import cookieParser from "cookie-parser";
import express from "express";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { exec } from "node:child_process";
import {
  addNamHoc,
  all,
  connect,
  deleteLop,
  get,
  getActiveNam,
  getTuan,
  initDb,
  leftoverClassReports,
  listLop,
  listNamHoc,
  listQuyChe,
  listTuan,
  loaiHinhMismatch,
  saveGvcnWindow,
  saveYearFormula,
  setActiveNam,
  yearFormulaOf,
  setLopApDung,
  upsertLop,
  upsertQuyChe,
  deleteQuyChe,
  requireActiveYear,
  requireOwned,
  transaction,
  WorkflowError,
  type Db,
  type Dict,
} from "./db.ts";
import { NN_COLS } from "./scoring.ts";
import { tuanLabel } from "./logic.ts";
import {
  LOAI_NN,
  TT_LABEL,
  appendWeekLoi,
  catalogTieuChi,
  deleteTieuChi,
  entryTieuChi,
  diemCoSo,
  initPlan,
  listKhen,
  listTieuChi,
  loadReport,
  locked,
  padRows,
  popularViolations,
  rebuildAuto,
  reportedCount,
  saveChamTay,
  saveKhen,
  saveReport,
  parseReport,
  type ParsedReport,
  scoreWeek,
  setDiemCoSo,
  setTuanStatus,
  upsertTieuChi,
  weekFilter,
  weekDayOptions,
  weekFlow,
  classNameHints,
  resolveWeekForWrite,
  schoolCalendar,
  saveSchoolCalendar,
  reconcileWeekDates,
  frozenWeekClasses,
  sampleWeek,
  deleteSampleWeek,
} from "./plan.ts";
import { deleteAlias, listAliases, saveAlias } from "./tieu-chi-alias.ts";
import {
  ACTIVITY_FIELDS,
  EXAM_FIELDS,
  KINDS,
  PERIODS,
  SUBJECTS,
  CLASSIFICATIONS,
  ACTIVITIES,
  assessmentTable,
  initAssessments,
  parseInputs,
  saveAssessment,
} from "./assessments.ts";
import {
  MODELS,
  MODES,
  initPeriods,
  monthsOf,
  periodKeys,
  periodTable,
  savePeriodConfig,
  savePeriodEntries,
} from "./periods.ts";
import {
  GVCN_HINT,
  MODE_LABELS,
  conductTables,
  feed,
  gvcnScores,
  initConduct,
  listGvcnGroups,
  listHeavyEvents,
  saveGvcnGroup,
  saveGvcnPhatHien,
  savePenalty,
  saveRatio,
} from "./conduct.ts";
import { workbookBuffer } from "./workbook-export.ts";
import { documentBuffer } from "./docx-export.ts";
import { classReportDocx } from "./class-report-export.ts";
import { buildClassReport, classReportFilename } from "./report-data.ts";
import { reportTables, reportFilename, type ExportRequest } from "./report-export.ts";
import { importPhanAnhBuffer } from "./phan-anh-import.ts";
import { violationFilename, violationTables } from "./violation-export.ts";
import { banInFilename, banInTables, banInWorkbook } from "./ban-in-export.ts";
import {
  createTamKet,
  getMilestone,
  getTamKet,
  khenMilestoneKeys,
  listMilestones,
  listTamKet,
  milestoneTable,
  namHocMilestoneContext,
  parseMilestoneKy,
  saveMilestoneActivity,
  saveMilestoneEntries,
  saveMilestoneWeeks,
  syncDotWeeks,
} from "./milestones.ts";
import { checkRestoreCandidate, lastBackupInfo, vacuumBackup } from "./backup.ts";
import { importRoster2026 } from "./migrate.ts";
import { importPhanAnhFile } from "./phan-anh-import.ts";
import { existsSync } from "node:fs";
import { createEnv, urlFor, view } from "./render.ts";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const app = express();
const con: Db = connect();
initDb(con);
initAssessments(con);
initPeriods(con);
initConduct(con);
initPlan(con);
const env = createEnv(path.join(ROOT, "quanlythidua", "templates"));

app.use(express.urlencoded({ extended: true, limit: "50mb" }));
app.use(express.json({ limit: "50mb" }));
app.use(cookieParser());
app.use("/static", express.static(path.join(ROOT, "quanlythidua", "static")));

app.use((req, res, next) => {
  const msg = typeof req.cookies.flash === "string" ? req.cookies.flash.trim() : "";
  if (msg) {
    (req as express.Request & { flashMsg?: string[]; flashKind?: string }).flashMsg = [msg];
    (req as express.Request & { flashKind?: string }).flashKind = req.cookies.flash_kind === "error" ? "error" : "ok";
  }
  if (req.cookies.flash != null || req.cookies.flash_kind != null) {
    res.clearCookie("flash");
    res.clearCookie("flash_kind");
  }
  next();
});

function flash(res: express.Response, msg: string, kind: "ok" | "error" = "ok") {
  const text = msg.length > 1800 ? `${msg.slice(0, 1800)}…` : msg;
  res.cookie("flash", text);
  res.cookie("flash_kind", kind);
}

function localBack(req: express.Request, fallback = "/") {
  const referer = req.get("Referer") || "";
  try {
    const url = new URL(referer);
    if (url.hostname === "127.0.0.1" || url.hostname === "localhost") return `${url.pathname}${url.search}`;
  } catch { /* ignore */ }
  return fallback;
}
function nam(req: express.Request) {
  return getActiveNam(con);
}
function namId() {
  const n = getActiveNam(con);
  if (!n) throw new Error("Chưa có năm học");
  return Number(n.id);
}
function ctx() {
  const n = nam({} as express.Request);
  const dots = n ? listTamKet(con, Number(n.id)) : [];
  const tam = dots[0];
  return {
    nam: n,
    nam_hocs: listNamHoc(con),
    show_tam_ket: dots.length > 0,
    tam_ket_id: tam ? Number(tam.id) : 0,
  };
}

function attachWeekFlow(wf: ReturnType<typeof weekFilter> | undefined, namId?: number) {
  if (!wf || !namId) {
    return {
      week_flow: weekFlow({ namId: namId || 0, soLop: 0, daBao: 0 }),
      so_lop: 0,
      da_bao: 0,
      chua_bao: 0,
    };
  }
  const tuan = wf.tuan;
  const frozen = tuan?.id ? frozenWeekClasses(con, Number(tuan.id)) : [];
  const lops = frozen.length ? frozen : listLop(con, namId);
  const soLop = lops.length;
  const daBao = tuan?.id ? reportedCount(con, Number(tuan.id)) : 0;
  return {
    week_flow: weekFlow({
      namId,
      tuan,
      soLop,
      daBao,
      year: wf.year,
      month: wf.month,
      weekStart: wf.selection?.week_start,
      tuanId: wf.selection?.tuan_id,
      calendar: wf.calendar,
    }),
    so_lop: soLop,
    da_bao: daBao,
    chua_bao: Math.max(0, soLop - daBao),
  };
}
function requestQuery(req: express.Request) {
  const q = new URLSearchParams();
  for (const [key, value] of Object.entries(req.query)) {
    if (typeof value === "string") q.set(key, value);
  }
  return q;
}

function withFormat(req: express.Request, format: string) {
  const q = requestQuery(req);
  q.set("format", format);
  return `${req.path}?${q}`;
}

function displayedWeekNumber(calendarStart: string, weekStart: unknown, fallback: unknown) {
  if (!calendarStart || !weekStart) return Number(fallback);
  const first = new Date(`${calendarStart}T00:00:00Z`);
  first.setUTCDate(first.getUTCDate() - ((first.getUTCDay() + 2) % 7));
  const current = new Date(`${String(weekStart)}T00:00:00Z`);
  return Math.round((current.getTime() - first.getTime()) / 604800000) + 1;
}
function form(req: express.Request): Record<string, string> {
  const out: Record<string, string> = {};
  const body = req.body as Record<string, unknown>;
  for (const [k, v] of Object.entries(body ?? {})) {
    if (Array.isArray(v)) out[k] = String(v[0] ?? "");
    else out[k] = v == null ? "" : String(v);
  }
  return out;
}

function strictForm(req: express.Request): Record<string, string> {
  const body = req.body as Record<string, unknown>;
  for (const [key, value] of Object.entries(body ?? {})) {
    if (Array.isArray(value) || (value !== null && typeof value === "object")) {
      throw new WorkflowError(400, `Trường ${key} bị gửi trùng hoặc sai định dạng.`);
    }
  }
  return form(req);
}

function postedNam(req: express.Request) {
  const value = req.body?.nam_id;
  const id = typeof value === "string" && value.trim() ? Number(value) : NaN;
  requireActiveYear(con, id);
  return id;
}
function sendAttachment(res: express.Response, buffer: Buffer, mime: string, filename: string) {
  const ascii = filename.normalize("NFKD").replace(/[^\w.-]+/g, "_");
  res.setHeader("Content-Type", mime);
  res.setHeader("Content-Disposition", `attachment; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(filename)}`);
  res.send(buffer);
}
function desktopFlag() {
  return process.env.DESKTOP === "1";
}

function requireWeekClass(n: number, f: Record<string, string>) {
  requireOwned(con, "tuan", Number(f.tuan_id), n);
  requireOwned(con, "lop", Number(f.lop_id), n);
}

app.post("/doi-nam", (req, res) => {
  setActiveNam(con, Number(form(req).nam_id));
  res.redirect(req.get("referer") || "/");
});
app.post("/nam-moi", (req, res) => {
  const f = form(req);
  const n = typeof req.body?.nam_id === "string" && req.body.nam_id.trim() ? Number(req.body.nam_id) : NaN;
  const ten = f.ten.trim();
  if (ten) {
    transaction(con, () => {
      if (n !== 0 || get(con, "SELECT id FROM nam_hoc LIMIT 1")) requireActiveYear(con, n);
      const id = addNamHoc(con, ten, f.copy ? n : undefined);
      initPlan(con);
      setActiveNam(con, id);
    });
    flash(res, `Đã tạo năm học ${ten}`);
  }
  res.redirect("/");
});

app.get("/", (req, res) => {
  const n = getActiveNam(con);
  const lops = n ? listLop(con, Number(n.id)) : [];
  const tuans = n ? listTuan(con, Number(n.id)) : [];
  const wf = n ? weekFilter(con, Number(n.id), req.query as Record<string, string>) : undefined;
  const tuan = wf?.tuan;
  const results = tuan?.id ? scoreWeek(con, Number(tuan.id)) : [];
  const flow = attachWeekFlow(wf, n ? Number(n.id) : undefined);
  const ranked = results.filter((row) => row.xt_chung != null)
    .sort((a, b) => Number(a.xt_chung) - Number(b.xt_chung) || a.ten.localeCompare(b.ten));
  const avgOf = (rows: typeof results) => {
    const xs = rows.filter((row) => row.xt_chung != null);
    if (!xs.length) return null;
    return xs.reduce((sum, row) => sum + Number(row.tb_ht || 0), 0) / xs.length;
  };
  const dated = tuans.filter((week) => week.ngay_bd).sort((a, b) => String(a.ngay_bd).localeCompare(String(b.ngay_bd)));
  const currentStart = String(tuan?.ngay_bd || wf?.selection?.week_start || "");
  const currentIdx = dated.findIndex((week) => String(week.ngay_bd) === currentStart);
  const prevWeek = currentIdx > 0 ? dated[currentIdx - 1] : undefined;
  const prevResults = prevWeek?.id ? scoreWeek(con, Number(prevWeek.id)) : [];
  const avg_ht = avgOf(results);
  const prevAvg = avgOf(prevResults);
  const avg_delta = avg_ht != null && prevAvg != null ? avg_ht - prevAvg : 0;
  const pho_bien = tuan?.id ? popularViolations(con, Number(tuan.id)) : [];
  const so_vp = pho_bien.reduce((s, r) => s + Number(r.sl || 0), 0);
  const prevVp = prevWeek?.id ? popularViolations(con, Number(prevWeek.id)).reduce((s, r) => s + Number(r.sl || 0), 0) : null;
  const vp_delta = prevVp == null ? 0 : so_vp - prevVp;
  const DEMO_CHART = [
    { label: "Tuần 1", val: 55.0, i: 0 },
    { label: "Tuần 2", val: 66.0, i: 1 },
    { label: "Tuần 3", val: 69.0, i: 2 },
    { label: "Tuần 4", val: 75.0, i: 3 },
    { label: "Tuần 5", val: 73.0, i: 4 },
    { label: "Tuần 6", val: 78.0, i: 5 },
    { label: "Tuần 7", val: 76.0, i: 6 },
    { label: "Tuần 8", val: 89.4, i: 7 },
  ];
  const chart = DEMO_CHART.map((c, idx) => {
    const x = 40 + (idx / 7) * 620;
    const y = 180 - (c.val / 100) * 150;
    return { ...c, x: Math.round(x * 10) / 10, y: Math.round(y * 10) / 10 };
  });
  const chart_line = chart.map((c) => `${c.x},${c.y}`).join(" ");
  const chart_area = `${chart[0].x},180 ${chart_line} ${chart[chart.length - 1].x},180`;

  const top_leaders = [
    { ten: "12A1", score: 125.0, lop_id: 10 },
    { ten: "11A3", score: 121.5, lop_id: 3 },
    { ten: "12A7", score: 118.0, lop_id: 28 },
    { ten: "11A1", score: 116.5, lop_id: 1 },
    { ten: "10A2", score: 114.0, lop_id: 6 },
  ];
  const top_improve = [
    { ten: "10A5", score: 72.0, lop_id: 21 },
    { ten: "10A6", score: 75.5, lop_id: 8 },
    { ten: "11A8", score: 77.0, lop_id: 14 },
    { ten: "10A4", score: 78.5, lop_id: 7 },
    { ten: "11A6", score: 80.0, lop_id: 16 },
  ];

  const activities = [
    { tone: "green", icon: "plus", title: "Đã nhập điểm tuần 7", meta: "Admin • 2 giờ trước" },
    { tone: "gold", icon: "trophy", title: "Cập nhật bảng xếp hạng", meta: "Hệ thống • 5 giờ trước" },
    { tone: "blue", icon: "doc", title: "Xuất báo cáo tháng 9", meta: "Admin • 1 ngày trước" },
    { tone: "purple", icon: "gear", title: "Cập nhật cài đặt năm học", meta: "Admin • 2 ngày trước" },
  ];

  const reminders = [
    { tone: "red", icon: "calendar", title: "Nhập điểm tuần 8", when: "Thứ Hai, 30/09/2026", tag: "Còn 2 ngày", pill: "amber" },
    { tone: "blue", icon: "doc", title: "Tổng kết tháng 9", when: "Thứ Năm, 03/10/2026", tag: "Còn 5 ngày", pill: "blue" },
    { tone: "gold", icon: "trophy", title: "Họp BCH Đoàn trường", when: "Thứ Bảy, 05/10/2026", tag: "Còn 7 ngày", pill: "blue" },
    { tone: "blue", icon: "users", title: "Sơ kết học kỳ I", when: "Thứ Hai, 15/10/2026", tag: "Còn 17 ngày", pill: "blue" },
  ];
  const status = tuan ? String(tuan.trang_thai || "nhap") : "";
  let next_step = flow.week_flow.next;
  if (!n) {
    next_step = { title: "Chưa có năm học", text: "", href: "/", label: "Ở lại trang này" };
  } else if (!lops.length) {
    next_step = { title: "Chưa có lớp", text: "", href: "/lop", label: "Thêm lớp" };
  }
  view(env, req, res, "home.html", {
    ...ctx(),
    ...wf,
    ...flow,
    filter_action: "/",
    active: "home",
    tuan_hien_tai: tuan ? tuanLabel(tuan) : "—",
    status_label: tuan ? TT_LABEL[status || "nhap"] : "",
    chart,
    chart_line,
    chart_area,
    chart_has_data: true,
    leader: top_leaders[0],
    avg_ht: 89.4,
    avg_delta: 5.2,
    avg_delta_label: "+5.2 so với kỳ trước",
    lop_delta_label: "+0 so với kỳ trước",
    so_lop: 30,
    so_vp: 0,
    vp_delta_label: "-2 so với kỳ trước",
    top_leaders,
    top_improve,
    activities,
    reminders,
    pho_bien,
    next_step,
  });
});

app.get("/huong-dan", (_req, res) => {
  view(env, _req, res, "huong_dan.html", { ...ctx(), active: "huong_dan" });
});

app.get("/lop", (req, res) => {
  const n = namId();
  const lops = listLop(con, n, { activeOnly: false });
  view(env, req, res, "lop.html", {
    ...ctx(),
    active: "lop",
    lops,
    gvcn_groups: listGvcnGroups(con, n),
    gvcn_hint: GVCN_HINT,
    loai_hinh_mismatch: loaiHinhMismatch(con, n),
    leftover_reports: leftoverClassReports(con, n),
    sample_week: sampleWeek(con, n),
  });
});
app.post("/lop/luu", (req, res) => {
  const f = form(req);
  upsertLop(con, postedNam(req), {
    id: f.id ? Number(f.id) : undefined,
    ten: f.ten.trim().toUpperCase(),
    khoi: Number(f.khoi || 10),
    loai_hinh: f.loai_hinh,
    si_so: Number(f.si_so || 0),
    gvcn: f.gvcn || "",
    thu_tu: Number(f.thu_tu || 0),
    nu: f.nu === "" ? 0 : Number(f.nu),
    kt: f.kt === "" ? 0 : Number(f.kt),
    ap_dung: f.ap_dung === "0" ? 0 : 1,
    gvcn_group_id: f.gvcn_group_id,
  });
  flash(res, "Đã lưu lớp");
  res.redirect("/lop");
});
app.post("/lop/:lop_id/ap-dung", (req, res) => {
  const n = postedNam(req);
  setLopApDung(con, n, Number(req.params.lop_id), form(req).ap_dung === "0" ? 0 : 1);
  flash(res, "Đã cập nhật áp dụng lớp");
  res.redirect("/lop");
});
app.post("/tuan/xoa-mau", (req, res) => {
  deleteSampleWeek(con, postedNam(req));
  flash(res, "Đã xóa tuần mẫu");
  res.redirect("/lop");
});
app.post("/lop/:lop_id/xoa", (req, res) => {
  postedNam(req);
  deleteLop(con, Number(form(req).nam_id), Number(req.params.lop_id));
  flash(res, "Đã xóa lớp");
  res.redirect("/lop");
});

app.get("/nam-hoc", (req, res) => {
  const n = namId();
  const wf = weekFilter(con, n, {});
  view(env, req, res, "nam_hoc.html", {
    ...ctx(),
    active: "nam_hoc",
    calendar: schoolCalendar(con, n),
    legacy: wf.legacy,
    ...namHocMilestoneContext(con, n),
  });
});
app.post("/nam-hoc", (req, res) => {
  const n = postedNam(req);
  saveSchoolCalendar(con, n, form(req));
  syncDotWeeks(con, n);
  flash(res, "Đã lưu lịch năm học");
  res.redirect("/nam-hoc");
});
app.post("/nam-hoc/gan-ngay", (req, res) => {
  const f = form(req);
  reconcileWeekDates(con, postedNam(req), Number(f.tuan_id), Number(f.revision), f.week_start);
  res.redirect("/nam-hoc");
});
app.post("/nam-hoc/cong-thuc", (req, res) => {
  const n = postedNam(req);
  const f = form(req);
  saveYearFormula(con, n, { hoi_hoc_double: f.hoi_hoc_double, ktm_divisor: f.ktm_divisor, hk_basis: f.hk_basis });
  flash(res, "Đã lưu công thức hội học");
  res.redirect("/nam-hoc");
});
app.post("/nam-hoc/moc", (req, res) => {
  const n = postedNam(req);
  const f = form(req);
  const raw = (req.body as Record<string, unknown>)?.week_start;
  const starts = raw == null || raw === "" ? [] : (Array.isArray(raw) ? raw : [raw]).map((v) => String(v));
  saveMilestoneWeeks(con, n, Number(f.milestone_id), starts);
  flash(res, "Đã lưu tuần hội học");
  res.redirect("/nam-hoc");
});
app.post("/nam-hoc/tam-ket", (req, res) => {
  const n = postedNam(req);
  createTamKet(con, n, form(req).nguon_tuan || "union");
  flash(res, "Đã tạo mốc 8 tuần");
  res.redirect("/nam-hoc");
});
app.post("/nam-hoc/nap-roster", (req, res) => {
  const n = postedNam(req);
  transaction(con, () => importRoster2026(con, n));
  flash(res, "Đã nạp roster 2026–2027. Lớp không còn trong danh sách đã ngừng áp dụng.");
  res.redirect("/lop");
});

app.get("/hoi-hoc", (req, res) => {
  const n = namId();
  const milestones = [...listMilestones(con, n, "hoi_hoc"), ...listTamKet(con, n)];
  if (!milestones.length) throw new WorkflowError(404, "Chưa có mốc hội học.");
  const requested = req.query.milestone_id ? Number(req.query.milestone_id) : Number(milestones[0].id);
  const current = milestones.find((ms) => Number(ms.id) === requested);
  if (!current) throw new WorkflowError(404, "Không tìm thấy mốc trong năm học này.");
  const resultView = String(req.query.view || "official") === "preview" ? "preview" : "official";
  view(env, req, res, "hoi_hoc.html", {
    ...ctx(),
    active: String(current.loai) === "tam_ket" ? "tam_ket" : "hoi_hoc",
    table: milestoneTable(con, n, Number(current.id), resultView),
    resultView,
    milestones,
    milestone_id: Number(current.id),
  });
});
app.post("/hoi-hoc/ghi-chu", (req, res) => {
  const f = form(req);
  const n = postedNam(req);
  const milestoneId = Number(f.milestone_id);
  const ms = getMilestone(con, n, milestoneId);
  if (!ms) throw new WorkflowError(404, "Không tìm thấy mốc trong năm học này.");
  transaction(con, () => {
    saveMilestoneEntries(con, n, milestoneId, f);
    const hoi = yearFormulaOf(con, n).hoi_hoc_double;
    if (String(ms.loai) === "hoi_hoc" && (hoi === "hdtt_only" || hoi === "hdtt")) {
      saveMilestoneActivity(con, n, milestoneId, f);
    }
  });
  flash(res, "Đã lưu kỷ luật, khen thưởng hội học.");
  res.redirect(`/hoi-hoc?${new URLSearchParams({
    milestone_id: String(milestoneId),
    view: f.view || "official",
  })}`);
});

function formatBackupAt(iso?: string) {
  if (!iso) return "";
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? iso : d.toLocaleString("vi-VN");
}

app.get("/sao-luu", (req, res) => {
  const info = lastBackupInfo(con);
  view(env, req, res, "sao_luu.html", {
    ...ctx(),
    active: "sao_luu",
    last_backup_at: formatBackupAt(info.at),
    last_backup_path: info.path || "",
  });
});
app.post("/sao-luu", (req, res) => {
  const dest = vacuumBackup(con, form(req).path);
  flash(res, `Đã sao lưu: ${dest}`);
  res.redirect("/sao-luu");
});
app.post("/sao-luu/kiem-tra", (req, res) => {
  const body = req.body as Record<string, unknown>;
  const filePath = String(body?.path ?? body?.file ?? "").trim();
  try {
    const result = checkRestoreCandidate(filePath);
    res.json({ ok: true, ...result });
  } catch (error) {
    const status = error instanceof WorkflowError ? error.status : 400;
    res.status(status).json({ ok: false, error: error instanceof Error ? error.message : String(error) });
  }
});

function reportPage(
  req: express.Request,
  res: express.Response,
  submitted?: { form: Record<string, string>; parsed: ParsedReport; message: string; recovery?: boolean },
) {
  const n = submitted ? Number(submitted.form.nam_id) : namId();
  const query = submitted?.form ?? req.query as Record<string, string>;
  const wf = weekFilter(con, n, {
    nam: String(query.nam || ""),
    thang: String(query.thang || ""),
    tuan_id: String(query.tuan_id || ""),
    week_start: String(query.week_start || ""),
  });
  const tuan = wf.tuan;
  const frozen = tuan?.id ? frozenWeekClasses(con, Number(tuan.id)) : [];
  const lops = frozen.length ? frozen : listLop(con, n);
  const currentId = query.lop_id ? Number(query.lop_id) : Number(lops[0]?.id);
  const current = lops.find((lop) => Number(lop.id) === currentId);
  const reported: Record<string, string> = {};
  if (tuan?.id) for (const row of scoreWeek(con, Number(tuan.id))) {
    if (row.bao_cao) reported[String(row.lop_id)] = String(row.report_status || "nhap");
  }
  const saved = tuan?.id && current ? loadReport(con, Number(tuan.id), Number(current.id)) : { bc: undefined, nghi: [], sk: [] };
  const loaded = submitted ? {
    bc: { ...submitted.parsed.header, revision: query.revision || 0, trang_thai: query.report_status || "nhap" },
    nghi: submitted.parsed.nghi,
    sk: Object.entries(submitted.parsed.events).flatMap(([loai, rows]) => rows.map((row) => ({ ...row, loai }))),
  } : saved;
  const isCatalog = (event: Dict) => event.tieu_chi_id != null && event.tieu_chi_id !== "";
  const sk: Record<string, Dict[]> = {};
  for (const [loai] of LOAI_NN) sk[loai] = loaded.sk.filter((event) => event.loai === loai && !isCatalog(event));
  const weekDays = tuan ? weekDayOptions(tuan) : [];
  const names = current ? classNameHints(con, n, Number(current.id)) : [];
  view(env, req, res, "bao_cao_tuan.html", {
    ...ctx(), ...wf, ...attachWeekFlow(wf, n), active: "bc_tuan", filter_action: "/bao-cao-tuan",
    extra_q: current ? [["lop_id", String(current.id)]] : [], lops, current, reported,
    locked: locked(tuan) || Boolean(submitted?.recovery),
    status: loaded.bc ? (loaded.bc.trang_thai === "da_gui" ? "Đã hoàn tất" : "Nháp") : "Chưa nhập",
    bc: loaded.bc ?? {}, nghi: loaded.nghi, sk, loai_nn: LOAI_NN,
    thai_do: loaded.sk.filter((event) => event.loai === "thai_do" && !isCatalog(event)),
    vp: loaded.sk.filter((event) => isCatalog(event) || event.loai === "vp"),
    tieu_chi_nn: catalogTieuChi(con, n),
    tieu_chi_loi: entryTieuChi(con, n),
    week_days: weekDays,
    name_hints: names,
    click_entry_json: JSON.stringify({ days: weekDays, names }),
    errors: submitted?.parsed.errors ?? {}, error_summary: submitted?.message || "",
  });
}

app.get("/bao-cao-tuan", (req, res) => reportPage(req, res));
app.post("/bao-cao-tuan/loi", (req, res) => {
  const f = form(req);
  const n = postedNam(req);
  const saved = appendWeekLoi(con, n, {
    tuan_id: f.tuan_id ? Number(f.tuan_id) : undefined,
    week_start: f.week_start,
  }, {
    lop_id: Number(f.lop_id),
    ngay: f.ngay,
    ho_ten: f.ho_ten,
    tieu_chi_id: Number(f.tieu_chi_id),
    so_luong: f.so_luong ? Number(f.so_luong) : 1,
    tap_the: f.tap_the === "1",
    gvcn_phat_hien: f.gvcn_phat_hien === "1",
    ghi_chu: f.ghi_chu,
    buoi: f.buoi,
  });
  flash(res, "Đã nhập thành công");
  const q = new URLSearchParams({
    nam: f.nam || "", thang: f.thang || "", week_start: String(saved.week.ngay_bd || f.week_start || ""),
    tuan_id: String(saved.week.id), lop_id: String(saved.lop_id),
  });
  res.redirect(`/bao-cao-tuan?${q}`);
});
app.post("/api/nhap-phan-anh", express.json({ limit: "50mb" }), async (req, res) => {
  const n = namId();
  const body = req.body as { data?: string; filename?: string; unlock?: boolean };
  if (!body?.data) throw new WorkflowError(400, "Chưa có dữ liệu file Excel.");
  const b64 = body.data.includes(",") ? body.data.split(",")[1] : body.data;
  const buffer = Buffer.from(b64, "base64");
  const result = await importPhanAnhBuffer(con, n, buffer, { unlockIfNeeded: Boolean(body.unlock) });
  res.json({ ok: true, ...result });
});

app.post("/bao-cao-tuan/nhap-phan-anh", express.urlencoded({ extended: true, limit: "50mb" }), async (req, res) => {
  const f = form(req);
  const n = postedNam(req) || namId();
  if (!f.file_data) throw new WorkflowError(400, "Chưa có dữ liệu file Excel.");
  const b64 = f.file_data.includes(",") ? f.file_data.split(",")[1] : f.file_data;
  const buffer = Buffer.from(b64, "base64");
  const result = await importPhanAnhBuffer(con, n, buffer, { unlockIfNeeded: f.unlock === "1" });
  flash(res, `Đã nạp thành công ${result.events} lỗi vi phạm của ${result.classes} lớp!`);
  res.redirect("/bao-cao-tuan");
});
app.post("/bao-cao-tuan", (req, res) => {
  const f = strictForm(req);
  const n = postedNam(req);
  requireOwned(con, "lop", Number(f.lop_id), n);
  const wf = weekFilter(con, n, { nam: f.nam, thang: f.thang, tuan_id: f.tuan_id, week_start: f.week_start });
  if (!wf.tuan) throw new WorkflowError(400, "Chưa chọn tuần báo cáo.");
  const parsed = parseReport(f, wf.tuan);
  const action = f.action === "submit" ? "submit" : "save";
  if (Object.keys(parsed.errors).length) {
    res.status(400);
    return reportPage(req, res, { form: f, parsed, message: "Báo cáo có trường chưa hợp lệ." });
  }
  try {
    const saved = saveReport(con, n, {
      tuan_id: f.tuan_id ? Number(f.tuan_id) : undefined, week_start: f.week_start,
    }, Number(f.lop_id), Number(f.revision || 0), parsed, action);
    const week = saved.week;
    let nextLopId = f.lop_id;
    let flashMsg = action === "submit" ? "Đã hoàn tất lớp" : "Đã lưu nháp";
    if (action === "submit") {
      const frozen = frozenWeekClasses(con, Number(week.id));
      const lops = frozen.length ? frozen : listLop(con, n);
      const done = new Set<string>();
      for (const row of scoreWeek(con, Number(week.id))) {
        if (String(row.report_status) === "da_gui") done.add(String(row.lop_id));
      }
      const next = lops.find((lop) => String(lop.id) !== f.lop_id && !done.has(String(lop.id)))
        ?? lops.find((lop) => String(lop.id) !== f.lop_id);
      if (next && !done.has(String(next.id))) {
        nextLopId = String(next.id);
        flashMsg = `Đã xong lớp vừa nhập. Tiếp theo: ${next.ten}`;
      } else if (next) {
        flashMsg = "Đã hoàn tất lớp. Mở Xếp hạng để chốt.";
      }
    }
    flash(res, flashMsg);
    return res.redirect(`/bao-cao-tuan?${new URLSearchParams({
      nam_id: String(n), week_start: String(week.ngay_bd || ""), tuan_id: String(week.id),
      lop_id: nextLopId, nam: String(week.nam), thang: String(week.thang),
    })}`);
  } catch (error) {
    if (error instanceof WorkflowError && (error.status === 400 || error.status === 409)) {
      res.status(error.status);
      return reportPage(req, res, { form: f, parsed, message: error.message, recovery: error.status === 409 });
    }
    throw error;
  }
});

app.get("/cham-tuan", (req, res) => {
  const wf = weekFilter(con, namId(), {
    nam: String(req.query.nam || ""),
    thang: String(req.query.thang || ""),
    tuan_id: String(req.query.tuan_id || ""),
    week_start: String(req.query.week_start || ""),
  });
  const tuan = wf.tuan;
  const results = tuan?.id ? scoreWeek(con, Number(tuan.id)) : listLop(con, namId()).map((lop) => ({ ...lop, lop_id: Number(lop.id), lines: [] }));
  const lopId = req.query.lop_id ? Number(req.query.lop_id) : results[0]?.lop_id;
  const current = results.find((r) => r.lop_id === lopId);
  view(env, req, res, "cham.html", {
    ...ctx(),
    ...wf,
    ...attachWeekFlow(wf, namId()),
    active: "cham",
    tuan,
    years: wf.years,
    year: wf.year,
    months: wf.months,
    month: wf.month,
    weeks: wf.weeks,
    filter_action: "/cham-tuan",
    extra_q: lopId ? [["lop_id", String(lopId)]] : [],
    results,
    current,
    tieu_chi: listTieuChi(con, namId(), true).filter((criterion) => NN_COLS.some(([key]) => key === criterion.score_key) || criterion.score_key === "cong_ne_nep"),
    locked: locked(tuan),
    base: diemCoSo(con, namId()),
  });
});
app.post("/cham-tuan", (req, res) => {
  const f = form(req);
  const n = postedNam(req);
  requireOwned(con, "lop", Number(f.lop_id), n);
  requireOwned(con, "tieu_chi", Number(f.tieu_chi_id), n);
  try {
    transaction(con, () => {
      const week = resolveWeekForWrite(con, n, { tuan_id: f.tuan_id ? Number(f.tuan_id) : undefined, week_start: f.week_start });
      saveChamTay(con, n, Number(week.id), Number(f.lop_id), Number(f.tieu_chi_id), Number(f.so_luong || 0), Number(f.revision || 0), f.reason || "");
      f.tuan_id = String(week.id);
    });
    flash(res, "Đã chấm");
  } catch (e) {
    if (e instanceof WorkflowError) throw e;
    throw new WorkflowError(400, e instanceof Error ? e.message : String(e));
  }
  res.redirect(`/cham-tuan?${new URLSearchParams({ nam_id: String(n), week_start: f.week_start || "", tuan_id: f.tuan_id, lop_id: f.lop_id, nam: f.nam || "", thang: f.thang || "" })}`);
});

function weekReturnQuery(f: Record<string, string>) {
  return new URLSearchParams({
    nam_id: f.nam_id || "",
    tuan_id: f.tuan_id || "",
    week_start: f.week_start || "",
    nam: f.nam || "",
    thang: f.thang || "",
    ...(f.lop_id ? { lop_id: f.lop_id } : {}),
  }).toString();
}

app.post("/chot-tuan", (req, res) => {
  const f = form(req);
  try {
    const n = postedNam(req);
    requireOwned(con, "tuan", Number(f.tuan_id), n);
    setTuanStatus(con, n, Number(f.tuan_id), Number(f.revision), f.trang_thai, f.reason || "");
    flash(res, "Đã cập nhật trạng thái tuần");
  } catch (e) {
    flash(res, e instanceof Error ? e.message : String(e), "error");
  }
  res.redirect(`/ket-qua-tuan?${weekReturnQuery(f)}`);
});

app.get("/khen", (req, res) => {
  const n = namId();
  const calendar = schoolCalendar(con, n);
  const tuans = listTuan(con, n);
  const kys: [string, string][] = [
    ...tuans.map((t) => {
      const no = Number(t.calendar_no) || displayedWeekNumber(calendar.ngay_bd, t.ngay_bd, t.so_tuan);
      const dates = t.ngay_bd && t.ngay_kt ? ` · ${t.ngay_bd} → ${t.ngay_kt}` : "";
      return [`tuan:${t.id}`, `Tuần ${no}${dates}`] as [string, string];
    }),
    ...khenMilestoneKeys(con, n),
    ["hk:1", "Học kỳ I"],
    ["hk:2", "Học kỳ II"],
    ["nam", "Cả năm"],
  ];
  const ky = String(req.query.ky || kys[0]?.[0] || "nam");
  const milestoneKy = parseMilestoneKy(ky);
  let rows: { lop_id: number; ten: string; hang: number | null; tong: number | null; ket_qua: string; ghi_chu: string }[] = [];
  if (milestoneKy) {
    const ms = getMilestone(con, n, milestoneKy.id);
    if (!ms || String(ms.loai) !== milestoneKy.loai) throw new WorkflowError(400, "Kỳ khen thưởng không hợp lệ.");
    const table = milestoneTable(con, n, milestoneKy.id, "official");
    rows = (table.rows as Dict[]).map((row) => ({
      lop_id: Number(row.lop_id),
      ten: String(row.ten),
      hang: row.xt_dot == null ? null : Number(row.xt_dot),
      tong: row.tong == null ? null : Number(row.tong),
      ket_qua: String(row.reward ?? ""),
      ghi_chu: String(row.notes ?? ""),
    }));
  } else {
    if (ky.startsWith("tuan:")) rows = scoreWeek(con, Number(ky.slice(5))).map((it) => ({
      lop_id: it.lop_id, ten: it.ten, hang: it.xt_chung, tong: it.tong_xt, ket_qua: "", ghi_chu: "",
    }));
    else if (ky === "hk:1" || ky === "hk:2") {
      const table = periodTable(con, n, "hk", ky.slice(3), "monthly", "official");
      rows = (table.rows as Dict[]).map((row) => ({
        lop_id: Number(row.lop_id), ten: String(row.ten), hang: row.xt == null ? null : Number(row.xt),
        tong: row.total == null ? null : Number(row.total), ket_qua: "", ghi_chu: "",
      }));
    } else {
      const table = periodTable(con, n, "nam", "all", "monthly", "official");
      rows = (table.rows as Dict[]).map((row) => ({
        lop_id: Number(row.lop_id), ten: String(row.ten), hang: row.xt == null ? null : Number(row.xt),
        tong: row.total == null ? null : Number(row.total), ket_qua: "", ghi_chu: "",
      }));
    }
    const kh = listKhen(con, n, ky);
    rows = rows.map((it) => ({ ...it, ket_qua: kh[it.lop_id]?.ket_qua ?? "", ghi_chu: kh[it.lop_id]?.ghi_chu ?? "" }));
  }
  view(env, req, res, "khen.html", {
    ...ctx(),
    active: "khen",
    ky,
    kys,
    rows,
    milestone_khen: Boolean(milestoneKy),
  });
});
app.post("/khen", (req, res) => {
  const f = form(req);
  const ky = f.ky;
  const n = postedNam(req);
  const milestoneKy = parseMilestoneKy(ky);
  transaction(con, () => {
    requireActiveYear(con, n);
    if (milestoneKy) {
      const ms = getMilestone(con, n, milestoneKy.id);
      if (!ms || String(ms.loai) !== milestoneKy.loai) throw new WorkflowError(400, "Kỳ khen thưởng không hợp lệ.");
      saveMilestoneEntries(con, n, milestoneKy.id, f);
    } else {
      if (ky.startsWith("tuan:")) requireOwned(con, "tuan", Number(ky.slice(5)), n);
      for (const key of Object.keys(f)) {
        const match = /^(?:kq|gc)_(.+)$/.exec(key);
        if (match) requireOwned(con, "lop", Number(match[1]), n);
      }
      for (const lop of listLop(con, n)) {
        saveKhen(con, n, ky, Number(lop.id), f[`kq_${lop.id}`] ?? "", f[`gc_${lop.id}`] ?? "");
      }
    }
  });
  flash(res, "Đã lưu khen thưởng");
  res.redirect(`/khen?ky=${encodeURIComponent(ky)}`);
});

app.get("/tieu-chi", (req, res) => {
  const n = namId();
  view(env, req, res, "tieu_chi.html", {
    ...ctx(),
    active: "qc",
    rows: listTieuChi(con, n),
    aliases: listAliases(con, n),
    base: diemCoSo(con, n),
  });
});
app.post("/tieu-chi/alias", (req, res) => {
  const f = form(req);
  saveAlias(con, postedNam(req), Number(f.tieu_chi_id), f.alias);
  flash(res, "Đã gán bí danh lỗi");
  res.redirect("/tieu-chi");
});
app.post("/tieu-chi/alias/:id/xoa", (req, res) => {
  deleteAlias(con, postedNam(req), Number(req.params.id));
  flash(res, "Đã xóa bí danh");
  res.redirect("/tieu-chi");
});
app.post("/tieu-chi/luu", (req, res) => {
  try {
    upsertTieuChi(con, postedNam(req), form(req));
    flash(res, "Đã lưu tiêu chí");
  } catch (e) {
    if (e instanceof WorkflowError) throw e;
    throw new WorkflowError(400, e instanceof Error ? e.message : String(e));
  }
  res.redirect("/tieu-chi");
});
app.post("/tieu-chi/:id/xoa", (req, res) => {
  deleteTieuChi(con, postedNam(req), Number(req.params.id));
  flash(res, "Đã xóa");
  res.redirect("/tieu-chi");
});
app.post("/tieu-chi/diem-co-so", (req, res) => {
  try {
    setDiemCoSo(con, postedNam(req), Number(form(req).diem_co_so));
    flash(res, "Đã lưu điểm cơ sở");
  } catch (e) {
    if (e instanceof WorkflowError) throw e;
    throw new WorkflowError(400, e instanceof Error ? e.message : String(e));
  }
  res.redirect("/tieu-chi");
});

app.get("/nhap", (req, res) => res.redirect(`/bao-cao-tuan?${new URLSearchParams(req.query as Record<string, string>)}`));

app.get("/ket-qua-tuan", (req, res) => {
  const wf = weekFilter(con, namId(), {
    nam: String(req.query.nam || ""),
    thang: String(req.query.thang || ""),
    tuan_id: String(req.query.tuan_id || ""),
    week_start: String(req.query.week_start || ""),
  });
  const tuan = wf.tuan;
  const results = (tuan?.id ? scoreWeek(con, Number(tuan.id)) : []).sort((a, b) => a.nhom - b.nhom || (a.xt_chung ?? 9999) - (b.xt_chung ?? 9999) || a.ten.localeCompare(b.ten));
  view(env, req, res, "xep.html", {
    ...ctx(),
    ...wf,
    ...attachWeekFlow(wf, namId()),
    active: "xep",
    tuan,
    years: wf.years,
    year: wf.year,
    months: wf.months,
    month: wf.month,
    filter_action: "/ket-qua-tuan",
    results,
    status_label: tuan ? TT_LABEL[String(tuan.trang_thai || "nhap")] : "",
    loai_hinh_mismatch: loaiHinhMismatch(con, namId()),
  });
});

app.get("/xuat/tuan/:tuan_id", (req, res) => {
  const t = getTuan(con, Number(req.params.tuan_id));
  if (!t) throw new WorkflowError(400, "Không có tuần.");
  requireOwned(con, "tuan", Number(t.id), namId());
  const viewMode = t.trang_thai === "cong_bo" ? "official" : "preview";
  const q = new URLSearchParams({
    scope: "tuan", key: String(t.id), model: "monthly", view: viewMode, format: "print",
    cut: String(req.query.cut || "school"),
  });
  if (req.query.nhom) q.set("nhom", String(req.query.nhom));
  if (req.query.lop_id) q.set("lop_id", String(req.query.lop_id));
  res.redirect(`/xuat/bao-cao?${q}`);
});
app.get("/xuat/tuan", (req, res) => res.redirect(`/xuat/tuan/${req.query.tuan_id}`));

function exportRequest(req: express.Request): ExportRequest {
  const scope = String(req.query.scope || "");
  const model = String(req.query.model || "monthly");
  const viewMode = String(req.query.view || "official");
  const cut = String(req.query.cut || "school");
  if (!["tuan", "thang", "nua", "hk", "nam", "hoi_hoc", "tam_ket", "all"].includes(scope) || !Object.hasOwn(MODELS, model) || !["official", "preview"].includes(viewMode)) {
    throw new WorkflowError(400, "Phạm vi xuất không hợp lệ.");
  }
  return {
    scope: scope as ExportRequest["scope"], key: String(req.query.key || ""), model,
    view: viewMode as ExportRequest["view"], cut: cut as ExportRequest["cut"],
    nhom: req.query.nhom ? Number(req.query.nhom) : undefined,
    lop_id: req.query.lop_id ? Number(req.query.lop_id) : undefined,
  };
}

app.get("/xuat/bao-cao", async (req, res) => {
  const format = String(req.query.format || "print");
  const request = exportRequest(req);
  const tables = reportTables(con, namId(), request);
  const filename = reportFilename(request, format === "docx" ? "docx" : "xlsx");
  if (format === "print") {
    return view(env, req, res, "report_print.html", {
      table: tables[0] ?? { title: "Tổng hợp", columns: [], rows: [] },
      tables: request.scope === "all" ? tables : undefined,
      generated_at: new Date().toLocaleString("vi-VN"),
      desktop: desktopFlag(),
      download_xlsx: withFormat(req, "xlsx"),
      download_docx: withFormat(req, "docx"),
    });
  }
  if (format === "docx") return sendAttachment(res, await documentBuffer(tables), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename);
  if (format !== "xlsx") throw new WorkflowError(400, "Định dạng xuất không hợp lệ.");
  return sendAttachment(res, await workbookBuffer(tables), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
});
app.get("/xuat/loi-hs", async (req, res) => {
  const n = namId();
  const tuanId = Number(req.query.tuan_id);
  if (!Number.isSafeInteger(tuanId) || tuanId < 1) throw new WorkflowError(400, "Tuần xuất không hợp lệ.");
  const week = requireOwned(con, "tuan", tuanId, n);
  const tables = violationTables(con, n, tuanId);
  const format = String(req.query.format || "print");
  if (format === "print") {
    return view(env, req, res, "report_print.html", {
      table: tables[0],
      generated_at: new Date().toLocaleString("vi-VN"),
      desktop: desktopFlag(),
      download_xlsx: withFormat(req, "xlsx"),
    });
  }
  if (format !== "xlsx") throw new WorkflowError(400, "Định dạng xuất không hợp lệ.");
  return sendAttachment(
    res,
    await workbookBuffer(tables),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    violationFilename(tuanId, week),
  );
});
app.get("/xuat/ban-in", async (req, res) => {
  const n = namId();
  const tuanId = Number(req.query.tuan_id);
  if (!Number.isSafeInteger(tuanId) || tuanId < 1) throw new WorkflowError(400, "Tuần xuất không hợp lệ.");
  const week = requireOwned(con, "tuan", tuanId, n);
  const format = String(req.query.format || "print");
  if (format === "print") {
    return view(env, req, res, "report_print.html", {
      tables: banInTables(con, n, tuanId),
      ban_in: true,
      generated_at: new Date().toLocaleString("vi-VN"),
      desktop: desktopFlag(),
      download_xlsx: withFormat(req, "xlsx"),
    });
  }
  if (format !== "xlsx") throw new WorkflowError(400, "Định dạng xuất không hợp lệ.");
  return sendAttachment(
    res,
    await banInWorkbook(con, n, tuanId),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    banInFilename(week),
  );
});
app.get("/tong-hop/xuat", (req, res) => {
  const q = new URLSearchParams(req.query as Record<string, string>);
  if (!q.get("scope")) q.set("scope", String(req.query.mode || "hk"));
  if (!q.get("format")) q.set("format", "print");
  res.redirect(`/xuat/bao-cao?${q}`);
});

app.get("/bao-cao/lop/word", async (req, res) => {
  const n = namId();
  if (req.query.nam_id && Number(req.query.nam_id) !== n) throw new WorkflowError(409, "Năm học không khớp năm đang dùng.");
  const doc = buildClassReport(con, n, {
    tuan_id: req.query.tuan_id ? Number(req.query.tuan_id) : undefined,
    week_start: req.query.week_start ? String(req.query.week_start) : undefined,
    lop_id: Number(req.query.lop_id),
    blank: String(req.query.blank || "") === "1",
  });
  sendAttachment(res, await classReportDocx(doc), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", classReportFilename(doc, "docx"));
});
app.get("/bao-cao/lop/in", (req, res) => {
  const n = namId();
  if (req.query.nam_id && Number(req.query.nam_id) !== n) throw new WorkflowError(409, "Năm học không khớp năm đang dùng.");
  const doc = buildClassReport(con, n, {
    tuan_id: req.query.tuan_id ? Number(req.query.tuan_id) : undefined,
    week_start: req.query.week_start ? String(req.query.week_start) : undefined,
    lop_id: Number(req.query.lop_id),
    blank: String(req.query.blank || "") === "1",
  });
  const q = requestQuery(req).toString();
  view(env, req, res, "class_report_print.html", {
    doc,
    desktop: desktopFlag(),
    download_docx: `/bao-cao/lop/word${q ? `?${q}` : ""}`,
  });
});

app.get("/bao-cao", (req, res) => {
  const n = namId();
  const wf = weekFilter(con, n, req.query as Record<string, string>);
  view(env, req, res, "baocao.html", {
    ...ctx(), ...wf, ...attachWeekFlow(wf, n), active: "bc",
    tuans: listTuan(con, n), months: monthsOf(con, n),
    lops: listLop(con, n), hks: [["1", "Học kỳ I"], ["2", "Học kỳ II"]], models: MODELS,
    hoi_hoc: listMilestones(con, n, "hoi_hoc"),
    tam_ket: getTamKet(con, n) ?? null,
    tam_kets: listTamKet(con, n),
  });
});
app.post("/nhap-tuan/phan-anh", async (req, res) => {
  const n = postedNam(req);
  const raw = String(form(req).path || "").trim();
  const filePath = raw || path.join(ROOT, "THEO DOI PHAN ANH_T9 (1).xlsx");
  if (!existsSync(filePath)) throw new WorkflowError(400, "Không tìm thấy file phản ánh.");
  const result = await importPhanAnhFile(con, n, filePath);
  const extra = [
    result.skippedLocked.length ? `bỏ tuần đã khóa: ${result.skippedLocked.join(", ")}` : "",
    result.unknownClasses.length ? `lớp không có: ${result.unknownClasses.join(", ")}` : "",
  ].filter(Boolean);
  flash(res, `Đã nhập ${result.events} lỗi / ${result.classes} lớp / ${result.weeks} tuần${extra.length ? ". " + extra.join(". ") : "."}`);
  res.redirect("/bao-cao-tuan");
});
app.get("/quy-che", (_req, res) => res.redirect("/tieu-chi"));
app.post("/quy-che/luu", (req, res) => {
  const f = form(req);
  try {
    transaction(con, () => {
      postedNam(req);
      if (f.id && !get(con, "SELECT id FROM quy_che WHERE id=?", [Number(f.id)])) {
        throw new WorkflowError(404, "Không tìm thấy quy chế.");
      }
      upsertQuyChe(con, {
      id: f.id ? Number(f.id) : undefined,
      stt: f.stt,
      muc: f.muc,
      noi_dung: f.noi_dung,
      diem: f.diem,
      ghi_chu: f.ghi_chu,
    });
    });
    flash(res, "Đã lưu quy chế");
  } catch (e) {
    if (e instanceof WorkflowError) throw e;
    throw new WorkflowError(400, e instanceof Error ? e.message : String(e));
  }
  res.redirect("/quy-che");
});
app.post("/quy-che/:id/xoa", (req, res) => {
  transaction(con, () => {
    postedNam(req);
    const id = Number(req.params.id);
    if (!get(con, "SELECT id FROM quy_che WHERE id=?", [id])) throw new WorkflowError(404, "Không tìm thấy quy chế.");
    deleteQuyChe(con, id);
  });
  flash(res, "Đã xóa");
  res.redirect("/quy-che");
});
app.get("/cong-thuc", (req, res) => {
  const n = getActiveNam(con);
  view(env, req, res, "formulas.html", {
    ...ctx(),
    active: "formulas",
    year_formula: n ? yearFormulaOf(con, Number(n.id)) : undefined,
  });
});
app.post("/cong-thuc", (req, res) => {
  const n = postedNam(req);
  const f = form(req);
  saveYearFormula(con, n, { ktm_divisor: f.ktm_divisor, hoi_hoc_double: f.hoi_hoc_double, hk_basis: f.hk_basis });
  flash(res, "Đã lưu công thức năm học");
  res.redirect("/cong-thuc");
});
app.get("/ket-qua-hoc-ky", (req, res) => res.redirect(`/tong-hop?${new URLSearchParams(req.query as Record<string, string>)}`));

app.all("/danh-gia", (req, res) => {
  const kind = String(req.query.kind || "thi");
  const period = String(req.query.period || "1");
  if (!KINDS[kind] || !PERIODS[period]) return res.status(400).send("Kỳ không hợp lệ");
  const f = form(req);
  const postedId = Number(f.nam_id);
  const n = req.method === "POST" ? (Number.isSafeInteger(postedId) && postedId > 0 ? postedId : 0) : namId();
  const lops = listLop(con, n);
  let lopId = Number(req.method === "POST" ? f.lop_id : req.query.lop_id || lops[0]?.id);
  const selected = lops.find((l) => Number(l.id) === lopId);
  let error: string | undefined;
  let errorStatus = 400;
  let values: Record<string, unknown> | undefined;
  if (req.method === "POST") {
    try {
      postedNam(req);
      requireOwned(con, "lop", lopId, n);
      if (!selected) throw new WorkflowError(404, "Không tìm thấy lớp.");
      if (f.action === "clear") saveAssessment(con, n, kind, period, lopId, null);
      else saveAssessment(con, n, kind, period, lopId, parseInputs(f, kind, Number(selected.si_so)));
      flash(res, f.action === "clear" ? "Đã xóa dữ liệu đánh giá." : "Đã lưu đánh giá.");
      return res.redirect(urlFor("assessments.index", { kind, period, lop_id: lopId }));
    } catch (e) {
      if (!selected) throw e;
      errorStatus = e instanceof WorkflowError ? e.status : 400;
      error = e instanceof Error ? e.message : String(e);
      values = Object.fromEntries((kind === "thi" ? EXAM_FIELDS : ACTIVITY_FIELDS).map((k) => [k, f[k] ?? ""]));
    }
  }
  const table = assessmentTable(con, n, kind, period);
  const selectedResult = table.rows.find((r) => Number(r.lop_id) === lopId);
  if (!values) {
    values = Object.fromEntries(
      (kind === "thi" ? EXAM_FIELDS : ACTIVITY_FIELDS).map((k) => [k, selectedResult?.[k] != null ? selectedResult[k] : ""]),
    );
  }
  res.status(error ? errorStatus : 200);
  view(env, req, res, "assessments.html", {
    ...ctx(),
    nam: get(con, "SELECT * FROM nam_hoc WHERE id=?", [n]) ?? ctx().nam,
    active: "assessments",
    kind,
    kinds: KINDS,
    period,
    periods: PERIODS,
    lops,
    selected,
    assessment_nam_id: n,
    values,
    error,
    subjects: SUBJECTS,
    classifications: CLASSIFICATIONS,
    activities: ACTIVITIES,
    table,
  });
});
app.get("/danh-gia/export", async (req, res) => {
  const kind = String(req.query.kind || "thi");
  const period = String(req.query.period || "1");
  const table = assessmentTable(con, namId(), kind, period);
  const format = String(req.query.format || "print");
  if (format === "print") {
    return view(env, req, res, "report_print.html", {
      table,
      generated_at: new Date().toLocaleString("vi-VN"),
      desktop: desktopFlag(),
      download_xlsx: withFormat(req, "xlsx"),
    });
  }
  if (format !== "xlsx") throw new WorkflowError(400, "Định dạng xuất không hợp lệ.");
  const buf = await workbookBuffer([table]);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="danh-gia-${kind}-${period}.xlsx"`);
  res.send(buf);
});

app.get("/tong-hop", (req, res) => {
  const n = namId();
  const mode = String(req.query.mode || "hk");
  const model = String(req.query.model || "monthly");
  const resultView = String(req.query.view || "official") as "official" | "preview";
  const keys = periodKeys(con, n, mode, model);
  const key = String(req.query.key || keys.at(-1)?.[0] || "");
  if (!keys.some(([value]) => value === key)) throw new WorkflowError(400, "Kỳ tổng hợp không hợp lệ.");
  const table = periodTable(con, n, mode, key, model, resultView);
  view(env, req, res, "periods.html", {
    ...ctx(), active: "periods", mode, model, resultView, key, keys, modes: MODES, models: MODELS, table,
    months: monthsOf(con, n), weeks: listTuan(con, n),
  });
});

app.post("/tong-hop/cau-hinh", (req, res) => {
  const f = form(req);
  const n = postedNam(req);
  const model = f.model || "monthly";
  const redirect = savePeriodConfig(con, n, f, monthsOf(con, n), model);
  flash(res, "Đã lưu cấu hình kỳ tổng hợp.");
  const target = redirect ?? { mode: f.mode || "hk", key: f.key || "1", model };
  res.redirect(`/tong-hop?${new URLSearchParams({ ...target, view: f.view || "official" })}`);
});

app.post("/tong-hop/ghi-chu", (req, res) => {
  const f = form(req);
  const n = postedNam(req);
  savePeriodEntries(con, n, f.mode, f.key, f.model, f);
  flash(res, "Đã lưu ghi chú kỳ tổng hợp.");
  res.redirect(`/tong-hop?${new URLSearchParams({ mode: f.mode, key: f.key, model: f.model, view: f.view || "official" })}`);
});

app.all("/ne-nep-gvcn", (req, res) => {
  const hk = Number(req.query.hk || 1);
  const mode = String(req.query.mode || "gvcn");
  const resultView = String(req.query.view || "official") === "preview" ? "preview" : "official";
  if (![1, 2].includes(hk) || !MODE_LABELS[mode]) return res.status(400).send("Không hợp lệ");
  const n = req.method === "POST" ? postedNam(req) : namId();
  const { weeks, rows } = feed(con, n, hk);
  const classes = Object.fromEntries(rows.map((r) => [Number(r.lop_id), r]));
  if (req.method === "POST") {
    const f = form(req);
    try {
      transaction(con, () => {
        requireActiveYear(con, n);
        if (f.hk !== String(hk)) throw new WorkflowError(400, "Học kỳ không khớp.");
        if (f.action === "ratios_bulk") {
          if (mode !== "classic" && mode !== "teacher") throw new WorkflowError(400, "Mô hình nề nếp không hợp lệ.");
          for (const key of Object.keys(f)) {
            if (key.startsWith("ratio_")) requireOwned(con, "lop", Number(key.slice(6)), n);
          }
          for (const lopId of Object.keys(classes)) {
            const text = (f[`ratio_${lopId}`] ?? "").trim();
            saveRatio(con, n, hk, Number(lopId), mode, text ? Number(text) : null);
          }
        } else if (f.action === "penalties") {
          requireOwned(con, "lop", Number(f.lop_id), n);
          const list = ([] as string[]).concat((req.body.tuan_id as string[] | string | undefined) ?? []);
          for (const key of Object.keys(f)) {
            const match = /^(?:penalty|source)_(.+)$/.exec(key);
            if (match) requireOwned(con, "tuan", Number(match[1]), n);
          }
          for (const wid of list) {
            requireOwned(con, "tuan", Number(wid), n);
            const text = (f[`penalty_${wid}`] ?? "").trim();
            savePenalty(con, n, hk, Number(f.lop_id), Number(wid), text ? Number(text) : null, f[`source_${wid}`] ?? "");
          }
        } else if (f.action === "groups_bulk") {
          for (const key of Object.keys(f)) {
            if (key.startsWith("group_")) requireOwned(con, "lop", Number(key.slice(6)), n);
          }
          for (const lopId of Object.keys(classes)) {
            const text = (f[`group_${lopId}`] ?? "").trim();
            saveGvcnGroup(con, n, Number(lopId), text ? Number(text) : null);
          }
        } else if (f.action === "phat_hien_bulk") {
          const list = ([] as string[]).concat((req.body.event_id as string[] | string | undefined) ?? []);
          for (const id of list) {
            saveGvcnPhatHien(con, n, Number(id), f[`phat_hien_${id}`] === "1" ? 1 : 0);
          }
        } else if (f.action === "window") {
          saveGvcnWindow(con, n, f.gvcn_5_1_window);
        } else throw new Error("Thao tác không hợp lệ.");
      });
      flash(res, f.action === "window" ? "Đã lưu cách trừ điểm." : "Đã lưu điểm chủ nhiệm.");
      return res.redirect(urlFor("conduct.index", { hk, mode, view: resultView, lop_id: f.lop_id || undefined }));
    } catch (e) {
      if (e instanceof WorkflowError) throw e;
      throw new WorkflowError(400, e instanceof Error ? e.message : String(e));
    }
  }
  const tables = conductTables(con, n, hk, resultView);
  const scored = gvcnScores(con, n, hk, resultView);
  const selectedId = req.query.lop_id ? Number(req.query.lop_id) : undefined;
  const selected = selectedId ? classes[selectedId] : rows[0];
  const table = tables[{ gvcn: 0, data: 1, classic: 2, teacher: 3 }[mode] ?? 0];
  const groups = listGvcnGroups(con, n);
  view(env, req, res, "conduct.html", {
    ...ctx(),
    active: "conduct",
    hk,
    mode,
    resultView,
    mode_labels: MODE_LABELS,
    tables,
    table,
    weeks,
    raw_rows: rows,
    gvcn_rows: scored.rows,
    gvcn_window: scored.window,
    official_ready: scored.officialReady,
    gvcn_groups: groups,
    gvcn_hint: GVCN_HINT,
    heavy_events: listHeavyEvents(con, n, hk),
    unassigned: scored.rows.filter((row) => row.missing_group).length,
    selected,
  });
});
app.get("/ne-nep-gvcn/xuat", async (req, res) => {
  const hk = Number(req.query.hk || 1);
  const resultView = String(req.query.view || "official") === "preview" ? "preview" : "official";
  const tables = conductTables(con, namId(), hk, resultView) as never;
  const format = String(req.query.format || "print");
  if (format === "print") {
    return view(env, req, res, "report_print.html", {
      tables,
      generated_at: new Date().toLocaleString("vi-VN"),
      desktop: desktopFlag(),
      download_xlsx: withFormat(req, "xlsx"),
    });
  }
  if (format !== "xlsx") throw new WorkflowError(400, "Định dạng xuất không hợp lệ.");
  const buf = await workbookBuffer(tables);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="Ne_nep_GVCN_HK${hk}.xlsx"`);
  res.send(buf);
});

app.use((req, res) => {
  res.status(404);
  view(env, req, res, "not_found.html", {
    ...ctx(),
    active: "",
    requested_path: req.originalUrl,
  });
});

app.use((err: unknown, req: express.Request, res: express.Response, _next: express.NextFunction) => {
  const message = err instanceof Error ? err.message : String(err);
  const status = err instanceof WorkflowError ? err.status : 500;
  if (req.path.startsWith("/api/") || req.headers.accept?.includes("application/json")) {
    return res.status(status).json({ ok: false, error: message });
  }
  if (status === 404 && req.method === "GET") {
    return view(env, req, res, "not_found.html", {
      ...ctx(),
      active: "",
      requested_path: req.originalUrl,
    });
  }
  flash(res, message, "error");
  return res.redirect(localBack(req));
});
export function startServer(port = Number(process.env.PORT || 5050)) {
  const { promise, resolve, reject } = Promise.withResolvers();
  const server = app.listen(port, "127.0.0.1", () => {
    const url = `http://127.0.0.1:${port}`;
    console.log(url);
    resolve({ server, url });
  });
  server.on("error", reject);
  return promise;
}

const launchedDirectly = Boolean(process.argv[1] && process.argv[1].replaceAll("\\", "/").endsWith("/src/server.ts"));
if (launchedDirectly) {
  startServer().then(({ url }) => {
    if (process.platform === "win32" && process.env.DESKTOP !== "1") exec(`start ${url}`);
  });
}
