import ExcelJS from "exceljs";
import { get, listLop, transaction, WorkflowError, type Db } from "./db.ts";
import { matchLoi } from "./tieu-chi-alias.ts";
import {
  fridayOf,
  loadReport,
  locked,
  parseReport,
  resolveWeekForWrite,
  saveReport,
  setTuanStatus,
} from "./plan.ts";

export type MappedViolation = {
  ma: string;
  tapThe: boolean;
  paper: boolean;
};

export type ImportRow = {
  sheet: string;
  stt: string;
  ngay: string;
  lop: string;
  hoTen: string;
  loi: string;
  buoi: string;
  ghiChu: string;
};

export type ImportResult = {
  weeks: number;
  classes: number;
  events: number;
  skippedLocked: string[];
  unmatched: string[];
  unknownClasses: string[];
};

function fold(value: string) {
  return value
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/\s+/g, " ")
    .trim();
}

const PAPER = new Set(["nghi_hoc", "di_muon", "trang_phuc", "phu_hieu", "vp_khac", "thai_do"]);

export function mapViolation(loi: string, hoTen: string, con?: Db, namId?: number): MappedViolation {
  if (con && namId) {
    const hit = matchLoi(con, namId, loi, hoTen);
    if (hit) return { ma: hit.ma, tapThe: hit.tapThe, paper: PAPER.has(hit.ma) };
  }
  const text = fold(loi);
  const named = hoTen.trim().length > 0;
  const hit = matchLoiFallback(text, named);
  return hit;
}

function matchLoiFallback(text: string, named: boolean): MappedViolation {
  if (/nghi hoc/.test(text)) return { ma: "nghi_hoc", tapThe: false, paper: true };
  if (/di (hoc )?muon/.test(text)) return { ma: "di_muon", tapThe: false, paper: true };
  if (/so vin|dep sai|dep le|trang phuc/.test(text)) return { ma: "trang_phuc", tapThe: false, paper: true };
  return { ma: "vp_khac", tapThe: !named, paper: true };
}

function excelDay(value: unknown): string {
  if (value instanceof Date && Number.isFinite(value.getTime())) {
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate())).toISOString().slice(0, 10);
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000).toISOString().slice(0, 10);
  }
  const text = String(value ?? "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  return "";
}

function cellText(value: unknown): string {
  if (value == null) return "";
  if (value instanceof Date) return excelDay(value);
  if (typeof value === "object" && value && "text" in value) return String((value as { text: unknown }).text ?? "").trim();
  if (typeof value === "object" && value && "result" in value) return cellText((value as { result: unknown }).result);
  return String(value).trim();
}

export async function readPhanAnhWorkbook(input: string | Buffer): Promise<ImportRow[]> {
  const wb = new ExcelJS.Workbook();
  if (typeof input === "string") {
    await wb.xlsx.readFile(input);
  } else {
    await wb.xlsx.load(input);
  }
  const rows: ImportRow[] = [];
  for (const sheet of wb.worksheets) {
    sheet.eachRow((row, index) => {
      if (index === 1) return;
      const ngay = excelDay(row.getCell(3).value);
      const lop = cellText(row.getCell(5).value);
      const hoTen = cellText(row.getCell(6).value);
      const loi = cellText(row.getCell(7).value);
      if (!lop && !loi && !hoTen) return;
      rows.push({
        sheet: sheet.name,
        stt: cellText(row.getCell(1).value),
        ngay,
        lop,
        hoTen,
        loi,
        buoi: cellText(row.getCell(4).value),
        ghiChu: cellText(row.getCell(8).value),
      });
    });
  }
  return rows;
}

function fridayOfLocal(value: string) {
  return fridayOf(value);
}

export async function importPhanAnhFile(con: Db, namId: number, filePath: string, opts?: { unlockIfNeeded?: boolean }): Promise<ImportResult> {
  const rows = await readPhanAnhWorkbook(filePath);
  return importPhanAnhRows(con, namId, rows, opts);
}

export async function importPhanAnhBuffer(con: Db, namId: number, buffer: Buffer, opts?: { unlockIfNeeded?: boolean }): Promise<ImportResult> {
  const rows = await readPhanAnhWorkbook(buffer);
  return importPhanAnhRows(con, namId, rows, opts);
}
export function importPhanAnhRows(con: Db, namId: number, rows: ImportRow[], opts?: { unlockIfNeeded?: boolean }): ImportResult {
  return transaction(con, () => importPhanAnhRowsUnlocked(con, namId, rows, opts));
}

function importPhanAnhRowsUnlocked(con: Db, namId: number, rows: ImportRow[], opts?: { unlockIfNeeded?: boolean }): ImportResult {
  const classes = listLop(con, namId);
  const byTen = new Map(classes.map((lop) => [String(lop.ten).toUpperCase(), lop]));
  const skippedLocked = new Set<string>();
  const unmatched = new Set<string>();
  const unknownClasses = new Set<string>();
  type Bucket = { weekStart: string; lopId: number; rows: ImportRow[] };
  const buckets = new Map<string, Bucket>();
  for (const row of rows) {
    if (!row.ngay || !row.loi) continue;
    const lop = byTen.get(row.lop.toUpperCase());
    if (!lop) {
      if (row.lop) unknownClasses.add(row.lop);
      continue;
    }
    const weekStart = fridayOfLocal(row.ngay);
    const key = `${weekStart}:${lop.id}`;
    const bucket = buckets.get(key) ?? { weekStart, lopId: Number(lop.id), rows: [] };
    bucket.rows.push(row);
    buckets.set(key, bucket);
  }

  let events = 0;
  for (const bucket of buckets.values()) {
    const week = resolveWeekForWrite(con, namId, { week_start: bucket.weekStart });
    if (locked(week)) {
      if (opts?.unlockIfNeeded) {
        if (week.trang_thai === "cong_bo") {
          setTuanStatus(con, namId, Number(week.id), Number(week.revision), "chot", "Mở khóa nạp phản ánh");
        }
        const refreshed = get(con, "SELECT revision FROM tuan WHERE id=?", [week.id]);
        setTuanStatus(con, namId, Number(week.id), Number(refreshed?.revision ?? 0), "nhap", "Mở khóa nạp phản ánh");
      } else {
        skippedLocked.add(String(week.ngay_bd || bucket.weekStart));
        continue;
      }
    }
    const existing = loadReport(con, Number(week.id), bucket.lopId);
    const form: Record<string, string> = {
      gio_tong: String(existing.bc?.gio_tong ?? 0),
      gio_tot: String(existing.bc?.gio_tot ?? 0),
      gio_kha: String(existing.bc?.gio_kha ?? 0),
      gio_tb: String(existing.bc?.gio_tb ?? 0),
      gio_yeu: String(existing.bc?.gio_yeu ?? 0),
      gio_kem: String(existing.bc?.gio_kem ?? 0),
      ktm_9_10: String(existing.bc?.ktm_9_10 ?? 0),
      ktm_7_8: String(existing.bc?.ktm_7_8 ?? 0),
      ktm_5_6: String(existing.bc?.ktm_5_6 ?? 0),
      ktm_3_4: String(existing.bc?.ktm_3_4 ?? 0),
      ktm_0_2: String(existing.bc?.ktm_0_2 ?? 0),
      bi_thu: String(existing.bc?.bi_thu ?? ""),
      ngay_lap: String(existing.bc?.ngay_lap || week.ngay_bd || bucket.weekStart),
      ghi_chu_gio: String(existing.bc?.ghi_chu_gio || "Nhập từ file phản ánh"),
      ghi_chu_ktm: String(existing.bc?.ghi_chu_ktm || "Nhập từ file phản ánh"),
    };
    let nghiI = 0;
    const paperI: Record<string, number> = { di_muon: 0, trang_phuc: 0, phu_hieu: 0, vp_khac: 0 };
    let vpI = 0;
    const pushExisting = (prefix: string, recs: { ho_ten?: unknown; ngay?: unknown; ghi_chu?: unknown; so_luong?: unknown; tieu_chi_id?: unknown; tap_the?: unknown }[], fields: string[]) => {
      recs.forEach((rec, i) => {
        for (const field of fields) {
          const value = rec[field as keyof typeof rec];
          if (value != null && value !== "") form[`${prefix}_${i}_${field}`] = String(value);
        }
      });
      return recs.length;
    };
    nghiI = pushExisting("nghi", existing.nghi, ["ho_ten", "ngay", "ghi_chu"]);
    for (const loai of Object.keys(paperI)) {
      paperI[loai] = pushExisting(loai, existing.sk.filter((ev) => ev.loai === loai && !ev.tieu_chi_id), ["ho_ten", "ngay", "so_luong", "ghi_chu"]);
    }
    vpI = pushExisting("vp", existing.sk.filter((ev) => ev.tieu_chi_id), ["tieu_chi_id", "ho_ten", "ngay", "so_luong", "tap_the", "ghi_chu"]);

    for (const row of bucket.rows) {
      const mapped = mapViolation(row.loi, row.hoTen, con, namId);
      const note = [row.buoi, row.loi, row.ghiChu].filter(Boolean).join(" · ");
      if (mapped.ma === "nghi_hoc") {
        form[`nghi_${nghiI}_ho_ten`] = row.hoTen || "Tập thể";
        form[`nghi_${nghiI}_ngay`] = row.ngay;
        form[`nghi_${nghiI}_ghi_chu`] = note;
        nghiI += 1;
        events += 1;
        continue;
      }
      if (mapped.paper) {
        const loai = mapped.ma;
        const i = paperI[loai] ?? 0;
        form[`${loai}_${i}_ho_ten`] = mapped.tapThe ? "" : row.hoTen;
        form[`${loai}_${i}_ngay`] = row.ngay;
        form[`${loai}_${i}_so_luong`] = "1";
        form[`${loai}_${i}_ghi_chu`] = note;
        if (mapped.tapThe) form[`${loai}_${i}_tap_the`] = "1";
        if (!mapped.tapThe && !row.hoTen) form[`${loai}_${i}_ho_ten`] = "Không rõ tên";
        paperI[loai] = i + 1;
        events += 1;
        continue;
      }
      const criterion = get(con, "SELECT id FROM tieu_chi WHERE nam_hoc_id=? AND ma=? AND ap_dung=1", [namId, mapped.ma]);
      if (!criterion) {
        unmatched.add(row.loi);
        const i = paperI.vp_khac;
        form[`vp_khac_${i}_ho_ten`] = row.hoTen || "Không rõ tên";
        form[`vp_khac_${i}_ngay`] = row.ngay;
        form[`vp_khac_${i}_so_luong`] = "1";
        form[`vp_khac_${i}_ghi_chu`] = note;
        paperI.vp_khac = i + 1;
        events += 1;
        continue;
      }
      form[`vp_${vpI}_tieu_chi_id`] = String(criterion.id);
      form[`vp_${vpI}_ho_ten`] = mapped.tapThe ? "" : (row.hoTen || "Không rõ tên");
      form[`vp_${vpI}_ngay`] = row.ngay;
      form[`vp_${vpI}_so_luong`] = "1";
      form[`vp_${vpI}_tap_the`] = mapped.tapThe ? "1" : "0";
      form[`vp_${vpI}_ghi_chu`] = note;
      vpI += 1;
      events += 1;
    }
    const parsed = parseReport(form, week);
    if (Object.keys(parsed.errors).length) {
      throw new WorkflowError(400, `Không nhập được ${get(con, "SELECT ten FROM lop WHERE id=?", [bucket.lopId])?.ten} tuần ${week.ngay_bd}: ${Object.values(parsed.errors)[0]}`);
    }
    saveReport(con, namId, { tuan_id: Number(week.id), week_start: String(week.ngay_bd) }, bucket.lopId, Number(existing.bc?.revision ?? 0), parsed, "save");
  }
  return {
    weeks: new Set([...buckets.values()].map((b) => b.weekStart)).size,
    classes: new Set([...buckets.values()].map((b) => b.lopId)).size,
    events,
    skippedLocked: [...skippedLocked],
    unmatched: [...unmatched],
    unknownClasses: [...unknownClasses],
  };
}
