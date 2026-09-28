(() => {
  const prefixes = {
    nghi: ["ho_ten", "ngay", "ghi_chu"],
    di_muon: ["ho_ten", "ngay", "so_luong", "ghi_chu"],
    trang_phuc: ["ho_ten", "ngay", "so_luong", "ghi_chu"],
    phu_hieu: ["ho_ten", "ngay", "so_luong", "ghi_chu"],
    vp_khac: ["ho_ten", "ngay", "so_luong", "ghi_chu"],
    thai_do: ["ho_ten", "tiet_mon", "noi_dung", "ngay", "ghi_chu"],
    vp: ["tieu_chi_id", "ho_ten", "ngay", "so_luong", "tap_the", "gvcn_phat_hien", "ghi_chu"],
  };
  const GIO_BUCKETS = ["gio_tot", "gio_kha", "gio_tb", "gio_yeu", "gio_kem"];
  const KTM_BUCKETS = ["ktm_9_10", "ktm_7_8", "ktm_5_6", "ktm_3_4", "ktm_0_2"];
  const forms = [...document.querySelectorAll("form[data-dirty-guard]")];
  let dirty = false;
  const mark = () => {
    dirty = true;
    document.querySelectorAll("[data-save-state]").forEach((node) => { node.textContent = "Chưa lưu"; });
  };
  const entry = (() => {
    try {
      return JSON.parse(document.getElementById("click-entry-data")?.textContent || "{}");
    } catch {
      return {};
    }
  })();
  const days = Array.isArray(entry.days) ? entry.days : [];
  let lastDay = days[0]?.value || "";
  const today = (() => {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, "0");
    const d = String(now.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  })();
  if (days.some((day) => day.value === today)) lastDay = today;

  const num = (id) => {
    const el = document.getElementById(id);
    if (!(el instanceof HTMLInputElement)) return 0;
    const n = Number(el.value);
    return Number.isFinite(n) ? n : 0;
  };
  const setNum = (id, value) => {
    const el = document.getElementById(id);
    if (!(el instanceof HTMLInputElement)) return;
    el.value = String(Math.max(0, Math.trunc(value)));
    el.dispatchEvent(new Event("input", { bubbles: true }));
  };

  function wrapStepper(input) {
    if (!(input instanceof HTMLInputElement) || input.type !== "number" || input.closest(".stepper")) return;
    if (input.step === "any") return;
    const wrap = document.createElement("div");
    wrap.className = "stepper";
    const minus = document.createElement("button");
    minus.type = "button";
    minus.className = "step-btn";
    minus.dataset.step = "-1";
    minus.setAttribute("aria-label", "Giảm");
    minus.textContent = "−";
    const plus = document.createElement("button");
    plus.type = "button";
    plus.className = "step-btn";
    plus.dataset.step = "1";
    plus.setAttribute("aria-label", "Tăng");
    plus.textContent = "+";
    input.parentNode?.insertBefore(wrap, input);
    wrap.append(minus, input, plus);
  }

  function replaceDate(input) {
    if (!(input instanceof HTMLInputElement) || input.type !== "date" || !days.length) return;
    if (input.dataset.daySelect === "1") return;
    const select = document.createElement("select");
    select.name = input.name;
    if (input.id) select.id = input.id;
    select.dataset.daySelect = "1";
    const blank = document.createElement("option");
    blank.value = "";
    blank.textContent = "— Ngày —";
    select.append(blank);
    let matched = false;
    for (const day of days) {
      const opt = document.createElement("option");
      opt.value = day.value;
      opt.textContent = day.label;
      if (input.value === day.value) {
        opt.selected = true;
        matched = true;
      }
      select.append(opt);
    }
    if (input.value && !matched) {
      const extra = document.createElement("option");
      extra.value = input.value;
      extra.textContent = `${input.value} · chọn lại`;
      extra.selected = true;
      select.append(extra);
    }
    input.replaceWith(select);
  }

  function bindName(input) {
    if (!(input instanceof HTMLInputElement)) return;
    input.setAttribute("list", "ten-lop");
    input.autocomplete = "off";
  }

  function fillDayChips(root, selected) {
    root.replaceChildren();
    for (const day of days) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "chip" + (day.value === selected ? " on" : "");
      btn.dataset.day = day.value;
      btn.textContent = day.label;
      root.append(btn);
    }
  }

  function selectedDay(composer) {
    const sel = composer.querySelector("select[data-composer-day], .composer-day-select");
    if (sel && sel.value) return sel.value;
    return composer.querySelector(".chip.on")?.dataset.day || lastDay || days[0]?.value || "";
  }

  function nextIndex(prefix) {
    const tbody = document.querySelector(`#tbl-${CSS.escape(prefix)} tbody`);
    if (!tbody) return 0;
    const indexes = [...tbody.querySelectorAll(`[name^="${prefix}_"]`)]
      .map((el) => Number(el.name.match(new RegExp(`^${prefix}_(\\d+)_`))?.[1]))
      .filter(Number.isFinite);
    return indexes.length ? Math.max(...indexes) + 1 : 0;
  }

  function fieldControl(prefix, index, field) {
    const name = `${prefix}_${index}_${field}`;
    if (field === "tieu_chi_id") {
      const options = document.getElementById("vp-tieu-chi-options")?.innerHTML
        || `<option value="">— Chọn tiêu chí —</option>`;
      return `<select name="${name}">${options}</select>`;
    }
    if (field === "tap_the" || field === "gvcn_phat_hien") {
      const label = field === "tap_the" ? "Tập thể" : "GVCN phát hiện";
      return `<input type="checkbox" name="${name}" value="1" aria-label="${label}">`;
    }
    if (field === "ngay") {
      const opts = [`<option value="">— Ngày —</option>`]
        .concat(days.map((day) => `<option value="${day.value}">${day.label}</option>`))
        .join("");
      return days.length
        ? `<select name="${name}" data-day-select="1">${opts}</select>`
        : `<input name="${name}" type="date">`;
    }
    if (field === "so_luong") return `<input name="${name}" type="number" min="1" step="1" value="1" inputmode="numeric">`;
    if (["ghi_chu", "noi_dung"].includes(field)) return `<textarea name="${name}"></textarea>`;
    if (field === "ho_ten") return `<input name="${name}" list="ten-lop" autocomplete="off">`;
    return `<input name="${name}">`;
  }

  function addRow(prefix) {
    const tbody = document.querySelector(`#tbl-${CSS.escape(prefix)} tbody`);
    if (!tbody || !prefixes[prefix]) return null;
    const index = nextIndex(prefix);
    const tr = document.createElement("tr");
    tr.innerHTML = prefixes[prefix].map((field) => `<td>${fieldControl(prefix, index, field)}</td>`).join("")
      + `<td><button type="button" class="btn ghost sm" data-delete-row aria-label="Xóa dòng ${index + 1}">Xóa dòng</button></td>`;
    tbody.appendChild(tr);
    tr.querySelectorAll("input[type=number]").forEach(wrapStepper);
    refreshEntryCount(tbody.closest(".entry-block"));
    return { tr, index };
  }

  function refreshEntryCount(block) {
    if (!(block instanceof HTMLElement)) return;
    const n = block.querySelector("summary .entry-n");
    if (!n || /không/.test(n.textContent || "")) return;
    n.textContent = String(block.querySelectorAll("tbody tr").length);
  }

  function setField(tr, prefix, index, field, value) {
    const el = tr.querySelector(`[name="${prefix}_${index}_${field}"]`);
    if (el instanceof HTMLInputElement) {
      if (el.type === "checkbox") el.checked = value === "1" || value === true;
      else el.value = value ?? "";
    } else if (el instanceof HTMLSelectElement || el instanceof HTMLTextAreaElement) {
      el.value = value ?? "";
    }
  }

  function notifyError(message) {
    window.alert(message);
  }

  function flashComposer(composer, message) {
    notifyError(message);
    composer.classList.add("invalid");
    const old = composer.querySelector(".composer-error");
    old?.remove();
    const note = document.createElement("span");
    note.className = "composer-error";
    note.textContent = message;
    composer.append(note);
    setTimeout(() => {
      composer.classList.remove("invalid");
      note.remove();
    }, 4000);
  }

  function addFromComposer(composer) {
    const prefix = composer.dataset.composer;
    if (!prefix) return;
    const day = selectedDay(composer);
    const name = composer.querySelector(".name-pick")?.value.trim() || "";
    const tap = composer.querySelector("[data-composer-tap]")?.checked;
    const gvcn = composer.querySelector("[data-composer-gvcn]")?.checked;
    const tc = composer.querySelector("[data-composer-tc]")?.value || "";
    const tiet = composer.querySelector("[data-composer-tiet]")?.value.trim() || "";
    const nd = composer.querySelector("[data-composer-nd]")?.value.trim() || "";
    if (prefix === "vp" && !tc) return flashComposer(composer, "Chọn tiêu chí trước.");
    if (prefix === "vp" && !tap && !name) return flashComposer(composer, "Gõ hoặc chọn họ tên, hoặc đánh Tập thể.");
    if (prefix !== "vp" && !name) return flashComposer(composer, "Gõ hoặc chọn họ tên.");
    if (prefix === "thai_do" && !nd) return flashComposer(composer, "Ghi nội dung lỗi.");
    const added = addRow(prefix);
    if (!added) return;
    const { tr, index } = added;
    setField(tr, prefix, index, "ho_ten", name);
    setField(tr, prefix, index, "ngay", day);
    setField(tr, prefix, index, "so_luong", "1");
    if (prefix === "vp") {
      setField(tr, prefix, index, "tieu_chi_id", tc);
      setField(tr, prefix, index, "tap_the", tap ? "1" : "0");
      setField(tr, prefix, index, "gvcn_phat_hien", gvcn ? "1" : "0");
    }
    if (prefix === "thai_do") {
      setField(tr, prefix, index, "tiet_mon", tiet);
      setField(tr, prefix, index, "noi_dung", nd);
    }
    lastDay = day || lastDay;
    const nameInput = composer.querySelector(".name-pick");
    if (nameInput instanceof HTMLInputElement) nameInput.value = "";
    const ndInput = composer.querySelector("[data-composer-nd]");
    if (ndInput instanceof HTMLInputElement) ndInput.value = "";
    mark();
    nameInput?.focus();
  }

  function updateGioMeter() {
    const meter = document.querySelector("[data-gio-meter]");
    if (!meter) return;
    const total = num("gio_tong");
    const classified = GIO_BUCKETS.reduce((sum, key) => sum + num(key), 0);
    const rest = total - classified;
    meter.classList.remove("ok", "warn", "bad");
    if (classified === total && total >= 0) {
      meter.classList.add("ok");
      meter.textContent = `Đã xếp loại ${classified}/${total} giờ.`;
    } else if (classified > total) {
      meter.classList.add("bad");
      meter.textContent = `Đã xếp ${classified} giờ — vượt tổng ${total}. Bấm − ở mức thừa.`;
    } else {
      meter.classList.add("warn");
      meter.textContent = `Còn ${rest} giờ chưa xếp loại. Bấm + phần còn ở đúng mức, hoặc ghi lý do bên dưới.`;
    }
    document.querySelectorAll("[data-fill-rest]").forEach((btn) => {
      btn.hidden = rest <= 0;
    });
  }

  function updateKtmMeter() {
    const meter = document.querySelector("[data-ktm-meter]");
    if (!meter) return;
    const total = KTM_BUCKETS.reduce((sum, key) => sum + num(key), 0);
    meter.classList.remove("ok", "warn", "bad");
    meter.classList.add(total === 0 ? "warn" : "ok");
    meter.textContent = total === 0
      ? "Chưa có lượt điểm miệng — bấm «Không phát sinh» nếu giấy ghi không có."
      : `${total} lượt điểm miệng.`;
  }

  function updateExamMeter() {
    const meter = document.querySelector("[data-exam-meter]");
    if (!meter) return;
    const max = Number(meter.dataset.max || 0);
    const keys = (meter.dataset.keys || "").split(",").filter(Boolean);
    const total = keys.reduce((sum, key) => sum + num(key), 0);
    meter.classList.remove("ok", "warn", "bad");
    if (!max) {
      meter.textContent = `Đã nhập ${total}.`;
      return;
    }
    if (total === max) {
      meter.classList.add("ok");
      meter.textContent = `Tổng xếp loại ${total} = sĩ số ${max}.`;
    } else if (total > max) {
      meter.classList.add("bad");
      meter.textContent = `Tổng ${total} vượt sĩ số ${max}.`;
    } else {
      meter.classList.add("warn");
      meter.textContent = `Đã nhập ${total} / sĩ số ${max} — còn ${max - total}.`;
    }
  }

  function enhance(root) {
    root.querySelectorAll("input[type=date]").forEach(replaceDate);
    root.querySelectorAll("input[type=number]").forEach(wrapStepper);
    root.querySelectorAll("input[name*='_ho_ten']").forEach(bindName);
  }

  document.querySelectorAll("[data-composer-days]").forEach((root) => fillDayChips(root, lastDay));
  document.querySelectorAll("form[data-click-entry], form[data-stepper-form]").forEach(enhance);
  updateGioMeter();
  updateKtmMeter();
  updateExamMeter();

  for (const form of forms) {
    form.addEventListener("input", mark);
    form.addEventListener("change", mark);
    form.addEventListener("submit", (event) => {
      const submitter = event.submitter;
      if (submitter instanceof HTMLButtonElement && submitter.name === "action" && submitter.value === "submit") {
        const total = num("gio_tong");
        const classified = GIO_BUCKETS.reduce((sum, key) => sum + num(key), 0);
        const gioNote = document.getElementById("ghi_chu_gio");
        const gioReason = gioNote instanceof HTMLTextAreaElement ? gioNote.value.trim() : "";
        if (classified !== total && !gioReason) {
          event.preventDefault();
          notifyError(`Chưa xếp đủ giờ. Đã xếp ${classified}/${total}. Bấm + phần còn ở đúng mức, hoặc ghi lý do bên dưới.`);
          document.getElementById("gio_tong")?.scrollIntoView({ block: "center" });
          return;
        }
        const ktmTotal = KTM_BUCKETS.reduce((sum, key) => sum + num(key), 0);
        const ktmNote = document.getElementById("ghi_chu_ktm");
        const ktmReason = ktmNote instanceof HTMLTextAreaElement ? ktmNote.value.trim() : "";
        if (ktmTotal === 0 && !ktmReason) {
          event.preventDefault();
          notifyError("Chưa có điểm miệng. Bấm «Không phát sinh điểm miệng» nếu giấy không có, hoặc nhập số lượt − / +.");
          document.getElementById("ktm_9_10")?.scrollIntoView({ block: "center" });
          return;
        }
      }
      dirty = false;
    });
  }
  document.addEventListener("input", (event) => {
    if (!(event.target instanceof Element)) return;
    if (event.target.id?.startsWith("gio_") || event.target.closest("[data-count-pad=gio]")) updateGioMeter();
    if (event.target.id?.startsWith("ktm_") || event.target.closest("[data-count-pad=ktm]")) updateKtmMeter();
    if (event.target.closest("[data-stepper-form]")) updateExamMeter();
  });

  document.addEventListener("click", (event) => {
    const chip = event.target.closest(".chip[data-day]");
    if (chip) {
      const group = chip.parentElement;
      group?.querySelectorAll(".chip").forEach((node) => node.classList.remove("on"));
      chip.classList.add("on");
      lastDay = chip.dataset.day || lastDay;
      const ngayLap = document.getElementById("ngay_lap");
      if (ngayLap instanceof HTMLSelectElement && chip.closest(".field")?.contains(ngayLap)) {
        ngayLap.value = lastDay;
        mark();
      }
      return;
    }
    const stepBtn = event.target.closest(".step-btn");
    if (stepBtn) {
      const wrap = stepBtn.closest(".stepper");
      const input = wrap?.querySelector("input[type=number]")
        || (stepBtn.dataset.for ? document.getElementById(stepBtn.dataset.for) : null);
      if (input instanceof HTMLInputElement) {
        const delta = Number(stepBtn.dataset.step || 1);
        const min = input.min === "" ? 0 : Number(input.min);
        const next = Math.max(min, Math.trunc(Number(input.value || 0) + delta));
        input.value = String(next);
        input.dispatchEvent(new Event("input", { bubbles: true }));
        mark();
      }
      return;
    }
    const fill = event.target.closest("[data-fill-rest]");
    if (fill) {
      const rest = num("gio_tong") - GIO_BUCKETS.reduce((sum, key) => sum + num(key), 0);
      if (rest > 0) setNum(fill.dataset.fillRest, num(fill.dataset.fillRest) + rest);
      return;
    }
    if (event.target.closest("[data-ktm-none]")) {
      for (const key of KTM_BUCKETS) setNum(key, 0);
      const note = document.getElementById("ghi_chu_ktm");
      if (note instanceof HTMLTextAreaElement && !note.value.trim()) note.value = "Không phát sinh";
      mark();
      updateKtmMeter();
      return;
    }
    const composerAdd = event.target.closest("[data-composer-add]");
    if (composerAdd) {
      const composer = composerAdd.closest("[data-composer]");
      if (composer) addFromComposer(composer);
      return;
    }
    const add = event.target.closest("[data-add-row]");
    if (add) {
      const added = addRow(add.dataset.addRow);
      added?.tr.querySelector("input,select,textarea")?.focus();
      mark();
      return;
    }
    const remove = event.target.closest("[data-delete-row]");
    if (remove) {
      const block = remove.closest(".entry-block");
      remove.closest("tr")?.remove();
      refreshEntryCount(block);
      mark();
      return;
    }
    const link = event.target.closest("a[href]");
    if (dirty && link && !confirm("Có thay đổi chưa lưu. Chọn OK để bỏ thay đổi, Cancel để ở lại.")) event.preventDefault();
  });
  document.addEventListener("keydown", (event) => {
    const target = event.target;
    if (!(target instanceof Element)) return;
    if (event.key === "Enter" && target.closest("[data-composer]")) {
      const composer = target.closest("[data-composer]");
      if (composer && !(target instanceof HTMLTextAreaElement) && !(target instanceof HTMLButtonElement && target.name === "action")) {
        event.preventDefault();
        addFromComposer(composer);
        return;
      }
    }
    const form = target.closest("form.report-form[data-dirty-guard]");
    if (!form || !event.ctrlKey) return;
    const key = event.key.toLowerCase();
    const submitWith = (value) => {
      const button = form.querySelector(`button[name=action][value="${value}"]`);
      if (!(button instanceof HTMLButtonElement) || button.disabled) return;
      event.preventDefault();
      form.requestSubmit(button);
    };
    if (key === "s") submitWith("save");
    else if (key === "enter") submitWith("submit");
  });
  addEventListener("beforeunload", (event) => { if (dirty) { event.preventDefault(); event.returnValue = ""; } });
})();
