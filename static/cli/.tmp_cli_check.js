
(() => {
  // ---------- DOM ----------
  const $ = (id) => document.getElementById(id);
  const fileInput = $("fileInput");
  const btnSample = $("btnSample");
  const btnConvert = $("btnConvert");
  const btnDownload = $("btnDownload");
  const statusPill = $("statusPill");
  const metaBox = $("meta");
  const inpBox = $("inp");
  const outBox = $("out");
  const inpCount = $("inpCount");
  const outCount = $("outCount");
  const btnCopyOut = $("btnCopyOut");
  const optOnlyChanges = $("optOnlyChanges");
  const diffBox = $("diff");
  const diffSummary = $("diffSummary");
  const downloadLinks = $("downloadLinks");
  const previewRow = $("previewRow");
  const previewSelect = $("previewSelect");
  const fileDropZone = $("file-drop-zone");
  const filePickButton = $("fileInput-pick");
  const fileSelectedName = $("fileInput-selected");

  const optCustom = $("optCustom");
  const customRow = $("customRow");
  const pickDate = $("pickDate");
  const pickDateNative = $("pickDateNative");
  const pickDatePickerBtn = $("pickDatePickerBtn");
  const timeStartHour = $("timeStartHour");
  const timeStartMinute = $("timeStartMinute");
  const timeStartSecond = $("timeStartSecond");
  const timeEndHour = $("timeEndHour");
  const timeEndMinute = $("timeEndMinute");
  const timeEndSecond = $("timeEndSecond");
  const rangeHint = $("rangeHint");
  const autoFixModal = $("auto-fix-modal");
  const autoFixModalText = $("auto-fix-modal-text");
  const autoFixModalOk = $("auto-fix-modal-ok");
  const autoFixModalClose = $("auto-fix-modal-close");

  // ---------- State ----------
  let currentText = "";
  let lastResult = null; // { output, meta, serials, input }
  let selectedInputs = []; // [{ name, text }]
  let convertedResults = []; // [{ fileName, input, output, meta, serials }]
  let activeIndex = -1;
  let metaDetailsPinned = false;
  const OLD_CLOCK_TRANSFER_QUERY = "import_old_clock";
  const OLD_CLOCK_TRANSFER_PREFIX = "cli_old_clock_";
  const OLD_CLOCK_TRANSFER_MESSAGE = "cli_old_clock_transfer";
  const OLD_CLOCK_TRANSFER_ACK = "cli_old_clock_transfer_ack";

  // ---------- Small utils ----------
  const pad2 = (n) => String(n).padStart(2, "0");
  const pad3 = (n) => String(n).padStart(3, "0");
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const monthMap = Object.fromEntries(months.map((m,i)=>[m,i]));

  function setStatus(text, ok=true){
    statusPill.textContent = text;
    statusPill.style.borderColor = ok ? "#ddd" : "#ffccd5";
    statusPill.style.background = ok ? "#fcfcfc" : "#fff5f7";
  }

  function buildAutoFixNotice(issue){
    const foundDate = String(issue?.foundDate || "").trim();
    const oldestAllowedDate = String(issue?.oldestAllowedDate || "").trim();
    const reason = foundDate && oldestAllowedDate
      ? `เนื่องจากระบบตรวจพบเวลา show clock เก่าเกินเงื่อนไข (วันที่ที่พบ ${foundDate} เก่ากว่าวันที่ต่ำสุดที่ยอมรับ ${oldestAllowedDate})`
      : "เนื่องจากระบบตรวจพบเวลา show clock เก่าเกินเงื่อนไข";
    return `ระบบทำการ Convert ใหม่และตั้งวันที่เป็นวันที่ปัจจุบันแล้ว ${reason} หากต้องการเปลี่ยนวัน/เดือน/ปี ให้เลื่อนขึ้นไปด้านบนแล้วปรับค่าได้ทันที`;
  }

  function showAutoFixModal(message){
    if(!autoFixModal || !autoFixModalText) return;
    autoFixModalText.textContent = String(message || "");
    autoFixModal.hidden = false;
    autoFixModal.setAttribute("aria-hidden", "false");
    document.body.style.overflow = "hidden";
  }

  function hideAutoFixModal(){
    if(!autoFixModal) return;
    autoFixModal.hidden = true;
    autoFixModal.setAttribute("aria-hidden", "true");
    document.body.style.overflow = "";
  }

  function bindAutoFixModal(){
    autoFixModalOk?.addEventListener("click", hideAutoFixModal);
    autoFixModalClose?.addEventListener("click", hideAutoFixModal);
    autoFixModal?.addEventListener("click", (event) => {
      if(event.target === autoFixModal){
        hideAutoFixModal();
      }
    });
    window.addEventListener("keydown", (event) => {
      if(event.key === "Escape" && autoFixModal && !autoFixModal.hidden){
        hideAutoFixModal();
      }
    });
  }
  bindAutoFixModal();

  function setFileSelectionUI(entries){
    const count = Array.isArray(entries) ? entries.length : 0;
    if(fileSelectedName){
      if(count === 0) fileSelectedName.textContent = "ยังไม่เลือกไฟล์";
      else if(count === 1) fileSelectedName.textContent = entries[0].name || "1 file selected";
      else fileSelectedName.textContent = `เลือกแล้ว ${count} ไฟล์`;
    }
    if(fileDropZone){
      fileDropZone.classList.toggle("filled", count > 0);
    }
  }

  function escapeHtml(s){
    return String(s)
      .replaceAll("&","&amp;")
      .replaceAll("<","&lt;")
      .replaceAll(">","&gt;")
      .replaceAll('"',"&quot;")
      .replaceAll("'","&#039;");
  }

  function countLines(text){
    if(text === null || text === undefined) return 0;
    const t = String(text);
    if(t.length === 0) return 0;
    // Keep consistent with split used elsewhere
    return t.split(/\r?\n/).length;
  }

  function updateLineCounts(){
    if(inpCount) inpCount.textContent = `บรรทัดทั้งหมด: ${countLines(inpBox.value)}`;
    if(outCount) outCount.textContent = `บรรทัดทั้งหมด: ${countLines(outBox.value)}`;
  }


  function sanitizeFilenamePart(name, fallback="file"){
    const cleaned = String(name || "")
      .replace(/[^\w.-]+/g, "_")
      .replace(/^_+|_+$/g, "");
    return cleaned || fallback;
  }

  function splitFilename(name){
    const s = String(name || "");
    const idx = s.lastIndexOf(".");
    if(idx <= 0) return { stem: s || "file", ext: "" };
    return { stem: s.slice(0, idx), ext: s.slice(idx) };
  }

  function makeUniqueFilename(name, used){
    const base = splitFilename(name);
    let candidate = name;
    let n = 2;
    while(used.has(candidate.toLowerCase())){
      candidate = `${base.stem}_${n}${base.ext}`;
      n++;
    }
    used.add(candidate.toLowerCase());
    return candidate;
  }

  function normalizeTimeValue(v){
    if(!v) return "00:00:00";
    return v.length===5 ? (v + ":00") : v;
  }

  function buildTimeOptions(selectEl, max){
    if(!selectEl) return;
    selectEl.innerHTML = "";
    for(let i=0;i<=max;i++){
      const v = pad2(i);
      const opt = document.createElement("option");
      opt.value = v;
      opt.textContent = v;
      selectEl.appendChild(opt);
    }
  }

  function setSelectTime(hourEl, minuteEl, secondEl, hms){
    const [h,m,s] = normalizeTimeValue(hms).split(":").map(n => pad2(Number(n)));
    hourEl.value = h;
    minuteEl.value = m;
    secondEl.value = s;
  }

  function getSelectTime(hourEl, minuteEl, secondEl){
    const h = Number(hourEl?.value ?? 0);
    const m = Number(minuteEl?.value ?? 0);
    const s = Number(secondEl?.value ?? 0);
    return `${pad2(h)}:${pad2(m)}:${pad2(s)}`;
  }

  function initTimeSelectors(){
    buildTimeOptions(timeStartHour, 23);
    buildTimeOptions(timeEndHour, 23);
    buildTimeOptions(timeStartMinute, 59);
    buildTimeOptions(timeEndMinute, 59);
    buildTimeOptions(timeStartSecond, 59);
    buildTimeOptions(timeEndSecond, 59);
    setSelectTime(timeStartHour, timeStartMinute, timeStartSecond, "08:00:00");
    setSelectTime(timeEndHour, timeEndMinute, timeEndSecond, "18:00:00");
  }

  function formatDateYmd(dateObj){
    const d = dateObj instanceof Date ? dateObj : new Date();
    return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
  }

  function formatDateDmy(dateObj){
    const d = dateObj instanceof Date ? dateObj : new Date();
    return `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}/${d.getFullYear()}`;
  }

  function formatDatePartsToDmy(parts){
    if(!parts) return "";
    return `${pad2(parts.day)}/${pad2(parts.month)}/${parts.year}`;
  }

  function parseDateInputToYmd(value){
    const raw = String(value || "").trim();
    if(!raw) return null;

    const validate = (year, month, day) => {
      if(!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return null;
      if(year < 1900 || year > 9999) return null;
      if(month < 1 || month > 12) return null;
      if(day < 1 || day > 31) return null;
      const check = new Date(Date.UTC(year, month - 1, day));
      if(
        check.getUTCFullYear() !== year ||
        check.getUTCMonth() !== month - 1 ||
        check.getUTCDate() !== day
      ){
        return null;
      }
      return { year, month, day };
    };

    let m = raw.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
    if(m){
      return validate(Number(m[3]), Number(m[2]), Number(m[1]));
    }

    // Backward-compatible: accept YYYY-MM-DD from old saved values.
    m = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if(m){
      return validate(Number(m[1]), Number(m[2]), Number(m[3]));
    }

    return null;
  }

  function isLeapYear(year){
    if(!Number.isInteger(year)) return false;
    return year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  }

  function getMaxDayByMonth(month, year){
    if(!Number.isInteger(month) || month < 1 || month > 12) return 31;
    if(month === 2){
      if(Number.isInteger(year)) return isLeapYear(year) ? 29 : 28;
      return 29;
    }
    if(month === 4 || month === 6 || month === 9 || month === 11) return 30;
    return 31;
  }

  function clampNumber(value, min, max){
    return Math.min(max, Math.max(min, value));
  }

  const DATE_SLOT_POSITIONS = [0, 1, 3, 4, 6, 7, 8, 9];

  function findDateSlotAtOrAfter(index){
    for(const pos of DATE_SLOT_POSITIONS){
      if(pos >= index) return pos;
    }
    return null;
  }

  function findDateSlotBefore(index){
    for(let i=DATE_SLOT_POSITIONS.length-1;i>=0;i--){
      if(DATE_SLOT_POSITIONS[i] < index) return DATE_SLOT_POSITIONS[i];
    }
    return null;
  }

  function clearDateSlotsInRange(chars, start, end){
    for(const pos of DATE_SLOT_POSITIONS){
      if(pos >= start && pos < end){
        chars[pos] = "_";
      }
    }
  }

  function normalizeDateTextInput(value){
    const raw = String(value || "");
    const chars = "__/__/____".split("");

    if(raw.length >= 10 && raw[2] === "/" && raw[5] === "/"){
      for(const pos of DATE_SLOT_POSITIONS){
        const ch = raw[pos];
        chars[pos] = /\d/.test(ch || "") ? ch : "_";
      }
    } else {
      const digits = raw.replace(/\D+/g, "").slice(0, 8);
      for(let i=0;i<digits.length;i++){
        chars[DATE_SLOT_POSITIONS[i]] = digits[i];
      }
    }

    const mmRaw = `${chars[3]}${chars[4]}`;
    let monthForCalc = null;
    if(/^\d\d$/.test(mmRaw)){
      const m2 = clampNumber(Number(mmRaw), 1, 12);
      const mm = pad2(m2);
      chars[3] = mm[0];
      chars[4] = mm[1];
      monthForCalc = m2;
    } else if(/^\d_$/.test(mmRaw)){
      const m1 = Number(chars[3]);
      monthForCalc = m1 >= 1 && m1 <= 9 ? m1 : null;
    }

    const yyyyRaw = `${chars[6]}${chars[7]}${chars[8]}${chars[9]}`;
    const yearMaybe = /^\d{4}$/.test(yyyyRaw) ? Number(yyyyRaw) : null;
    const maxDay = getMaxDayByMonth(monthForCalc, yearMaybe);
    const ddRaw = `${chars[0]}${chars[1]}`;
    if(/^\d\d$/.test(ddRaw)){
      const d2 = clampNumber(Number(ddRaw), 1, maxDay);
      const dd = pad2(d2);
      chars[0] = dd[0];
      chars[1] = dd[1];
    }

    return chars.join("");
  }

  function countDigitsBeforeIndex(text, index){
    const src = String(text || "");
    const end = Math.max(0, Math.min(Number(index) || 0, src.length));
    let count = 0;
    for(let i=0;i<end;i++){
      if(/\d/.test(src[i])) count++;
    }
    return count;
  }

  function caretIndexForDigitCount(text, digitCount){
    const src = String(text || "");
    const target = Math.max(0, Number(digitCount) || 0);
    if(target === 0) return 0;
    let seen = 0;
    for(let i=0;i<src.length;i++){
      if(/\d/.test(src[i])){
        seen++;
        if(seen >= target){
          return i + 1;
        }
      }
    }
    return src.length;
  }

  function syncNativeDateFromText(){
    if(!pickDateNative || !pickDate) return false;
    const parsed = parseDateInputToYmd(pickDate.value);
    if(!parsed){
      pickDateNative.value = "";
      return false;
    }
    pickDateNative.value = `${parsed.year}-${pad2(parsed.month)}-${pad2(parsed.day)}`;
    pickDate.value = formatDatePartsToDmy(parsed);
    return true;
  }

  function syncTextDateFromNative(){
    if(!pickDateNative || !pickDate) return false;
    const parsed = parseDateInputToYmd(pickDateNative.value);
    if(!parsed) return false;
    pickDate.value = formatDatePartsToDmy(parsed);
    return true;
  }

  function setDateToToday(){
    const now = new Date();
    if(pickDate){
      pickDate.value = formatDateDmy(now);
    }
    if(pickDateNative){
      pickDateNative.value = formatDateYmd(now);
    }
  }

  function bindDatePickerUi(){
    pickDatePickerBtn?.addEventListener("click", () => {
      if(!syncNativeDateFromText() && pickDateNative){
        // Keep native value empty so selecting a day always emits a value update.
        pickDateNative.value = "";
      }
      if(!pickDateNative) return;
      if(typeof pickDateNative.showPicker === "function"){
        pickDateNative.showPicker();
      } else {
        pickDateNative.focus();
        pickDateNative.click();
      }
    });

    pickDateNative?.addEventListener("input", () => {
      syncTextDateFromNative();
    });

    pickDateNative?.addEventListener("change", () => {
      syncTextDateFromNative();
    });

    pickDate?.addEventListener("keydown", (event) => {
      if(event.ctrlKey || event.metaKey || event.altKey) return;
      const key = event.key;
      const start = pickDate.selectionStart ?? 0;
      const end = pickDate.selectionEnd ?? start;
      const chars = normalizeDateTextInput(pickDate.value).split("");
      const applyValue = (caretPos) => {
        pickDate.value = chars.join("");
        const c = Math.max(0, Math.min(Number(caretPos) || 0, pickDate.value.length));
        pickDate.setSelectionRange(c, c);
        syncNativeDateFromText();
      };
      if(
        key === "Tab" ||
        key === "ArrowLeft" ||
        key === "ArrowRight" ||
        key === "ArrowUp" ||
        key === "ArrowDown" ||
        key === "Home" ||
        key === "End" ||
        key === "Enter"
      ){
        return;
      }
      if(key === "Backspace"){
        event.preventDefault();
        if(end > start){
          clearDateSlotsInRange(chars, start, end);
          applyValue(start);
          return;
        }
        const pos = findDateSlotBefore(start);
        if(pos === null) return;
        chars[pos] = "_";
        applyValue(pos);
        return;
      }
      if(key === "Delete"){
        event.preventDefault();
        if(end > start){
          clearDateSlotsInRange(chars, start, end);
          applyValue(start);
          return;
        }
        const pos = findDateSlotAtOrAfter(start);
        if(pos === null) return;
        chars[pos] = "_";
        applyValue(pos);
        return;
      }
      if(/^\d$/.test(key)){
        event.preventDefault();
        if(end > start){
          clearDateSlotsInRange(chars, start, end);
        }
        let pos = findDateSlotAtOrAfter(start);
        if(pos === null){
          pos = DATE_SLOT_POSITIONS[DATE_SLOT_POSITIONS.length - 1];
        }
        chars[pos] = key;
        const adjusted = normalizeDateTextInput(chars.join("")).split("");
        for(let i=0;i<adjusted.length;i++){
          chars[i] = adjusted[i];
        }
        const next = findDateSlotAtOrAfter(pos + 1);
        applyValue(next === null ? pos + 1 : next);
        return;
      }
      event.preventDefault();
    });

    pickDate?.addEventListener("input", () => {
      const beforeValue = pickDate.value;
      const caretBefore = pickDate.selectionStart ?? beforeValue.length;
      const digitCountBefore = countDigitsBeforeIndex(beforeValue, caretBefore);
      const normalized = normalizeDateTextInput(beforeValue);
      if(beforeValue !== normalized){
        pickDate.value = normalized;
        const nextCaret = caretIndexForDigitCount(normalized, digitCountBefore);
        pickDate.setSelectionRange(nextCaret, nextCaret);
      }
      syncNativeDateFromText();
    });

    pickDate?.addEventListener("blur", () => {
      syncNativeDateFromText();
    });
  }

  function validateCustomDateOrNotify(){
    if(!optCustom.checked) return true;
    const parsed = parseDateInputToYmd(pickDate?.value);
    if(parsed){
      syncNativeDateFromText();
      return true;
    }
    setStatus("วันที่ไม่ถูกต้อง", false);
    metaBox.textContent = "กรุณากรอกวันที่เป็น วัน/เดือน/ปี และกรอกได้เฉพาะตัวเลข เช่น 25/02/2026";
    pickDate?.focus();
    return false;
  }

  function toSec(hms){
    const [h,m,s] = hms.split(":").map(Number);
    return h*3600 + m*60 + s;
  }

  function randInt(min, max){
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }

  function dowStr(y,m,d){
    const days = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
    return days[new Date(Date.UTC(y,m,d)).getUTCDay()];
  }

  // ---------- Time line parse/format (preserve prefix incl. '*') ----------
  function parseTimeLine(line){
    // Example: "*08:12:21.902 UTC Tue Feb 3 2026"
    // Capture prefix (spaces + optional *), then time, optional ms, tz token, DOW, MON, DAY, YEAR
    const re = /^(\s*\*?\s*)(\d{1,2}):(\d{2}):(\d{2})(?:\.(\d{1,3}))?\s+(\S+)\s+([A-Za-z]{3})\s+([A-Za-z]{3})\s+(\d{1,2})\s+(\d{4})\s*$/;
    const m = line.match(re);
    if(!m) return null;
    const prefix = m[1] || "";
    const hh = +m[2], mm = +m[3], ss = +m[4];
    const ms = m[5] ? +(m[5].padEnd(3,"0")) : null;
    const tz = m[6];
    const dow = m[7];
    const mon = m[8];
    const day = +m[9];
    const year = +m[10];
    const monIdx = monthMap[mon];
    if(monIdx === undefined) return null;

    const dt = new Date(Date.UTC(year, monIdx, day, hh, mm, ss, ms ?? 0));
    return { dt, hasMs: ms !== null, tz, prefix, rawDow: dow };
  }

  function formatTimeLine(dt, tmpl){
    const y = dt.getUTCFullYear();
    const m = dt.getUTCMonth();
    const d = dt.getUTCDate();
    const hh = dt.getUTCHours();
    const mm = dt.getUTCMinutes();
    const ss = dt.getUTCSeconds();
    const ms = dt.getUTCMilliseconds();
    const dow = dowStr(y,m,d);
    const mon = months[m];
    const base = `${pad2(hh)}:${pad2(mm)}:${pad2(ss)}`;
    const t = tmpl.hasMs ? `${base}.${pad3(ms)}` : base;
    return `${tmpl.prefix}${t} ${tmpl.tz} ${dow} ${mon} ${d} ${y}`;
  }

  // ---------- Serial numbers ----------
  function findUniqueSerials(text){
    const re = /System\s+serial\s+number\s*:\s*([A-Za-z0-9_-]+)/gi;
    const set = new Set();
    let m;
    while((m = re.exec(text)) !== null){
      set.add(m[1]);
    }
    return Array.from(set);
  }

  // ---------- Diff (Myers line diff) ----------
  function backtrack(trace, a, b, offset){
    let x = a.length;
    let y = b.length;
    const edits = [];
    for(let d=trace.length-1; d>=0; d--){
      const v = trace[d];
      const k = x - y;
      const ki = k + offset;

      let prevK;
      if(k === -d || (k !== d && v[ki-1] < v[ki+1])) prevK = k + 1;
      else prevK = k - 1;

      const prevKi = prevK + offset;
      const prevX = v[prevKi];
      const prevY = prevX - prevK;

      while(x > prevX && y > prevY){
        edits.push({ type:"equal", line:a[x-1] });
        x--; y--;
      }
      if(d === 0) break;

      if(x === prevX){
        edits.push({ type:"insert", line:b[y-1] });
        y--;
      } else {
        edits.push({ type:"delete", line:a[x-1] });
        x--;
      }
    }
    edits.reverse();
    return edits;
  }

  function myersDiffLines(a, b){
    const N = a.length, M = b.length;
    const max = N + M;
    const offset = max;
    let v = new Array(2*max + 1).fill(0);
    const trace = [];

    for(let d=0; d<=max; d++){
      trace.push(v.slice());
      for(let k=-d; k<=d; k+=2){
        const ki = k + offset;
        let x;
        if(k === -d || (k !== d && v[ki-1] < v[ki+1])) x = v[ki+1];
        else x = v[ki-1] + 1;

        let y = x - k;
        while(x < N && y < M && a[x] === b[y]){ x++; y++; }
        v[ki] = x;

        if(x >= N && y >= M){
          return backtrack(trace, a, b, offset);
        }
      }
    }
    return backtrack(trace, a, b, offset);
  }

  function buildSideBySideRows(edits, onlyChanges){
    const rows = [];
    let inNo = 1;
    let outNo = 1;
    let skippedEqual = 0;

    const flushSkipped = () => {
      if(skippedEqual > 0){
        rows.push({ kind:"skip", count: skippedEqual });
        skippedEqual = 0;
      }
    };

    let i = 0;
    while(i < edits.length){
      const e = edits[i];

      if(e.type === "equal"){
        if(onlyChanges){
          skippedEqual++;
          inNo++;
          outNo++;
        } else {
          flushSkipped();
          rows.push({
            kind:"pair",
            left: { no: inNo, text: e.line, type: "equal" },
            right:{ no: outNo, text: e.line, type: "equal" },
          });
          inNo++;
          outNo++;
        }
        i++;
        continue;
      }

      const delBlock = [];
      const addBlock = [];
      while(i < edits.length && edits[i].type !== "equal"){
        if(edits[i].type === "delete") delBlock.push(edits[i].line);
        else if(edits[i].type === "insert") addBlock.push(edits[i].line);
        i++;
      }

      flushSkipped();
      const maxLen = Math.max(delBlock.length, addBlock.length);
      for(let idx=0; idx<maxLen; idx++){
        const hasDel = idx < delBlock.length;
        const hasAdd = idx < addBlock.length;
        const inline = (hasDel && hasAdd) ? buildInlineTokenDiff(delBlock[idx], addBlock[idx]) : null;
        rows.push({
          kind:"pair",
          left: hasDel
            ? { no: inNo++, text: delBlock[idx], type: "delete", html: inline?.leftHtml || "" }
            : { no: null, text: "", type: "empty" },
          right: hasAdd
            ? { no: outNo++, text: addBlock[idx], type: "insert", html: inline?.rightHtml || "" }
            : { no: null, text: "", type: "empty" },
        });
      }
    }

    if(onlyChanges) flushSkipped();
    return rows;
  }

  function normalizeDiffLineKey(text){
    return String(text || "")
      .toLowerCase()
      .replace(/\d+/g, "#")
      .replace(/[^a-z#]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function tokenOverlapScore(aKey, bKey){
    if(!aKey || !bKey) return 0;
    if(aKey === bKey) return 1;
    const aTokens = new Set(aKey.split(" ").filter(Boolean));
    const bTokens = new Set(bKey.split(" ").filter(Boolean));
    if(aTokens.size === 0 || bTokens.size === 0) return 0;
    let inter = 0;
    for(const tok of aTokens){
      if(bTokens.has(tok)) inter++;
    }
    const union = (aTokens.size + bTokens.size - inter) || 1;
    return inter / union;
  }

  function parseInterfaceErrorLine(text){
    const line = String(text || "");
    const m = line.match(/^\s*\d+\s+input\s+errors,\s*\d+\s+CRC,\s*\d+\s+frame,\s*\d+\s+overrun,\s*\d+\s+ignored(?:,\s*\d+\s+abort)?\s*$/i);
    if(!m) return null;
    return { hasAbort: /,\s*\d+\s+abort\s*$/i.test(line) };
  }

  function buildNormalizedInterfaceErrorLine(text){
    const line = String(text || "");
    const parsed = parseInterfaceErrorLine(line);
    if(!parsed) return null;
    const lead = (line.match(/^\s*/) || [""])[0];
    const tail = (line.match(/\s*$/) || [""])[0];
    return `${lead}0 input errors, 0 CRC, 0 frame, 0 overrun, 0 ignored${parsed.hasAbort ? ", 0 abort" : ""}${tail}`;
  }

  function pickNearestCandidateIndex(indices, pivot, consumed){
    if(!Array.isArray(indices) || indices.length === 0) return -1;
    let bestForward = -1;
    let bestForwardDist = Number.POSITIVE_INFINITY;
    let bestAny = -1;
    let bestAnyDist = Number.POSITIVE_INFINITY;

    for(const idx of indices){
      if(consumed.has(idx)) continue;
      const dist = Math.abs(idx - pivot);
      if(idx >= pivot && dist < bestForwardDist){
        bestForward = idx;
        bestForwardDist = dist;
      }
      if(dist < bestAnyDist){
        bestAny = idx;
        bestAnyDist = dist;
      }
    }
    return bestForward !== -1 ? bestForward : bestAny;
  }

  function buildInsertOnlyIndex(rows){
    const byNormalizedKey = new Map();
    const interfaceWithAbort = [];
    const interfaceWithoutAbort = [];
    const allInsertOnly = [];

    for(let idx=0; idx<rows.length; idx++){
      const row = rows[idx];
      if(!row || row.kind !== "pair") continue;
      if(row.left?.type !== "empty" || row.right?.type !== "insert") continue;

      allInsertOnly.push(idx);

      const key = normalizeDiffLineKey(row.right.text);
      if(key){
        if(!byNormalizedKey.has(key)) byNormalizedKey.set(key, []);
        byNormalizedKey.get(key).push(idx);
      }

      const parsed = parseInterfaceErrorLine(row.right.text);
      if(parsed){
        if(parsed.hasAbort) interfaceWithAbort.push(idx);
        else interfaceWithoutAbort.push(idx);
      }
    }

    return { byNormalizedKey, interfaceWithAbort, interfaceWithoutAbort, allInsertOnly };
  }

  function reconcileDeleteInsertRows(rows){
    const consumedInsertRows = new Set();
    const { byNormalizedKey, interfaceWithAbort, interfaceWithoutAbort, allInsertOnly } = buildInsertOnlyIndex(rows);
    const maxFuzzyDistance = 220;

    for(let i=0; i<rows.length; i++){
      const row = rows[i];
      if(!row || row.kind !== "pair") continue;
      if(row.left?.type !== "delete" || row.right?.type !== "empty") continue;

      const leftKey = normalizeDiffLineKey(row.left.text);
      if(!leftKey) continue;

      let bestIdx = -1;

      const parsedLeft = parseInterfaceErrorLine(row.left.text);
      if(parsedLeft){
        bestIdx = pickNearestCandidateIndex(
          parsedLeft.hasAbort ? interfaceWithAbort : interfaceWithoutAbort,
          i,
          consumedInsertRows
        );
      }

      if(bestIdx === -1){
        bestIdx = pickNearestCandidateIndex(byNormalizedKey.get(leftKey), i, consumedInsertRows);
      }

      if(bestIdx === -1){
        let bestScore = 0;
        let bestDist = Number.POSITIVE_INFINITY;
        for(const idx of allInsertOnly){
          if(consumedInsertRows.has(idx)) continue;
          const dist = Math.abs(idx - i);
          if(dist > maxFuzzyDistance) continue;
          const candidate = rows[idx];
          const rightKey = normalizeDiffLineKey(candidate?.right?.text);
          if(!rightKey) continue;
          const score = tokenOverlapScore(leftKey, rightKey);
          if(score > bestScore || (score === bestScore && dist < bestDist)){
            bestScore = score;
            bestDist = dist;
            bestIdx = idx;
          }
        }
        if(bestScore < 0.78) bestIdx = -1;
      }
      if(bestIdx === -1) continue;

      const matched = rows[bestIdx];
      row.right = {
        no: matched.right.no,
        text: matched.right.text,
        type: "insert",
        html: matched.right.html || "",
      };
      const inline = buildInlineTokenDiff(row.left.text, row.right.text);
      row.left.html = inline.leftHtml;
      row.right.html = inline.rightHtml;
      consumedInsertRows.add(bestIdx);
    }

    for(const row of rows){
      if(!row || row.kind !== "pair") continue;
      if(row.left?.type !== "delete" || row.right?.type !== "empty") continue;
      const normalized = buildNormalizedInterfaceErrorLine(row.left.text);
      if(!normalized || normalized === row.left.text) continue;
      row.right = { no: null, text: normalized, type: "insert", html: "" };
      const inline = buildInlineTokenDiff(row.left.text, row.right.text);
      row.left.html = inline.leftHtml;
      row.right.html = inline.rightHtml;
    }

    const compact = rows.filter((_, idx) => !consumedInsertRows.has(idx));
    const merged = [];
    for(const row of compact){
      const last = merged[merged.length - 1];
      if(last && last.kind === "skip" && row.kind === "skip"){
        last.count += row.count;
      } else {
        merged.push(row);
      }
    }
    return merged;
  }

  function diffSideClass(type){
    if(type === "delete") return "diffSideDel";
    if(type === "insert") return "diffSideAdd";
    if(type === "empty") return "diffSideEmpty";
    return "diffSideEq";
  }

  function tokenizeForInlineDiff(text){
    return String(text || "").match(/\s+|[^\s]+/g) || [];
  }

  function buildInlineTokenDiff(leftText, rightText){
    const a = tokenizeForInlineDiff(leftText);
    const b = tokenizeForInlineDiff(rightText);
    const n = a.length;
    const m = b.length;
    const dp = Array.from({ length: n + 1 }, () => Array(m + 1).fill(0));

    for(let i=n-1; i>=0; i--){
      for(let j=m-1; j>=0; j--){
        if(a[i] === b[j]) dp[i][j] = dp[i+1][j+1] + 1;
        else dp[i][j] = Math.max(dp[i+1][j], dp[i][j+1]);
      }
    }

    const leftSeg = [];
    const rightSeg = [];
    let i = 0;
    let j = 0;
    while(i < n && j < m){
      if(a[i] === b[j]){
        leftSeg.push({ t: a[i], c: false });
        rightSeg.push({ t: b[j], c: false });
        i++;
        j++;
      } else if(dp[i+1][j] >= dp[i][j+1]){
        leftSeg.push({ t: a[i], c: true });
        i++;
      } else {
        rightSeg.push({ t: b[j], c: true });
        j++;
      }
    }
    while(i < n){ leftSeg.push({ t: a[i], c: true }); i++; }
    while(j < m){ rightSeg.push({ t: b[j], c: true }); j++; }

    const renderSeg = (segments, cls) => segments.map((seg) => {
      const text = escapeHtml(seg.t);
      if(!seg.c || /^\s+$/.test(seg.t)) return text;
      return `<span class="${cls}">${text}</span>`;
    }).join("");

    return {
      leftHtml: renderSeg(leftSeg, "diffTokDel"),
      rightHtml: renderSeg(rightSeg, "diffTokAdd"),
    };
  }

  function renderDiffSideCell(side){
    const noText = side.no === null ? "" : String(side.no);
    const textHtml = side.text === ""
      ? "<span class=\"diffBlank\">&nbsp;</span>"
      : (side.html || escapeHtml(side.text));
    return `<div class="diffCell ${diffSideClass(side.type)}"><span class="diffNo">${noText}</span><span class="diffTxt">${textHtml}</span></div>`;
  }

  function renderDiffTable(rows){
    const html = [
      '<table class="diffSplit" aria-label="เทียบซ้ายขวา Config ต้นฉบับ และ Config หลังแก้ไข">',
      "<thead><tr><th>Config ต้นฉบับ</th><th>Config หลังแก้ไข</th></tr></thead>",
      "<tbody>",
    ];

    for(const row of rows){
      if(row.kind === "skip"){
        html.push(`<tr class="diffSkipRow"><td colspan="2">… (ข้าม ${row.count} บรรทัดที่เหมือนกัน) …</td></tr>`);
        continue;
      }
      html.push("<tr>");
      html.push(`<td>${renderDiffSideCell(row.left)}</td>`);
      html.push(`<td>${renderDiffSideCell(row.right)}</td>`);
      html.push("</tr>");
    }

    html.push("</tbody></table>");
    return html.join("");
  }

  function renderDiff(edits, onlyChanges){
    let add=0, del=0;
    for(const e of edits){
      if(e.type==="insert") add++;
      else if(e.type==="delete") del++;
    }
    diffSummary.textContent = `+${add} / -${del}`;

    if(add === 0 && del === 0){
      diffBox.innerHTML = "ไม่มีความต่าง";
      return;
    }

    let rows = buildSideBySideRows(edits, onlyChanges);
    rows = reconcileDeleteInsertRows(rows);
    diffBox.innerHTML = rows.length ? renderDiffTable(rows) : "ไม่มีความต่าง";
  }

  function refreshDiff(){
    if(!lastResult) return;
    const inLines = (lastResult.input || "").split(/\r?\n/);
    const outLines = (lastResult.output || "").split(/\r?\n/);
    const edits = myersDiffLines(inLines, outLines);
    renderDiff(edits, !!optOnlyChanges.checked);
  }

  // ---------- Download buttons ----------
  function clearDownloads(){
    downloadLinks.innerHTML = "";
  }

  function createDownloadButton(filename, content){
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([content], { type:"text/plain;charset=utf-8" }));
    a.download = filename;
    a.className = "btnLink";
    a.innerHTML = `<span class="btnIcon">⬇</span> <span>${escapeHtml(filename)}</span>`;
    a.addEventListener("click", () => setTimeout(()=>URL.revokeObjectURL(a.href), 1500));
    downloadLinks.appendChild(a);
  }

  function buildButtonsForSingleResult(result, index, usedNames){
    const fileName = result.fileName || `file_${index+1}.log`;
    const sourceStem = sanitizeFilenamePart(splitFilename(fileName).stem, `file_${index+1}`);
    const serials = (result.serials || []).map((s) => sanitizeFilenamePart(s, sourceStem));
    if(serials.length === 0){
      createDownloadButton(makeUniqueFilename(`${sourceStem}_converted.log`, usedNames), result.output);
      return 1;
    }

    let count = 0;
    for(const serial of serials){
      createDownloadButton(makeUniqueFilename(`${serial}.log`, usedNames), result.output);
      count++;
    }
    return count;
  }

  function buildDownloadButtons(){
    clearDownloads();
    const results = convertedResults.length
      ? convertedResults
      : (lastResult ? [{ fileName:"converted.log", output:lastResult.output, serials:lastResult.serials || [] }] : []);
    if(results.length === 0) return 0;

    const usedNames = new Set();
    let count = 0;
    results.forEach((result, index) => {
      count += buildButtonsForSingleResult(result, index, usedNames);
    });
    return count;
  }

  const relevantCmdRegex = {
    showClock: /#\s*(?:sh|sho|show)\s+cl(?:o(?:c(?:k)?)?)?\b/i,
    showVersion: /#\s*(?:sh|sho|show)\s+ver(?:s(?:i(?:o(?:n)?)?)?)?\b/i,
    showRun: /#\s*(?:sh|sho|show)\s+run(?:n(?:i(?:n(?:g(?:-c(?:o(?:n(?:f(?:i(?:g)?)?)?)?)?)?)?)?)?)?\b/i,
    showLog: /#\s*(?:sh|sho|show)\s+lo(?:g(?:g(?:i(?:n(?:g)?)?)?)?)?\b/i,
    showEnvAll: /#\s*(?:sh|sho|show)\s+en(?:v(?:i(?:r(?:o(?:n(?:m(?:e(?:n(?:t)?)?)?)?)?)?)?)?)\s+all\b/i,
    showIntCrc: /#\s*(?:sh|sho|show)\s+int(?:e(?:r(?:f(?:a(?:c(?:e(?:s)?)?)?)?)?)?)?\s*\|\s*i(?:n(?:c(?:l(?:u(?:d(?:e)?)?)?)?)?)?\s+crc\b/i,
  };

  function detectRelevantCommand(line){
    const txt = String(line || "");
    const promptMatch = txt.match(/^\s*[^\s#]+#(.*)$/);
    if(!promptMatch) return null;
    const display = String(promptMatch[1] || "")
      .trim()
      .replace(/\s+/g, " ")
      .replace(/\s*\|\s*/g, " | ");
    if(!display) return null;

    const probe = `# ${display}`;
    if(relevantCmdRegex.showClock.test(probe)) return { kind: "show clock", display };
    if(relevantCmdRegex.showVersion.test(probe)) return { kind: "show version", display };
    if(relevantCmdRegex.showRun.test(probe)) return { kind: "show run", display };
    if(relevantCmdRegex.showLog.test(probe)) return { kind: "show log", display };
    if(relevantCmdRegex.showEnvAll.test(probe)) return { kind: "show env all", display };
    if(relevantCmdRegex.showIntCrc.test(probe)) return { kind: "show interface | I CRC", display };
    return { kind: "unknown command", display };
  }

  function findTooOldShowClockIssue(text){
    const lines = String(text || "").split(/\r?\n/);
    const oldestAllowed = new Date();
    oldestAllowed.setUTCFullYear(oldestAllowed.getUTCFullYear() - 1);

    for(let i=0;i<lines.length;i++){
      if(!relevantCmdRegex.showClock.test(lines[i])) continue;
      let j = i + 1;
      while(j < lines.length && lines[j].trim() === "") j++;
      if(j >= lines.length) continue;
      const parsed = parseTimeLine(lines[j]);
      if(!parsed) continue;
      if(parsed.dt.getTime() < oldestAllowed.getTime()){
        return {
          lineNo: j + 1,
          foundDate: parsed.dt.toISOString().slice(0, 10),
          oldestAllowedDate: oldestAllowed.toISOString().slice(0, 10),
        };
      }
    }
    return null;
  }

  function formatProblemFileList(entries, maxNames=6){
    const names = (Array.isArray(entries) ? entries : [])
      .map((entry) => String(entry?.name || "").trim())
      .filter(Boolean);
    if(names.length <= maxNames) return names.join(", ");
    return `${names.slice(0, maxNames).join(", ")}, ...`;
  }

  function isSameOriginMessage(event){
    const currentOrigin = window.location.origin;
    if(currentOrigin === "null"){
      return event.origin === "null" || event.origin === "";
    }
    return event.origin === currentOrigin;
  }

  function normalizeImportedFiles(payload){
    if(!Array.isArray(payload?.files)) return [];
    return payload.files.map((entry) => ({
      name: String(entry?.name || "problem.log"),
      text: String(entry?.text || ""),
    }));
  }

  function applyImportedProblemFiles(payload){
    const files = normalizeImportedFiles(payload);
    if(files.length === 0) return false;
    fileInput.value = "";
    setSelectedInputs(files);
    if(payload?.autoCustomMode){
      enableCustomDateMode();
    }
    metaBox.textContent = "ในตอนนี้ถ้าคุณกด Convert จะได้เป็นวัน/เดือน/ปี ของปัจจุบัน\nคุณสามารถเลื่อนขึ้นไปข้างบนเพื่อเปลี่ยน วัน/เดือน/ปี ตามที่คุณต้องการได้";
    setStatus(`รับไฟล์มีปัญหา ${files.length} ไฟล์จากอีกแท็บแล้ว`);
    return true;
  }

  function enableCustomDateMode(){
    if(!optCustom.checked){
      optCustom.checked = true;
      optCustom.dispatchEvent(new Event("change"));
      return;
    }
    if(!pickDate.value){
      setDateToToday();
    } else {
      syncNativeDateFromText();
    }
    clampEnd();
  }

  function openProblemFilesInNewTab(problemEntries){
    if(!Array.isArray(problemEntries) || problemEntries.length === 0) return;
    const payload = {
      files: problemEntries.map((entry) => ({
        name: String(entry?.name || "problem.log"),
        text: String(entry?.text || ""),
      })),
      autoCustomMode: true,
    };
    const nonce = `${Date.now()}_${Math.random().toString(16).slice(2)}`;
    try{
      const nextTab = window.open("/cli", "_blank");
      if(!nextTab){
        throw new Error("popup blocked");
      }
      const targetOrigin = window.location.origin === "null" ? "*" : window.location.origin;
      let done = false;
      let sendTimer = null;
      let timeoutId = null;

      const cleanup = () => {
        if(sendTimer !== null){
          window.clearInterval(sendTimer);
          sendTimer = null;
        }
        if(timeoutId !== null){
          window.clearTimeout(timeoutId);
          timeoutId = null;
        }
        window.removeEventListener("message", onAckMessage);
      };

      const onAckMessage = (event) => {
        if(!isSameOriginMessage(event)) return;
        const data = event.data || {};
        if(data.type !== OLD_CLOCK_TRANSFER_ACK || data.nonce !== nonce) return;
        done = true;
        cleanup();
        removeProblemFilesFromCurrentTab(problemEntries);
      };
      window.addEventListener("message", onAckMessage);

      const sendPayload = () => {
        try{
          nextTab.postMessage(
            {
              type: OLD_CLOCK_TRANSFER_MESSAGE,
              nonce,
              payload,
            },
            targetOrigin
          );
        }catch(err){}
      };

      sendPayload();
      sendTimer = window.setInterval(sendPayload, 250);
      timeoutId = window.setTimeout(() => {
        if(done) return;
        cleanup();
        try{
          const storageKey = `${OLD_CLOCK_TRANSFER_PREFIX}${Date.now()}_${Math.random().toString(16).slice(2)}`;
          localStorage.setItem(storageKey, JSON.stringify(payload));
          const targetUrl = `/cli?${OLD_CLOCK_TRANSFER_QUERY}=${encodeURIComponent(storageKey)}`;
          nextTab.location.href = targetUrl;
          removeProblemFilesFromCurrentTab(problemEntries);
        }catch(err){
          setStatus("ย้ายไฟล์ไปแท็บใหม่ไม่สำเร็จ", false);
        }
      }, 6000);
    }catch(err){
      setStatus("ย้ายไฟล์ไปแท็บใหม่ไม่สำเร็จ", false);
    }
  }

  function removeProblemFilesFromCurrentTab(problemEntries){
    const indices = Array.from(
      new Set(
        (Array.isArray(problemEntries) ? problemEntries : [])
          .map((entry) => Number(entry?.sourceIndex))
          .filter((idx) => Number.isInteger(idx) && idx >= 0 && idx < selectedInputs.length)
      )
    ).sort((a, b) => a - b);

    if(indices.length === 0){
      metaBox.textContent = "ไฟล์ที่มีปัญหาถูกย้ายแล้วให้กดปุ่ม convert อีกที เพื่อความถูกต้อง";
      setStatus("ย้ายไฟล์ที่มีปัญหาไปแท็บใหม่แล้ว");
      return;
    }

    const blocked = new Set(indices);
    const remaining = selectedInputs.filter((_, idx) => !blocked.has(idx));
    setSelectedInputs(remaining);
    metaBox.textContent = "ไฟล์ที่มีปัญหาถูกย้ายแล้วให้กดปุ่ม convert อีกที เพื่อความถูกต้อง";
    if(remaining.length > 0){
      setStatus(`ย้ายไฟล์ที่มีปัญหา ${indices.length} ไฟล์แล้ว เหลือ ${remaining.length} ไฟล์`);
    } else {
      setStatus(`ย้ายไฟล์ที่มีปัญหา ${indices.length} ไฟล์แล้ว`);
    }
  }

  function showOldClockConvertError(problemEntries){
    const count = Array.isArray(problemEntries) ? problemEntries.length : 0;
    const namesText = formatProblemFileList(problemEntries);
    const message = `Convert error: มีทั้งหมด ${count} ไฟล์ (${namesText}) เวลา show clock เก่าเกิน ให้กดปุ่มด้านขวาเพื่อย้ายไฟล์ที่มีปัญหาไปอีกแท็บ`;
    metaBox.innerHTML = (
      `<div class="meta-error-row">`
      + `<span class="meta-error-text">${escapeHtml(message)}</span>`
      + `<button type="button" class="meta-move-btn" data-move-problem>ย้ายไฟล์มีปัญหาไปแท็บใหม่</button>`
      + `</div>`
    );
    const moveButton = metaBox.querySelector("[data-move-problem]");
    moveButton?.addEventListener("click", () => openProblemFilesInNewTab(problemEntries));
  }

  function importProblemFilesFromQuery(){
    const params = new URLSearchParams(window.location.search);
    const transferKey = params.get(OLD_CLOCK_TRANSFER_QUERY);
    if(!transferKey) return;

    try{
      const raw = localStorage.getItem(transferKey);
      if(!raw) return;
      const payload = JSON.parse(raw);
      const applied = applyImportedProblemFiles(payload);
      if(!applied){
        setStatus("โหลดไฟล์จากอีกแท็บไม่สำเร็จ", false);
      }
    }catch(err){
      setStatus("โหลดไฟล์จากอีกแท็บไม่สำเร็จ", false);
    }finally{
      localStorage.removeItem(transferKey);
      params.delete(OLD_CLOCK_TRANSFER_QUERY);
      const query = params.toString();
      const cleanUrl = `${window.location.pathname}${query ? `?${query}` : ""}`;
      window.history.replaceState({}, "", cleanUrl);
    }
  }

  function importProblemFilesFromMessage(){
    window.addEventListener("message", (event) => {
      if(!isSameOriginMessage(event)) return;
      const data = event.data || {};
      if(data.type !== OLD_CLOCK_TRANSFER_MESSAGE) return;

      const applied = applyImportedProblemFiles(data.payload || {});
      try{
        const replyOrigin = event.origin && event.origin !== "null" ? event.origin : "*";
        if(event.source && typeof event.source.postMessage === "function"){
          event.source.postMessage(
            {
              type: OLD_CLOCK_TRANSFER_ACK,
              nonce: data.nonce || "",
            },
            replyOrigin
          );
        }
      }catch(err){}

      if(!applied){
        setStatus("โหลดไฟล์จากอีกแท็บไม่สำเร็จ", false);
      }
    });
  }

  function findMissingCommands(expected, observed, observedDisplay=null){
    const exp = Array.isArray(expected) ? expected : [];
    const obs = Array.isArray(observed) ? observed : [];
    const obsDisplay = Array.isArray(observedDisplay) ? observedDisplay : [];
    const n = exp.length;
    const m = obs.length;
    const countOccurrence = (arr, idx) => {
      let count = 0;
      for(let x=0;x<=idx;x++){
        if(arr[x] === arr[idx]) count++;
      }
      return count;
    };
    const labelAt = (arr, idx) => `${arr[idx]} (ครั้งที่ ${countOccurrence(arr, idx)})`;
    const dp = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));

    for(let i=1;i<=n;i++){
      for(let j=1;j<=m;j++){
        if(exp[i - 1] === obs[j - 1]) dp[i][j] = dp[i - 1][j - 1] + 1;
        else dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }

    const matchedExpected = new Set();
    const matchedObserved = new Set();
    let i = n;
    let j = m;
    while(i > 0 && j > 0){
      if(exp[i - 1] === obs[j - 1]){
        matchedExpected.add(i - 1);
        matchedObserved.add(j - 1);
        i--;
        j--;
      } else if(dp[i - 1][j] >= dp[i][j - 1]){
        i--;
      } else {
        j--;
      }
    }

    const missing = [];
    const missingDetails = [];
    const extra = [];
    const extraDetails = [];
    for(let k=0;k<n;k++){
      if(!matchedExpected.has(k)){
        missing.push(exp[k]);
        let beforeIdx = k - 1;
        while(beforeIdx >= 0 && !matchedExpected.has(beforeIdx)) beforeIdx--;
        let afterIdx = k + 1;
        while(afterIdx < n && !matchedExpected.has(afterIdx)) afterIdx++;

        let location = "ไม่สามารถระบุตำแหน่งอ้างอิงได้";
        if(beforeIdx >= 0 && afterIdx < n){
          location = `หลัง ${labelAt(exp, beforeIdx)} ก่อน ${labelAt(exp, afterIdx)}`;
        } else if(beforeIdx >= 0){
          location = `หลัง ${labelAt(exp, beforeIdx)}`;
        } else if(afterIdx < n){
          location = `ก่อน ${labelAt(exp, afterIdx)}`;
        }
        missingDetails.push({
          command: exp[k],
          occurrence: countOccurrence(exp, k),
          location,
        });
      }
    }

    for(let k=0;k<m;k++){
      if(!matchedObserved.has(k)){
        const display = String(obsDisplay[k] || obs[k] || "").trim() || "-";
        extra.push(display);
        let beforeIdx = k - 1;
        while(beforeIdx >= 0 && !matchedObserved.has(beforeIdx)) beforeIdx--;
        let afterIdx = k + 1;
        while(afterIdx < m && !matchedObserved.has(afterIdx)) afterIdx++;

        let location = "ไม่สามารถระบุตำแหน่งอ้างอิงได้";
        if(beforeIdx >= 0 && afterIdx < m){
          location = `หลัง ${labelAt(obs, beforeIdx)} ก่อน ${labelAt(obs, afterIdx)}`;
        } else if(beforeIdx >= 0){
          location = `หลัง ${labelAt(obs, beforeIdx)}`;
        } else if(afterIdx < m){
          location = `ก่อน ${labelAt(obs, afterIdx)}`;
        }
        extraDetails.push({
          command: obs[k],
          display,
          occurrence: countOccurrence(obs, k),
          position: k + 1,
          location,
        });
      }
    }
    return { missing, missingDetails, extra, extraDetails };
  }

  // ---------- Conversion ----------
  function transform(text){
    const expectedCommandOrder = [
      "show clock",
      "show version",
      "show run",
      "show log",
      "show env all",
      "show clock",
      "show interface | I CRC",
      "show clock",
      "show interface | I CRC",
    ];
    const meta = {
      removedClear: 0,
      removedClearLines: [],
      changedToZero: {},   // field -> count
      changedFrom: {},     // field -> set of original numbers
      zeroChangedLines: [],
      commandOrder: {
        expected: expectedCommandOrder,
        observed: [],
        observedDisplay: [],
        missing: [],
        missingDetails: [],
        extra: [],
        extraDetails: [],
        pass: false,
      },
      clock: { found:0, adjusted:false, reason:"", rawDelta12Sec:null, newDelta12Sec:null, delta12Randomized:false, rawDelta23Sec:null, newDelta23Sec:null, usedCustom:false, changed1:false, changed2:false, changed3:false, details: [] }
    };

    // init meta maps
    const fields = ["input errors","CRC","frame","overrun","ignored","abort"];
    for(const f of fields){
      meta.changedToZero[f] = 0;
      meta.changedFrom[f] = new Set();
    }

    // 1) remove lines containing word 'clear'
    const rawLines = text.split(/\r?\n/);
    const kept = [];
    for(const ln of rawLines){
      if(/\bclear\b/i.test(ln)){
        meta.removedClear++;
        meta.removedClearLines.push(ln);
        continue;
      }
      kept.push(ln);
    }

    const observedOrder = [];
    const observedDisplay = [];
    for(const ln of kept){
      const cmd = detectRelevantCommand(ln);
      if(cmd){
        observedOrder.push(cmd.kind);
        observedDisplay.push(cmd.display);
      }
    }
    meta.commandOrder.observed = observedOrder;
    meta.commandOrder.observedDisplay = observedDisplay;
    meta.commandOrder.pass = (
      observedOrder.length === expectedCommandOrder.length
      && observedOrder.every((cmd, idx) => cmd === expectedCommandOrder[idx])
    );
    const missingResult = findMissingCommands(expectedCommandOrder, observedOrder, observedDisplay);
    meta.commandOrder.missing = missingResult.missing;
    meta.commandOrder.missingDetails = missingResult.missingDetails;
    meta.commandOrder.extra = missingResult.extra;
    meta.commandOrder.extraDetails = missingResult.extraDetails;

    let t = kept.join("\n");

    // 2) force related counters to 0 anywhere they appear (matches "<value> field")
    const reps = [
      { key:"input errors", re:/^(\s*)([^,\s]+)\s+input\s+errors\b/gi, withComma:false },
      { key:"CRC",         re:/(,\s*)([^,\s]+)\s+CRC\b/gi, withComma:true },
      { key:"frame",       re:/(,\s*)([^,\s]+)\s+frame\b/gi, withComma:true },
      { key:"overrun",     re:/(,\s*)([^,\s]+)\s+overrun\b/gi, withComma:true },
      { key:"ignored",     re:/(,\s*)([^,\s]+)\s+ignored\b/gi, withComma:true },
      { key:"abort",       re:/(,\s*)([^,\s]+)\s+abort\b/gi, withComma:true },
    ];
    const zeroLines = t.split("\n");
    for(let i=0;i<zeroLines.length;i++){
      const beforeLine = zeroLines[i];
      let afterLine = beforeLine;
      for(const r of reps){
        r.re.lastIndex = 0;
        afterLine = afterLine.replace(r.re, (...args) => {
          const prefix = String(args[1] || "");
          const num = String(args[2] || "");
          const raw = String(num || "").trim();
          const numeric = Number(raw);
          const isZeroLike = raw === "0" || (!Number.isNaN(numeric) && numeric === 0);
          if(!isZeroLike){
            meta.changedToZero[r.key]++;
            meta.changedFrom[r.key].add(raw);
          }
          if(r.withComma){
            return `${prefix}0 ${r.key}`;
          }
          return `${prefix}0 ${r.key}`;
        });
      }
      if(afterLine !== beforeLine){
        meta.zeroChangedLines.push({ before: beforeLine, after: afterLine });
      }
      zeroLines[i] = afterLine;
    }
    t = zeroLines.join("\n");

    // 3) adjust show clock blocks in first 3 occurrences
    const lines = t.split("\n");
    const cmdRe = relevantCmdRegex.showClock;

    const blocks = []; // {cmdIdx, timeIdx, parsed}
    for(let i=0;i<lines.length;i++){
      if(cmdRe.test(lines[i])){
        let j=i+1;
        while(j < lines.length && lines[j].trim()==="") j++;
        if(j < lines.length){
          const parsed = parseTimeLine(lines[j]);
          if(parsed) blocks.push({ cmdIdx:i, timeIdx:j, parsed });
        }
      }
    }
    meta.clock.found = blocks.length;

    if(blocks.length >= 3){
      const b1 = blocks[0], b2 = blocks[1], b3 = blocks[2];
      const dt1_raw = b1.parsed.dt;
      const dt2_raw = b2.parsed.dt;
      const dt3_raw = b3.parsed.dt;
      const clockOriginal = [
        lines[b1.timeIdx],
        lines[b2.timeIdx],
        lines[b3.timeIdx],
      ];

      const dayMs = 24*3600*1000;
      const maxDelta12Sec = 6*60;
      const randomDelta12MinSec = 30;
      const randomDelta12MaxSec = 5*60;
      let delta12ms = dt2_raw.getTime() - dt1_raw.getTime();
      // Keep only the time-of-day gap to avoid carrying over year/day jumps from raw logs.
      delta12ms = ((delta12ms % dayMs) + dayMs) % dayMs;
      meta.clock.rawDelta12Sec = Math.round(delta12ms/1000);
      meta.clock.rawDelta23Sec = Math.round((dt3_raw.getTime() - dt2_raw.getTime())/1000);
      let delta12ForCustomMs = delta12ms;
      if(meta.clock.rawDelta12Sec > maxDelta12Sec){
        delta12ForCustomMs = randInt(randomDelta12MinSec, randomDelta12MaxSec) * 1000;
        meta.clock.delta12Randomized = true;
      }
      meta.clock.newDelta12Sec = Math.round(delta12ForCustomMs/1000);

      let dt2_new = new Date(dt2_raw.getTime());
      let dt1_new = new Date(dt1_raw.getTime());

      meta.clock.usedCustom = !!optCustom.checked;
      if(!optCustom.checked){
        dt1_new = new Date(dt2_new.getTime() - delta12ForCustomMs);
        meta.clock.changed1 = (dt1_new.getTime() !== dt1_raw.getTime());
      }

      if(optCustom.checked){
        const dateVal = pickDate.value;
        if(!dateVal){
          meta.clock.adjusted = false;
          meta.clock.reason = "เปิดโหมดกำหนดเอง แต่ยังไม่ได้เลือกวันที่";
        } else {
          const parsedDate = parseDateInputToYmd(dateVal);
          if(!parsedDate){
            meta.clock.adjusted = false;
            meta.clock.reason = "รูปแบบวันที่ไม่ถูกต้อง (ใช้ วัน/เดือน/ปี เช่น 25/02/2026)";
          } else {
            const { year: Y, month: M, day: D } = parsedDate;
            const s = getSelectTime(timeStartHour, timeStartMinute, timeStartSecond);
            const e = getSelectTime(timeEndHour, timeEndMinute, timeEndSecond);

            // enforce end >= start
            let startSec = toSec(s);
            let endSec = toSec(e);
            if(endSec < startSec) endSec = startSec;

            // pick dt2 within range
            const sec2 = randInt(startSec, endSec);
            const ms2 = randInt(0,999);
            dt2_new = new Date(Date.UTC(Y, M-1, D, 0,0,0,0) + sec2*1000 + ms2);

            // preserve Δ(1→2), unless raw gap is > 6 minutes then randomize to 30..300 sec
            dt1_new = new Date(dt2_new.getTime() - delta12ForCustomMs);

            meta.clock.changed2 = (dt2_new.getTime() !== dt2_raw.getTime());
            meta.clock.changed1 = (dt1_new.getTime() !== dt1_raw.getTime());
          }
        }
      }

      // Always enforce Δ(2→3) in [420..450] seconds
      const delta23sec = randInt(420, 450);
      const dt3_new = new Date(dt2_new.getTime() + delta23sec*1000);

      meta.clock.newDelta23Sec = delta23sec;
      meta.clock.changed3 = (dt3_new.getTime() !== dt3_raw.getTime());

      // Write back lines
      if(meta.clock.changed1){
        lines[b1.timeIdx] = formatTimeLine(dt1_new, { hasMs: b1.parsed.hasMs, tz: b1.parsed.tz, prefix: b1.parsed.prefix });
      }
      if(optCustom.checked && meta.clock.changed2){
        lines[b2.timeIdx] = formatTimeLine(dt2_new, { hasMs: b2.parsed.hasMs, tz: b2.parsed.tz, prefix: b2.parsed.prefix });
      }
      // Always rewrite #3 (so it matches rule), keeping prefix and tz of its own line
      lines[b3.timeIdx] = formatTimeLine(dt3_new, { hasMs: b3.parsed.hasMs, tz: b3.parsed.tz, prefix: b3.parsed.prefix });

      if(meta.clock.reason === ""){
        meta.clock.adjusted = true;
        const delta12Summary = meta.clock.rawDelta12Sec === meta.clock.newDelta12Sec
          ? `raw Δ(1→2)=${meta.clock.rawDelta12Sec}s`
          : `raw Δ(1→2)=${meta.clock.rawDelta12Sec}s → new Δ(1→2)=${meta.clock.newDelta12Sec}s`;
        meta.clock.reason = `${delta12Summary} | raw Δ(2→3)=${meta.clock.rawDelta23Sec}s → new Δ(2→3)=${delta23sec}s`;
      }
      meta.clock.details = [
        { label: "clock #1", before: clockOriginal[0], after: lines[b1.timeIdx] },
        { label: "clock #2", before: clockOriginal[1], after: lines[b2.timeIdx] },
        { label: "clock #3", before: clockOriginal[2], after: lines[b3.timeIdx] },
      ];
    } else {
      meta.clock.adjusted = false;
      meta.clock.reason = `เจอ show clock แค่ ${blocks.length} ชุด (ต้องมี 3 เท่านั้น)`;
      meta.clock.details = [];
    }

    const output = lines.join("\n");
    const serials = findUniqueSerials(output);
    return { output, meta, serials };
  }

  function renderCommandFlow(commands, kinds=null){
    if(!Array.isArray(commands) || commands.length === 0){
      return "-";
    }
    return commands.map((command, index) => {
      const safe = escapeHtml(String(command || "").trim());
      let cmdClass = "meta-flow-cmd";
      if(Array.isArray(kinds)){
        const kind = String(kinds[index] || "");
        if(kind === "unknown command") cmdClass += " meta-flow-cmd-bad";
        else if(kind) cmdClass += " meta-flow-cmd-ok";
      }
      if(index === 0){
        return `<span class="${cmdClass}">${safe}</span>`;
      }
      return `<span class="meta-flow-sep"> -&gt; </span><wbr><span class="${cmdClass}">${safe}</span>`;
    }).join("");
  }

  function escapeRegex(s){
    return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  function highlightCommandsInText(text, commands, cssClass){
    const source = String(text || "");
    const list = Array.from(
      new Set(
        (Array.isArray(commands) ? commands : [])
          .map((cmd) => String(cmd || "").trim())
          .filter(Boolean)
      )
    ).sort((a, b) => b.length - a.length);
    if(list.length === 0){
      return escapeHtml(source);
    }
    const re = new RegExp(list.map((cmd) => escapeRegex(cmd)).join("|"), "gi");
    let html = "";
    let last = 0;
    let match;
    while((match = re.exec(source)) !== null){
      html += escapeHtml(source.slice(last, match.index));
      html += `<span class="${cssClass}">${escapeHtml(match[0])}</span>`;
      last = match.index + match[0].length;
    }
    html += escapeHtml(source.slice(last));
    return html;
  }

  function metaHtml(meta, serials, headText="", detailsOpen=false, summaryOverride=null){
    const detailsRows = [];
    const head = String(headText || "").trim();
    const listHtml = (items) => `<ul class="meta-list">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;
    const oldClockPrecheck = meta?.oldClockPrecheck || null;

    if(head){
      detailsRows.push(`<div class="meta-file">${escapeHtml(head)}</div>`);
    }
    if(oldClockPrecheck?.tooOld){
      detailsRows.push(`<div class="meta-section-title">แจ้งเตือนการปรับวันที่อัตโนมัติ</div>`);
      detailsRows.push(
        `<div class="meta-flow">${escapeHtml(buildAutoFixNotice(oldClockPrecheck))}</div>`
      );
    }

    detailsRows.push(
      `<div class="meta-section-title">ตรวจลำดับคำสั่ง: `
      + `<span class="${meta.commandOrder.pass ? "meta-value-ok" : "meta-value-bad"}">${meta.commandOrder.pass ? "ผ่าน" : "ไม่ผ่าน"}</span></div>`
    );
    detailsRows.push(`<div class="meta-kv"><span class="meta-label">รูปแบบที่ต้องการ:</span></div>`);
    detailsRows.push(`<div class="meta-flow">${renderCommandFlow(meta.commandOrder.expected)}</div>`);
    detailsRows.push(`<div class="meta-kv"><span class="meta-label">รูปแบบที่พบ:</span></div>`);
    detailsRows.push(
      `<div class="meta-flow">${meta.commandOrder.observedDisplay.length ? renderCommandFlow(meta.commandOrder.observedDisplay, meta.commandOrder.observed) : "-"}</div>`
    );
    if(meta.commandOrder.missing.length){
      const missingMessages = meta.commandOrder.missingDetails.length
        ? meta.commandOrder.missingDetails.map((entry) =>
            `ต้องมี!! ${entry.command} (ครั้งที่ ${entry.occurrence}) คำสั่งนี้หายไปช่วง: ${entry.location}`
          )
        : meta.commandOrder.missing.map((entry) => `ต้องมี!! ${entry} คำสั่งนี้หายไป`);
      detailsRows.push(`<div class="meta-kv"><span class="meta-label meta-label-missing">คำสั่งที่หายไป:</span></div>`);
      detailsRows.push(
        `<div class="meta-flow meta-flow-missing">${missingMessages.map((msg) => highlightCommandsInText(msg, meta.commandOrder.expected, "meta-flow-cmd meta-flow-cmd-bad")).join("<br>")}</div>`
      );
    }
    if(meta.commandOrder.extra.length){
      const extraMessagesHtml = meta.commandOrder.extraDetails.length
        ? meta.commandOrder.extraDetails.map((entry) => {
            const cmdTag = `<span class="meta-flow-cmd meta-flow-cmd-bad">${escapeHtml(entry.display)}</span>`;
            const locationHtml = highlightCommandsInText(entry.location, meta.commandOrder.expected, "meta-flow-cmd meta-flow-cmd-bad");
            return `ต้องลบ!! ${cmdTag} (ลำดับที่ ${entry.position} ในรูปแบบที่พบ) เกินช่วง: ${locationHtml}`;
          })
        : meta.commandOrder.extra.map((entry, idx) => {
            const cmdTag = `<span class="meta-flow-cmd meta-flow-cmd-bad">${escapeHtml(entry)}</span>`;
            return `ต้องลบ!! ${cmdTag} (ลำดับที่ ${idx + 1} ในรูปแบบที่พบ)`;
          });
      detailsRows.push(`<div class="meta-kv"><span class="meta-label meta-label-extra">คำสั่งที่เกินมา:</span></div>`);
      detailsRows.push(
        `<div class="meta-flow meta-flow-extra">${extraMessagesHtml.join("<br>")}</div>`
      );
    }

    detailsRows.push(`<div class="meta-section-title">การลบ clear</div>`);
    detailsRows.push(`<div class="meta-kv"><span class="meta-label">ลบคำว่า clear:</span> <span class="meta-value">${meta.removedClear} บรรทัด</span></div>`);
    if(meta.removedClearLines.length){
      detailsRows.push(listHtml(meta.removedClearLines));
    }

    const totalZero = Object.values(meta.changedToZero).reduce((a,b)=>a+b,0);
    detailsRows.push(`<div class="meta-section-title">การปรับตัวเลขเป็น 0</div>`);
    detailsRows.push(`<div class="meta-kv"><span class="meta-label">ปรับตัวเลขเป็น 0:</span> <span class="meta-value">${totalZero} จุด</span></div>`);
    if(meta.zeroChangedLines.length){
      detailsRows.push(listHtml(meta.zeroChangedLines.map((entry) => `${entry.before} -> ${entry.after}`)));
    }
    const zeroSummary = [];
    for(const k of Object.keys(meta.changedToZero)){
      if(meta.changedToZero[k] > 0){
        const orig = Array.from(meta.changedFrom[k]).sort((a,b) => {
          const na = Number(a);
          const nb = Number(b);
          const aIsNum = !Number.isNaN(na);
          const bIsNum = !Number.isNaN(nb);
          if(aIsNum && bIsNum) return na - nb;
          if(aIsNum) return -1;
          if(bIsNum) return 1;
          return String(a).localeCompare(String(b));
        }).join(", ");
        zeroSummary.push(`${k}: ${orig} → 0 (รวม ${meta.changedToZero[k]} ครั้ง)`);
      }
    }
    if(zeroSummary.length){
      detailsRows.push(listHtml(zeroSummary));
    }

    const clockFound = Number(meta.clock?.found || 0);
    const clockPass = clockFound === 3;
    const clockStatusText = clockPass ? "ผ่าน" : `ไม่ผ่าน (เพราะเจอเวลา: ${clockFound} ตัว)`;
    detailsRows.push(
      `<div class="meta-section-title">รายละเอียดเวลา clock ที่เปลี่ยน: `
      + `<span class="${clockPass ? "meta-value-ok" : "meta-value-bad"}">${clockStatusText}</span></div>`
    );
    if(meta.clock.details.length){
      detailsRows.push(listHtml(meta.clock.details.map((d) => `${d.label}: ${d.before} -> ${d.after}`)));
    } else {
      detailsRows.push(`<div class="meta-flow">ไม่พบ show clock ครบ 3 ชุด</div>`);
    }
    if(meta.clock.reason){
      detailsRows.push(`<div class="meta-kv"><span class="meta-label">หมายเหตุ:</span> <span class="meta-value">${escapeHtml(meta.clock.reason)}</span></div>`);
    }

    detailsRows.push(`<div class="meta-section-title">System serial number</div>`);
    if(serials.length){
      detailsRows.push(`<div class="meta-flow">${escapeHtml(serials.join(", "))}</div>`);
      detailsRows.push(`<div class="meta-kv"><span class="meta-label">Download:</span> <span class="meta-value">จะมีปุ่มให้กดโหลดทีละไฟล์ด้านล่าง</span></div>`);
    } else {
      detailsRows.push(`<div class="meta-flow">ไม่พบ (จะใช้ชื่อ converted.log)</div>`);
    }

    const summaryAction = `<button type="button" class="meta-more" data-meta-toggle>${detailsOpen ? "ซ่อนรายละเอียด" : "กดดูเพิ่มเติม"}</button>`;
    const detailsHiddenAttr = detailsOpen ? "" : " hidden";
    const statusClass = summaryOverride?.statusPass === false ? "meta-value-bad" : "meta-value-ok";
    const statusText = summaryOverride?.statusText || "สมบูรณ์";

    return (
      `<div class="meta-report">`
      + `<div class="meta-summary">`
      + `<div class="meta-kv"><span class="meta-label">สถานะ:</span> <span class="${statusClass}">${escapeHtml(statusText)}</span></div>`
      + summaryAction
      + `</div>`
      + `<div class="meta-details"${detailsHiddenAttr}>${detailsRows.join("")}</div>`
      + `</div>`
    );
  }

  function normalizePreviewName(name){
    const raw = String(name || "").trim();
    if(!raw) return "-";
    const noExt = raw.replace(/\.[^.]+$/, "");
    const noPrefix = noExt.replace(/^\d+\.\s*/, "");
    const compact = noPrefix.replace(/\s+rawfile$/i, "").trim();
    return compact || noPrefix || raw;
  }

  function isResultHealthy(result){
    const meta = result?.meta;
    if(!meta) return false;
    const orderPass = !!meta.commandOrder?.pass;
    const clockFound = Number(meta.clock?.found || 0);
    return orderPass && clockFound === 3;
  }

  function buildBatchSummary(){
    const activeResult = convertedResults[activeIndex] || convertedResults[0];
    const currentPass = isResultHealthy(activeResult);
    if(selectedInputs.length <= 1){
      if(currentPass){
        const oldClockHandled = !!activeResult?.meta?.oldClockPrecheck?.tooOld;
        return {
          statusPass: true,
          statusText: oldClockHandled
            ? "สมบูรณ์ (ระบบปรับวันที่ปัจจุบันให้อัตโนมัติ เพราะตรวจพบ show clock เก่าเกินเงื่อนไข)"
            : "สมบูรณ์",
        orderPass: true,
        orderText: "ผ่าน",
        showOrderSummary: true,
      };
    }
    return {
      statusPass: false,
      statusText: "ไม่สมบูรณ์",
      orderPass: false,
      orderText: "ไม่ผ่าน",
      showOrderSummary: false,
    };
  }

    const problemNames = convertedResults
      .filter((entry) => !isResultHealthy(entry))
      .map((entry) => normalizePreviewName(entry.fileName));
    const uniqueProblemNames = Array.from(new Set(problemNames));

    if(uniqueProblemNames.length === 0){
      return {
        statusPass: true,
        statusText: "สมบูรณ์ (ครบทุกตัว)",
        orderPass: true,
        orderText: "ผ่าน (ทุกตัว)",
        showOrderSummary: true,
      };
    }

    const listText = uniqueProblemNames.join(", ");
    return {
      statusPass: false,
      statusText: `ไม่สมบูรณ์ (ตัวที่มีปัญหา: ${listText})`,
      orderPass: false,
      orderText: "ไม่ผ่าน",
      showOrderSummary: false,
    };
  }

  // ---------- Time range UX ----------
  function updateTimeConstraints(){
    const s = getSelectTime(timeStartHour, timeStartMinute, timeStartSecond);
    const e = getSelectTime(timeEndHour, timeEndMinute, timeEndSecond);
    rangeHint.textContent = `ช่วงเวลา 24 ชม.: ${s} → ${e} (ถึงต้องไม่ก่อนเริ่ม)`;
  }
  function clampEnd(){
    const s = getSelectTime(timeStartHour, timeStartMinute, timeStartSecond);
    const e = getSelectTime(timeEndHour, timeEndMinute, timeEndSecond);
    if(toSec(e) < toSec(s)) setSelectTime(timeEndHour, timeEndMinute, timeEndSecond, s);
    updateTimeConstraints();
  }
  function clampStart(){
    const s = getSelectTime(timeStartHour, timeStartMinute, timeStartSecond);
    const e = getSelectTime(timeEndHour, timeEndMinute, timeEndSecond);
    if(toSec(s) > toSec(e)) setSelectTime(timeStartHour, timeStartMinute, timeStartSecond, e);
    updateTimeConstraints();
  }

  optCustom.addEventListener("change", () => {
    customRow.style.display = optCustom.checked ? "grid" : "none";
    if(optCustom.checked && !pickDate.value){
      setDateToToday();
    } else if(optCustom.checked){
      syncNativeDateFromText();
    }
    clampEnd();
  });
  for(const el of [timeStartHour, timeStartMinute, timeStartSecond]){
    el.addEventListener("change", () => { if(optCustom.checked) clampEnd(); });
  }
  for(const el of [timeEndHour, timeEndMinute, timeEndSecond]){
    el.addEventListener("change", () => { if(optCustom.checked) clampStart(); });
  }
  initTimeSelectors();
  bindDatePickerUi();
  updateTimeConstraints();

  // ---------- File read ----------
  function resetOutputViews(){
    lastResult = null;
    outBox.value = "";
    diffBox.textContent = "Not converted yet...";
    diffSummary.textContent = "No diff yet";
    clearDownloads();
    btnDownload.disabled = true;
    btnCopyOut.disabled = true;
  }

  function refreshPreviewSelector(){
    if(!previewSelect || !previewRow) return;
    previewSelect.innerHTML = "";
    selectedInputs.forEach((entry, idx) => {
      const opt = document.createElement("option");
      opt.value = String(idx);
      opt.textContent = `${idx + 1}. ${entry.name}`;
      previewSelect.appendChild(opt);
    });
    previewRow.style.display = selectedInputs.length > 1 ? "flex" : "none";
    if(activeIndex >= 0 && activeIndex < selectedInputs.length){
      previewSelect.value = String(activeIndex);
    }
  }

  function setConvertedView(result){
    lastResult = { input: result.input, output: result.output, meta: result.meta, serials: result.serials };
    outBox.value = result.output;
    const head = selectedInputs.length > 1 ? `File ${activeIndex + 1}/${selectedInputs.length}: ${result.fileName}\n` : "";
    const summaryInfo = buildBatchSummary();
    metaBox.innerHTML = metaHtml(result.meta, result.serials, head, metaDetailsPinned, summaryInfo);
    const moreButton = metaBox.querySelector("[data-meta-toggle]");
    const detailsBox = metaBox.querySelector(".meta-details");
    if(moreButton && detailsBox){
      moreButton.addEventListener("click", (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        const shouldOpen = detailsBox.hasAttribute("hidden");
        if(shouldOpen){
          detailsBox.removeAttribute("hidden");
          moreButton.textContent = "ซ่อนรายละเอียด";
          metaDetailsPinned = true;
        } else {
          detailsBox.setAttribute("hidden", "");
          moreButton.textContent = "กดดูเพิ่มเติม";
          metaDetailsPinned = false;
        }
      });
    }
    btnCopyOut.disabled = false;
    btnDownload.disabled = false;
    refreshDiff();
  }

  function renderActiveFileView(){
    if(selectedInputs.length === 0){
      activeIndex = -1;
      currentText = "";
      inpBox.value = "";
      resetOutputViews();
      metaBox.textContent = "Status: waiting for file selection.";
      updateLineCounts();
      refreshPreviewSelector();
      return;
    }

    if(activeIndex < 0 || activeIndex >= selectedInputs.length) activeIndex = 0;
    const active = selectedInputs[activeIndex];
    currentText = active.text || "";
    inpBox.value = currentText;

    const hasConverted = convertedResults.length === selectedInputs.length && convertedResults.length > 0;
    if(hasConverted){
      setConvertedView(convertedResults[activeIndex]);
    } else {
      resetOutputViews();
      metaBox.textContent = `Status: selected ${selectedInputs.length} file(s).\nPreview: ${active.name}\nClick Convert to process.`;
    }

    updateLineCounts();
    refreshPreviewSelector();
  }

  function setSelectedInputs(entries){
    selectedInputs = entries;
    convertedResults = [];
    activeIndex = entries.length ? 0 : -1;
    btnConvert.disabled = entries.length === 0;
    setFileSelectionUI(entries);
    renderActiveFileView();
  }

  previewSelect?.addEventListener("change", () => {
    const idx = Number(previewSelect.value);
    if(!Number.isInteger(idx)) return;
    if(idx < 0 || idx >= selectedInputs.length) return;
    activeIndex = idx;
    renderActiveFileView();
    const active = selectedInputs[activeIndex];
    if(active) setStatus(`Preview: ${active.name}`);
  });

  async function handleSelectedFiles(files){
    if(files.length === 0){
      setSelectedInputs([]);
      setStatus("No file selected");
      metaBox.textContent = "Status: waiting for file selection.";
      return;
    }
    try{
      const entries = await Promise.all(files.map(async (f) => ({ name: f.name, text: await f.text() })));
      setSelectedInputs(entries);
      setStatus(files.length > 1 ? `Selected ${files.length} files` : `Selected file: ${files[0].name}`);
    }catch(err){
      setSelectedInputs([]);
      setStatus("Read file failed", false);
      metaBox.textContent = "Error: " + (err && err.message ? err.message : String(err));
    }
  }

  fileInput.addEventListener("change", async () => {
    const files = Array.from(fileInput.files || []);
    await handleSelectedFiles(files);
  });

  filePickButton?.addEventListener("click", () => {
    fileInput.click();
  });

  fileDropZone?.addEventListener("click", (ev) => {
    if(ev.target instanceof HTMLElement && ev.target.closest(".pick-file-btn")) return;
    fileInput.click();
  });

  fileDropZone?.addEventListener("keydown", (ev) => {
    if(ev.key !== "Enter" && ev.key !== " ") return;
    ev.preventDefault();
    fileInput.click();
  });

  for(const dragEvent of ["dragenter", "dragover"]){
    fileDropZone?.addEventListener(dragEvent, (ev) => {
      ev.preventDefault();
      ev.stopPropagation();
      fileDropZone.classList.add("dragover");
    });
  }
  for(const dragEvent of ["dragleave", "drop"]){
    fileDropZone?.addEventListener(dragEvent, (ev) => {
      ev.preventDefault();
      ev.stopPropagation();
      fileDropZone.classList.remove("dragover");
    });
  }
  fileDropZone?.addEventListener("drop", async (ev) => {
    const files = Array.from(ev.dataTransfer?.files || []);
    await handleSelectedFiles(files);
  });

  btnSample?.addEventListener("click", () => {
    const sample =
`C3750-RDMO-DSW18#show clock
*08:12:21.902 UTC Tue Feb 3 2026

C3750-RDMO-DSW18#sh clock
 08:19:25.195 BKK Tue Feb 3 2026

C3750-RDMO-DSW18#show clock
*08:20:10.010 UTC Tue Feb 3 2026

Fa2/0/1  12 input errors, 3 CRC, 1 frame, 0 overrun, 9 ignored, 2 abort
clear mac address-table dynamic

System serial number : FOC1016Y0DF
System serial number : CAT1028RLCW
`;
    fileInput.value = "";
    setSelectedInputs([{ name:"sample.log", text:sample }]);
    setStatus("Sample loaded");
  });

  // ---------- Convert ----------
  btnConvert.addEventListener("click", () => {
    if(selectedInputs.length === 0){
      setStatus("No input to convert", false);
      metaBox.textContent = "Status: no input to convert.";
      return;
    }
    if(!validateCustomDateOrNotify()){
      return;
    }

    let oldClockIssuesByIndex = new Map();
    if(!optCustom.checked){
      const problematic = selectedInputs
        .map((entry, idx) => ({
          name: entry.name,
          text: entry.text || "",
          sourceIndex: idx,
          issue: findTooOldShowClockIssue(entry.text || ""),
        }))
        .filter((entry) => entry.issue !== null);
      oldClockIssuesByIndex = new Map(problematic.map((entry) => [entry.sourceIndex, entry.issue]));
      if(problematic.length > 0){
        const singleFileInPlace = selectedInputs.length === 1 && problematic.length === 1;
        const allFilesTooOldInPlace = problematic.length === selectedInputs.length;
        if(singleFileInPlace || allFilesTooOldInPlace){
          // Keep in current tab when only one file is selected, or when all selected files are too old.
          enableCustomDateMode();
        } else {
          convertedResults = [];
          resetOutputViews();
          setStatus("Convert failed", false);
          showOldClockConvertError(problematic);
          updateLineCounts();
          return;
        }
      }
    }

    try{
      convertedResults = selectedInputs.map((entry, idx) => {
        const res = transform(entry.text || "");
        const oldClockIssue = oldClockIssuesByIndex.get(idx);
        if(oldClockIssue){
          res.meta.oldClockPrecheck = {
            tooOld: true,
            lineNo: Number(oldClockIssue.lineNo || 0),
            foundDate: String(oldClockIssue.foundDate || ""),
            oldestAllowedDate: String(oldClockIssue.oldestAllowedDate || ""),
            autoHandled: true,
          };
        }
        return { fileName: entry.name, input: entry.text || "", output: res.output, meta: res.meta, serials: res.serials };
      });
      renderActiveFileView();
      const downloadCount = buildDownloadButtons();
      btnDownload.disabled = downloadCount === 0;
      const oldClockWarned = convertedResults.filter((entry) => entry.meta?.oldClockPrecheck?.tooOld);
      if(oldClockWarned.length > 0){
        const firstIssue = oldClockWarned[0]?.meta?.oldClockPrecheck || null;
        const autoFixNotice = buildAutoFixNotice(firstIssue);
        let popupMessage = autoFixNotice;
        if(convertedResults.length > 1){
          popupMessage = `ระบบทำการ Convert ใหม่ให้ ${oldClockWarned.length} ไฟล์และตั้งวันที่เป็นวันที่ปัจจุบันแล้ว เนื่องจากระบบตรวจพบเวลา show clock เก่าเกินเงื่อนไข หากต้องการเปลี่ยนวัน/เดือน/ปี ให้เลื่อนขึ้นไปด้านบนแล้วปรับค่าในโหมดกำหนดเอง`;
        }
        if(convertedResults.length === 1){
          setStatus(`Converted 1 file - ${autoFixNotice}`);
        } else {
          setStatus(`Converted ${convertedResults.length} files - ${autoFixNotice}`);
        }
        showAutoFixModal(popupMessage);
      } else {
        setStatus(convertedResults.length > 1 ? `Converted ${convertedResults.length} files` : "Converted 1 file");
      }
    }catch(err){
      setStatus("Convert failed", false);
      metaBox.textContent = "Convert error: " + (err && err.message ? err.message : String(err));
    }
  });

  // ---------- Diff toggle ----------
  optOnlyChanges.addEventListener("change", () => {
    try{ refreshDiff(); }catch(e){}
  });

  // ---------- Copy output ----------
  btnCopyOut.addEventListener("click", async () => {
    try{
      await navigator.clipboard.writeText(outBox.value || "");
      setStatus("คัดลอก Config หลังแก้ไขแล้ว");
    }catch(err){
      setStatus("Copy failed (try Ctrl+C)", false);
    }
  });

  // ---------- Download button (just scroll to buttons) ----------
  btnDownload.addEventListener("click", () => {
    if(convertedResults.length === 0 && !lastResult){
      setStatus("Nothing converted yet", false);
      return;
    }

    let n = downloadLinks.querySelectorAll("a").length;
    if(n === 0) n = buildDownloadButtons();

    const rect = downloadLinks.getBoundingClientRect();
    const inView = rect.top >= 0 && rect.bottom <= window.innerHeight;
    if(!inView){
      const targetTop = window.scrollY + rect.top - 90;
      window.scrollTo({ top: Math.max(0, targetTop), behavior:"smooth" });
    }

    setStatus(`พร้อมดาวน์โหลด ${n} ไฟล์`);
  });

  // ---------- Init ----------
  setFileSelectionUI([]);
  setStatus("No file selected");
  importProblemFilesFromMessage();
  importProblemFilesFromQuery();
})();


