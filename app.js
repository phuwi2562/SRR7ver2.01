const { meta, population } = window.SRR7_DATA;
let records = window.SRR7_DATA.records;
let unscreened = window.SRR7_DATA.unscreened || [];
const villages = Array.from({ length: 12 }, (_, i) => String(i + 1));
const colors = {
  normal: "#117553",
  risk: "#c9842b",
  dm: "#c54949",
  ht: "#376eb8",
  both: "#7a54c4",
  screened: "#128b96",
};

const els = {
  sourceInfo: document.querySelector("#sourceInfo"),
  totalRecords: document.querySelector("#totalRecords"),
  populationTotal: document.querySelector("#populationTotal"),
  rangeMode: document.querySelector("#rangeMode"),
  startDate: document.querySelector("#startDate"),
  endDate: document.querySelector("#endDate"),
  screeningImport: document.querySelector("#screeningImport"),
  importStatus: document.querySelector("#importStatus"),
  villageFilter: document.querySelector("#villageFilter"),
  activeRange: document.querySelector("#activeRange"),
  topVillages: document.querySelector("#topVillages"),
  topVhvs: document.querySelector("#topVhvs"),
  villageRows: document.querySelector("#villageRows"),
  registrySubtitle: document.querySelector("#registrySubtitle"),
  registryCount: document.querySelector("#registryCount"),
  registryRows: document.querySelector("#registryRows"),
  registryTabs: document.querySelectorAll(".registry-tab"),
  personDialog: document.querySelector("#personDialog"),
  personDialogTitle: document.querySelector("#personDialogTitle"),
  personDialogMeta: document.querySelector("#personDialogMeta"),
  personBp: document.querySelector("#personBp"),
  personBpStatus: document.querySelector("#personBpStatus"),
  personDtx: document.querySelector("#personDtx"),
  personDtxStatus: document.querySelector("#personDtxStatus"),
  personBmi: document.querySelector("#personBmi"),
  personBmiStatus: document.querySelector("#personBmiStatus"),
  personCvd: document.querySelector("#personCvd"),
  personHistoryRows: document.querySelector("#personHistoryRows"),
  measureDate: document.querySelector("#measureDate"),
  measureWeight: document.querySelector("#measureWeight"),
  measureHeight: document.querySelector("#measureHeight"),
  saveMeasure: document.querySelector("#saveMeasure"),
  measureMessage: document.querySelector("#measureMessage"),
  registryLogin: document.querySelector("#registryLogin"),
  registryLoginForm: document.querySelector("#registryLoginForm"),
  registryUsername: document.querySelector("#registryUsername"),
  registryPassword: document.querySelector("#registryPassword"),
  registryContent: document.querySelector("#registryContent"),
  registryActions: document.querySelector(".registry-actions"),
  registryAddressFilter: document.querySelector("#registryAddressFilter"),
  registryVolunteerFilter: document.querySelector("#registryVolunteerFilter"),
  clearRegistryFilters: document.querySelector("#clearRegistryFilters"),
  logoutRegistry: document.querySelector("#logoutRegistry"),
  loginMessage: document.querySelector("#loginMessage"),
};

const REGISTRY_AUTH = { username: "admin", password: "srr7@2569" };
let charts = {};
let activeRegistry = "htRisk";
let registryUnlocked = sessionStorage.getItem("srr7-registry-auth") === "ok";
let selectedPersonKey = "";
let personCharts = {};

Chart.defaults.font.family = '"Noto Sans Thai", "Segoe UI", Tahoma, sans-serif';
Chart.defaults.color = "#52615a";
Chart.defaults.plugins.tooltip.backgroundColor = "#10251f";
Chart.defaults.plugins.tooltip.padding = 12;
Chart.defaults.plugins.tooltip.cornerRadius = 10;

function fmt(n) {
  return Number(n || 0).toLocaleString("th-TH");
}

function fmtDecimal(n, digits = 1) {
  return Number.isFinite(Number(n)) ? Number(n).toLocaleString("th-TH", { minimumFractionDigits: digits, maximumFractionDigits: digits }) : "-";
}

function number(value) {
  const n = Number(String(value ?? "").replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : null;
}

function village(value) {
  const text = String(value || "").trim();
  const n = Number(text);
  return Number.isFinite(n) && n > 0 ? String(n) : text;
}

function cleanName(value) {
  return String(value || "")
    .replace(/[.]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function personKey(record) {
  return [cleanName(record.name), village(record.village), String(record.houseNo || "").replace(/,/g, "").trim()].join("|");
}

function keyToId(key) {
  return btoa(unescape(encodeURIComponent(key))).replace(/=+$/g, "");
}

function idToKey(id) {
  return decodeURIComponent(escape(atob(id)));
}

function parseThaiDateText(value) {
  if (!value) return null;
  const monthMap = {
    "ม.ค.": 1,
    "ก.พ.": 2,
    "มี.ค.": 3,
    "เม.ย.": 4,
    "พ.ค.": 5,
    "มิ.ย.": 6,
    "ก.ค.": 7,
    "ส.ค.": 8,
    "ก.ย.": 9,
    "ต.ค.": 10,
    "พ.ย.": 11,
    "ธ.ค.": 12,
  };
  const match = String(value).trim().match(/(\d{1,2})\s+(\S+)\s+(\d{4})/);
  if (!match) return null;
  const day = Number(match[1]);
  const month = monthMap[match[2]];
  const year = Number(match[3]) - 543;
  if (!day || !month || !year) return null;
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function parseDate(value) {
  return value ? new Date(`${value}T00:00:00+07:00`) : null;
}

function iso(date) {
  return date.toISOString().slice(0, 10);
}

function todayBangkok() {
  return new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Bangkok" }));
}

function fiscalBounds(now = todayBangkok()) {
  const year = now.getMonth() + 1 >= 10 ? now.getFullYear() : now.getFullYear() - 1;
  return [new Date(year, 9, 1), new Date(year + 1, 8, 30)];
}

function weekBounds(now = todayBangkok()) {
  const start = new Date(now);
  const day = start.getDay() || 7;
  start.setDate(start.getDate() - day + 1);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  return [start, end];
}

function monthBounds(now = todayBangkok()) {
  return [new Date(now.getFullYear(), now.getMonth(), 1), new Date(now.getFullYear(), now.getMonth() + 1, 0)];
}

function currentRange() {
  const mode = els.rangeMode.value;
  const now = todayBangkok();
  if (mode === "all") return [null, null, "ทั้งหมด"];
  if (mode === "today") return [new Date(now.getFullYear(), now.getMonth(), now.getDate()), new Date(now.getFullYear(), now.getMonth(), now.getDate()), "วันนี้"];
  if (mode === "week") return [...weekBounds(now), "สัปดาห์นี้"];
  if (mode === "month") return [...monthBounds(now), "เดือนนี้"];
  if (mode === "fiscal") return [...fiscalBounds(now), "ปีงบประมาณ"];
  return [parseDate(els.startDate.value), parseDate(els.endDate.value), "กำหนดเอง"];
}

function inRange(record, start, end) {
  const date = parseDate(record.screenedDate);
  if (!date) return false;
  if (start && date < start) return false;
  if (end && date > end) return false;
  return true;
}

function currentVillage() {
  return els.villageFilter.value;
}

function filteredPopulation() {
  const village = currentVillage();
  return population.filter((record) => village === "all" || record.village === village);
}

function filteredRecords() {
  const [start, end] = currentRange();
  const village = currentVillage();
  return records.filter((record) => {
    if (village !== "all" && record.village !== village) return false;
    return inRange(record, start, end);
  });
}

function blankCounts() {
  return { target: 0, screened: 0, matchedScreened: 0, normal: 0, risk: 0, dm: 0, ht: 0, both: 0, controlled: 0, uncontrolled: 0 };
}

function summarize(screenedList, populationList) {
  const total = blankCounts();
  const byVillage = Object.fromEntries(villages.map((v) => [v, blankCounts()]));

  for (const person of populationList) {
    if (!byVillage[person.village]) continue;
    byVillage[person.village].target += 1;
    total.target += 1;
  }

  for (const record of screenedList) {
    const bucket = byVillage[record.village];
    if (!bucket) continue;
    bucket.screened += 1;
    total.screened += 1;
    if (record.populationId) bucket.matchedScreened += 1, total.matchedScreened += 1;
    if (record.group === "ปกติ") bucket.normal += 1, total.normal += 1;
    if (record.group === "เสี่ยง") bucket.risk += 1, total.risk += 1;
    if (record.group === "DM") bucket.dm += 1, total.dm += 1;
    if (record.group === "HT") bucket.ht += 1, total.ht += 1;
    if (record.group === "DM+HT") bucket.both += 1, total.both += 1;
    if (record.control === "ควบคุมได้") bucket.controlled += 1, total.controlled += 1;
    if (record.control === "ควบคุมไม่ได้") bucket.uncontrolled += 1, total.uncontrolled += 1;
  }
  return { total, byVillage };
}

function updateKpis(total) {
  document.querySelector("#kpiScreened").textContent = fmt(total.screened);
  document.querySelector("#kpiUnscreened").textContent = fmt(Math.max(total.target - total.matchedScreened, 0));
  document.querySelector("#kpiNormal").textContent = fmt(total.normal);
  document.querySelector("#kpiRisk").textContent = fmt(total.risk);
  document.querySelector("#kpiDm").textContent = fmt(total.dm);
  document.querySelector("#kpiHt").textContent = fmt(total.ht);
  document.querySelector("#kpiBoth").textContent = fmt(total.both);
  document.querySelector("#kpiControlled").textContent = fmt(total.controlled);
  document.querySelector("#kpiUncontrolled").textContent = fmt(total.uncontrolled);
}

function updateTable(byVillage) {
  els.villageRows.innerHTML = villages
    .map((v) => {
      const c = byVillage[v];
      return `<tr>
        <td>หมู่ ${v}</td>
        <td>${fmt(c.target)}</td>
        <td>${fmt(c.screened)}</td>
        <td>${fmt(Math.max(c.target - c.matchedScreened, 0))}</td>
        <td>${fmt(c.normal)}</td>
        <td>${fmt(c.risk)}</td>
        <td>${fmt(c.dm)}</td>
        <td>${fmt(c.ht)}</td>
        <td>${fmt(c.both)}</td>
        <td>${fmt(c.controlled)}</td>
        <td>${fmt(c.uncontrolled)}</td>
      </tr>`;
    })
    .join("");
}

function updateLeaderboard(byVillage) {
  const leaders = villages
    .map((v) => {
      const c = byVillage[v];
      const unscreened = Math.max(c.target - c.matchedScreened, 0);
      const percent = c.target ? (c.matchedScreened / c.target) * 100 : 0;
      return { village: v, screened: c.matchedScreened, target: c.target, unscreened, percent };
    })
    .filter((item) => item.target > 0)
    .sort((a, b) => b.percent - a.percent || b.screened - a.screened)
    .slice(0, 3);

  els.topVillages.innerHTML = leaders
    .map((item, index) => {
      const stars = "★".repeat(3 - index);
      return `<article class="leader-card">
        <div class="leader-rank">
          <span>อันดับ ${index + 1}</span>
          <span class="leader-stars">${stars}</span>
        </div>
        <div class="leader-title">หมู่ ${item.village}</div>
        <div class="leader-percent">${item.percent.toFixed(1)}%</div>
        <div class="leader-meta">คัดกรอง ${fmt(item.screened)} จาก ${fmt(item.target)} คน | เหลือ ${fmt(item.unscreened)} คน</div>
      </article>`;
    })
    .join("");
}

function updateVhvLeaderboard(screenedList) {
  const byWorker = new Map();
  for (const record of screenedList) {
    const worker = workerName(record).trim();
    if (!worker) continue;
    if (!byWorker.has(worker)) {
      byWorker.set(worker, { worker, count: 0, villages: new Set() });
    }
    const item = byWorker.get(worker);
    item.count += 1;
    if (record.village) item.villages.add(record.village);
  }

  const leaders = Array.from(byWorker.values())
    .sort((a, b) => b.count - a.count || a.worker.localeCompare(b.worker, "th"))
    .slice(0, 10);

  if (!leaders.length) {
    els.topVhvs.innerHTML = `<tr><td class="empty-row" colspan="5">ไม่พบข้อมูล อสม. ในเงื่อนไขนี้</td></tr>`;
    return;
  }

  els.topVhvs.innerHTML = leaders
    .map((item, index) => {
      const villagesText = Array.from(item.villages)
        .sort((a, b) => Number(a) - Number(b))
        .map((v) => `หมู่ ${v}`)
        .join(", ");
      const stars = index < 3 ? "★".repeat(3 - index) : "";
      return `<tr>
        <td><span class="rank-badge">${index + 1}</span></td>
        <td>${escapeHtml(item.worker)}</td>
        <td>${escapeHtml(villagesText || "-")}</td>
        <td>${fmt(item.count)}</td>
        <td><span class="award-stars">${stars || "-"}</span></td>
      </tr>`;
    })
    .join("");
}

function isHtRisk(record) {
  if (record.diagnosedHt) return false;
  const sbp = record.sbp ?? 0;
  const dbp = record.dbp ?? 0;
  return (sbp >= 130 && sbp <= 139) || (dbp >= 80 && dbp <= 89);
}

function isDmRisk(record) {
  if (record.diagnosedDm) return false;
  const dtx = record.dtx ?? 0;
  return dtx >= 100 && dtx <= 125;
}

function screeningRisk(record) {
  const preHt = isHtRisk(record);
  const preDm = isDmRisk(record);
  const waistRisk = record.waist != null && ((record.sex === "ชาย" && record.waist >= 36) || (record.sex === "หญิง" && record.waist >= 32));
  const bmiRisk = record.bmi != null && record.bmi >= 25;
  return preHt || preDm || waistRisk || bmiRisk;
}

function recordGroup(record) {
  if (record.diagnosedDm && record.diagnosedHt) return "DM+HT";
  if (record.diagnosedDm) return "DM";
  if (record.diagnosedHt) return "HT";
  return screeningRisk(record) ? "เสี่ยง" : "ปกติ";
}

function controlStatus(record) {
  if (!record.diagnosedDm && !record.diagnosedHt) return "";
  const dmOk = !record.diagnosedDm || (record.dtx != null && record.dtx < 130);
  const htOk = !record.diagnosedHt || (record.sbp != null && record.dbp != null && record.sbp < 140 && record.dbp < 90);
  return dmOk && htOk ? "ควบคุมได้" : "ควบคุมไม่ได้";
}

function registryVillageFilter(list) {
  const village = currentVillage();
  return list.filter((record) => village === "all" || record.village === village);
}

function workerName(record) {
  return record.recorder || record.volunteer || "";
}

function addressText(record) {
  return [`บ้านเลขที่ ${record.houseNo || ""}`, `หมู่ ${record.village || ""}`, record.houseNo || "", record.village || ""].join(" ").toLowerCase();
}

function registrySort(a, b) {
  return Number(a.village) - Number(b.village) || String(a.name).localeCompare(String(b.name), "th");
}

function registryItems(screenedList) {
  let list;
  if (activeRegistry === "htRisk") list = registryVillageFilter(screenedList.filter(isHtRisk));
  else if (activeRegistry === "dmRisk") list = registryVillageFilter(screenedList.filter(isDmRisk));
  else list = registryVillageFilter(unscreened || []);

  const addressQuery = els.registryAddressFilter.value.trim().toLowerCase();
  const volunteer = els.registryVolunteerFilter.value;
  if (addressQuery) list = list.filter((record) => addressText(record).includes(addressQuery));
  if (volunteer !== "all") list = list.filter((record) => workerName(record) === volunteer);
  return list.sort(registrySort);
}

function baseRegistryItems(screenedList) {
  if (activeRegistry === "htRisk") return registryVillageFilter(screenedList.filter(isHtRisk));
  if (activeRegistry === "dmRisk") return registryVillageFilter(screenedList.filter(isDmRisk));
  return registryVillageFilter(unscreened || []);
}

function updateVolunteerOptions(screenedList) {
  const current = els.registryVolunteerFilter.value;
  const names = Array.from(new Set(baseRegistryItems(screenedList).map(workerName).filter(Boolean))).sort((a, b) => a.localeCompare(b, "th"));
  const village = currentVillage();
  const allLabel = village === "all" ? "ทุก อสม. / ผู้บันทึก" : `ทุก อสม. / ผู้บันทึกในหมู่ ${village}`;
  els.registryVolunteerFilter.innerHTML = `<option value="all">${allLabel}</option>${names.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("")}`;
  els.registryVolunteerFilter.value = names.includes(current) ? current : "all";
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function registryInfo() {
  if (activeRegistry === "htRisk") {
    return { title: "ทะเบียนกลุ่มเสี่ยงโรคความดันโลหิตสูง", note: "SBP 130-139 หรือ DBP 80-89 และยังไม่เป็นผู้ป่วย HT เดิม" };
  }
  if (activeRegistry === "dmRisk") {
    return { title: "ทะเบียนกลุ่มเสี่ยงโรคเบาหวาน", note: "DTX 100-125 และยังไม่เป็นผู้ป่วย DM เดิม" };
  }
  return { title: "ทะเบียนรายชื่อผู้ยังไม่ได้รับการคัดกรอง", note: "รายชื่อจากฐานประชากรที่ยังไม่พบผลคัดกรองในไฟล์คัดกรอง" };
}

function registryRemark(record) {
  if (activeRegistry === "htRisk") return "เสี่ยง HT";
  if (activeRegistry === "dmRisk") return "เสี่ยง DM";
  const disease = [];
  if (record.diagnosedDm) disease.push("DM เดิม");
  if (record.diagnosedHt) disease.push("HT เดิม");
  return disease.join(" / ") || "ยังไม่คัดกรอง";
}

function updateRegistry(screenedList) {
  updateRegistryLockState();
  if (!registryUnlocked) {
    els.registrySubtitle.textContent = "กรุณาเข้าสู่ระบบเพื่อดูทะเบียนรายชื่อ";
    els.registryCount.textContent = "ล็อก";
    els.registryRows.innerHTML = "";
    return;
  }
  updateVolunteerOptions(screenedList);
  const items = registryItems(screenedList);
  const info = registryInfo();
  const villageText = currentVillage() === "all" ? "หมู่ 1-12" : `หมู่ ${currentVillage()}`;
  els.registrySubtitle.textContent = `${info.title} | ${villageText} | ${info.note}`;
  els.registryCount.textContent = `${fmt(items.length)} ราย`;
  if (!items.length) {
    els.registryRows.innerHTML = `<tr><td class="empty-row" colspan="11">ไม่พบรายชื่อในเงื่อนไขนี้</td></tr>`;
    return;
  }
  els.registryRows.innerHTML = items
    .map((record, index) => {
      const bp = record.sbp != null || record.dbp != null ? `${record.sbp ?? "-"} / ${record.dbp ?? "-"}` : "-";
      const dtx = record.dtx != null ? fmt(record.dtx) : "-";
      const worker = workerName(record) || "-";
      const personId = keyToId(personKey(record));
      return `<tr>
        <td>${fmt(index + 1)}</td>
        <td>${escapeHtml(record.name || "-")}</td>
        <td>${escapeHtml(record.sex || "-")}</td>
        <td>${escapeHtml(record.houseNo || "-")}</td>
        <td>หมู่ ${escapeHtml(record.village || "-")}</td>
        <td>${escapeHtml(worker)}</td>
        <td>${bp}</td>
        <td>${dtx}</td>
        <td>${escapeHtml(record.screenedDateText || "-")}</td>
        <td>${escapeHtml(registryRemark(record))}</td>
        <td><button class="inline-action person-open" type="button" data-person="${personId}">ดู/บันทึก</button></td>
      </tr>`;
    })
    .join("");
}

function parseScreeningRows(rows) {
  const dataRows = rows.slice(3).filter((row) => row.some((cell) => cell !== undefined && cell !== null && String(cell).trim() !== ""));
  const popByKey = new Map(population.map((person) => [personKey(person), person]));
  const parsed = dataRows.map((row, index) => {
    const base = {
      id: index + 1,
      name: cleanName(row[0]),
      sex: row[1] || "",
      houseNo: String(row[2] || "").replace(/,/g, "").trim(),
      village: village(row[3]),
      subdistrict: row[4] || "",
      recorder: row[5] || "",
      sbp: number(row[6]),
      dbp: number(row[7]),
      dtx: number(row[8]),
      bmi: number(row[9]),
      waist: number(row[10]),
      alcohol: row[11] || "",
      smoking: row[12] || "",
      screenedDateText: row[13] || "",
      screenedDate: parseThaiDateText(row[13] || ""),
    };
    const pop = popByKey.get(personKey(base));
    const record = {
      ...base,
      populationId: pop?.id || null,
      diagnosedDm: pop?.diagnosedDm || false,
      diagnosedHt: pop?.diagnosedHt || false,
    };
    return { ...record, group: recordGroup(record), control: controlStatus(record) };
  });
  return parsed;
}

function rebuildUnscreened() {
  const screenedKeys = new Set(records.map(personKey));
  unscreened = population
    .filter((person) => !screenedKeys.has(personKey(person)))
    .map((person) => {
      const record = {
        ...person,
        sbp: null,
        dbp: null,
        dtx: null,
        bmi: null,
        waist: null,
        screenedDate: null,
        screenedDateText: "",
      };
      return { ...record, group: recordGroup(record), control: "" };
    });
}

function setImportStatus(message, type = "") {
  els.importStatus.textContent = message;
  els.importStatus.className = type;
}

async function importScreeningFile(file) {
  if (!window.XLSX) {
    setImportStatus("ไม่พบไลบรารีอ่าน Excel กรุณาเชื่อมต่ออินเทอร์เน็ตแล้วรีเฟรชหน้า", "error");
    return;
  }
  try {
    setImportStatus("กำลังอ่านไฟล์...");
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const imported = parseScreeningRows(rows);
    if (!imported.length) throw new Error("empty");
    records = imported;
    rebuildUnscreened();
    els.totalRecords.textContent = fmt(records.length);
    setImportStatus(`นำเข้า ${file.name} สำเร็จ: ${fmt(records.length)} รายการ`, "success");
    render();
  } catch (error) {
    setImportStatus("นำเข้าไฟล์ไม่สำเร็จ กรุณาตรวจสอบว่าเป็นไฟล์คัดกรองจาก SRR7", "error");
  }
}

function updateRegistryLockState() {
  els.registryLogin.classList.toggle("unlocked", registryUnlocked);
  els.registryContent.classList.toggle("locked", !registryUnlocked);
  els.registryActions.classList.toggle("locked", !registryUnlocked);
}

function measureStore() {
  try {
    return JSON.parse(localStorage.getItem("srr7-person-measures") || "{}");
  } catch {
    return {};
  }
}

function saveMeasureStore(store) {
  localStorage.setItem("srr7-person-measures", JSON.stringify(store));
}

function personMeasures(key) {
  return (measureStore()[key] || []).sort((a, b) => String(a.date).localeCompare(String(b.date)));
}

function findPerson(key) {
  return records.find((record) => personKey(record) === key) || unscreened.find((record) => personKey(record) === key) || population.find((record) => personKey(record) === key);
}

function bmiFromMeasure(measure) {
  const weight = Number(measure.weight);
  const heightM = Number(measure.height) / 100;
  return weight && heightM ? weight / (heightM * heightM) : null;
}

function bmiStatus(bmi) {
  if (!Number.isFinite(Number(bmi))) return "-";
  if (bmi >= 25) return "เสี่ยง";
  if (bmi >= 18.5) return "ปกติ";
  return "ต่ำกว่าเกณฑ์";
}

function bpStatus(record) {
  if (record?.sbp == null && record?.dbp == null) return "-";
  return (record.sbp ?? 0) >= 140 || (record.dbp ?? 0) >= 90 ? "เสี่ยง" : "ปกติ";
}

function dtxStatus(record) {
  if (record?.dtx == null) return "-";
  return record.dtx >= 126 ? "เสี่ยง" : record.dtx >= 100 ? "เฝ้าระวัง" : "ปกติ";
}

function thaiTodayIso() {
  return iso(todayBangkok());
}

function latestMeasure(key) {
  const list = personMeasures(key);
  return list[list.length - 1] || null;
}

function historyRows(person, key) {
  const rows = [];
  const measures = personMeasures(key);
  for (const measure of measures) {
    const bmi = bmiFromMeasure(measure);
    rows.push({
      date: measure.date,
      weight: measure.weight,
      height: measure.height,
      bmi,
      bp: "-",
      dtx: "-",
      smoking: "-",
      alcohol: "-",
    });
  }
  if (person?.screenedDate || person?.screenedDateText) {
    rows.push({
      date: person.screenedDate || person.screenedDateText,
      weight: "-",
      height: "-",
      bmi: person.bmi,
      bp: person.sbp != null || person.dbp != null ? `${person.sbp ?? "-"} / ${person.dbp ?? "-"}` : "-",
      dtx: person.dtx ?? "-",
      smoking: person.smoking || "-",
      alcohol: person.alcohol || "-",
    });
  }
  return rows.sort((a, b) => String(b.date).localeCompare(String(a.date)));
}

function renderPersonCharts(person, key) {
  for (const chart of Object.values(personCharts)) chart.destroy();
  personCharts = {};
  const rows = historyRows(person, key).slice().reverse();
  const labels = rows.map((row) => row.date);
  personCharts.bp = new Chart(document.querySelector("#personBpChart"), {
    type: "line",
    data: {
      labels,
      datasets: [
        { label: "SBP", data: rows.map((row) => (typeof row.bp === "string" && row.bp.includes("/") ? Number(row.bp.split("/")[0]) : null)), borderColor: colors.ht, tension: 0.35 },
        { label: "DBP", data: rows.map((row) => (typeof row.bp === "string" && row.bp.includes("/") ? Number(row.bp.split("/")[1]) : null)), borderColor: "#e0527f", tension: 0.35 },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: false } } },
  });
  personCharts.metabolic = new Chart(document.querySelector("#personMetabolicChart"), {
    type: "line",
    data: {
      labels,
      datasets: [
        { label: "DTX", data: rows.map((row) => Number(row.dtx) || null), borderColor: colors.screened, tension: 0.35 },
        { label: "BMI", data: rows.map((row) => Number(row.bmi) || null), borderColor: colors.risk, tension: 0.35 },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: false } } },
  });
}

function openPersonDialog(key) {
  if (!registryUnlocked) return;
  const person = findPerson(key);
  if (!person) return;
  selectedPersonKey = key;
  const measure = latestMeasure(key);
  const latestBmi = measure ? bmiFromMeasure(measure) : person.bmi;
  els.personDialogTitle.textContent = `${person.name || "-"} (${person.sex || "-"}, cid: -)`;
  els.personDialogMeta.textContent = `บ้านเลขที่ ${person.houseNo || "-"} | หมู่ ${person.village || "-"} | ${workerName(person) || "ไม่ระบุ อสม./ผู้บันทึก"}`;
  els.personBp.textContent = person.sbp != null || person.dbp != null ? `${person.sbp ?? "-"} / ${person.dbp ?? "-"}` : "-";
  els.personBpStatus.textContent = bpStatus(person);
  els.personDtx.textContent = person.dtx != null ? fmt(person.dtx) : "-";
  els.personDtxStatus.textContent = dtxStatus(person);
  els.personBmi.textContent = latestBmi ? fmtDecimal(latestBmi, 2) : "-";
  els.personBmiStatus.textContent = bmiStatus(latestBmi);
  els.personCvd.textContent = "0.00%";
  els.measureDate.value = thaiTodayIso();
  els.measureWeight.value = "";
  els.measureHeight.value = "";
  els.measureMessage.textContent = "";
  renderPersonHistory(person, key);
  if (!els.personDialog.open) els.personDialog.showModal();
  setTimeout(() => renderPersonCharts(person, key), 0);
}

function renderPersonHistory(person, key) {
  const rows = historyRows(person, key);
  if (!rows.length) {
    els.personHistoryRows.innerHTML = `<tr><td class="empty-row" colspan="8">ยังไม่มีประวัติ</td></tr>`;
    return;
  }
  els.personHistoryRows.innerHTML = rows
    .map((row) => `<tr>
      <td>${escapeHtml(row.date || "-")}</td>
      <td>${row.weight === "-" ? "-" : fmtDecimal(row.weight, 1)}</td>
      <td>${row.height === "-" ? "-" : fmtDecimal(row.height, 1)}</td>
      <td>${row.bmi ? fmtDecimal(row.bmi, 2) : "-"}</td>
      <td>${escapeHtml(row.bp || "-")}</td>
      <td>${escapeHtml(row.dtx || "-")}</td>
      <td>${escapeHtml(row.smoking || "-")}</td>
      <td>${escapeHtml(row.alcohol || "-")}</td>
    </tr>`)
    .join("");
}

function saveCurrentMeasure() {
  if (!selectedPersonKey) return;
  const weight = Number(els.measureWeight.value);
  const height = Number(els.measureHeight.value);
  const date = els.measureDate.value;
  if (!date || !weight || !height) {
    els.measureMessage.textContent = "กรุณากรอกวันที่ น้ำหนัก และส่วนสูง";
    return;
  }
  const store = measureStore();
  store[selectedPersonKey] = store[selectedPersonKey] || [];
  const existingIndex = store[selectedPersonKey].findIndex((item) => item.date === date);
  const entry = { date, weight, height, savedAt: new Date().toISOString() };
  if (existingIndex >= 0) store[selectedPersonKey][existingIndex] = entry;
  else store[selectedPersonKey].push(entry);
  saveMeasureStore(store);
  openPersonDialog(selectedPersonKey);
  els.measureMessage.textContent = "บันทึกสำเร็จ";
}

function chart(id, type, data, options = {}) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(document.querySelector(`#${id}`), {
    type,
    data,
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: "bottom", labels: { boxWidth: 10, boxHeight: 10, usePointStyle: true } },
      },
      scales:
        type === "doughnut"
          ? undefined
          : {
              x: { grid: { display: false }, border: { display: false } },
              y: { beginAtZero: true, ticks: { precision: 0 }, border: { display: false }, grid: { color: "rgba(22,33,29,.08)" } },
            },
      ...options,
    },
  });
}

function monthKey(record) {
  return record.screenedDate ? record.screenedDate.slice(0, 7) : null;
}

function updateCharts(list, byVillage, total) {
  const labels = villages.map((v) => `หมู่ ${v}`);
  chart("patientsBar", "bar", {
    labels,
    datasets: [
      {
        label: "ผู้รับการคัดกรองกลุ่ม DM / HT / DM+HT",
        data: villages.map((v) => byVillage[v].dm + byVillage[v].ht + byVillage[v].both),
        backgroundColor: colors.red,
        borderRadius: 10,
        maxBarThickness: 36,
      },
    ],
  });

  chart(
    "riskDoughnut",
    "doughnut",
    {
      labels: ["ปกติ", "เสี่ยง", "DM", "HT", "DM+HT"],
      datasets: [
        {
          data: [total.normal, total.risk, total.dm, total.ht, total.both],
          backgroundColor: [colors.normal, colors.risk, colors.dm, colors.ht, colors.both],
          borderWidth: 4,
          borderColor: "#ffffff",
        },
      ],
    },
    { cutout: "68%" }
  );

  const monthly = {};
  for (const record of list) {
    const key = monthKey(record);
    if (key) monthly[key] = (monthly[key] || 0) + 1;
  }
  const monthLabels = Object.keys(monthly).sort();
  chart("monthlyLine", "line", {
    labels: monthLabels,
    datasets: [
      {
        label: "คัดกรอง",
        data: monthLabels.map((m) => monthly[m]),
        borderColor: colors.screened,
        backgroundColor: "rgba(18,139,150,.16)",
        pointBackgroundColor: "#ffffff",
        pointBorderColor: colors.screened,
        pointRadius: 4,
        fill: true,
        tension: 0.35,
      },
    ],
  });

  chart(
    "stackedVillage",
    "bar",
    {
      labels,
      datasets: [
        { label: "DM", data: villages.map((v) => byVillage[v].dm), backgroundColor: colors.dm, borderRadius: 8 },
        { label: "HT", data: villages.map((v) => byVillage[v].ht), backgroundColor: colors.ht, borderRadius: 8 },
        { label: "DM+HT", data: villages.map((v) => byVillage[v].both), backgroundColor: colors.both, borderRadius: 8 },
      ],
    },
    { scales: { x: { stacked: true, grid: { display: false }, border: { display: false } }, y: { stacked: true, beginAtZero: true, ticks: { precision: 0 }, border: { display: false }, grid: { color: "rgba(22,33,29,.08)" } } } }
  );
}

function dateLabel(start, end, name) {
  if (!start && !end) return name;
  const th = new Intl.DateTimeFormat("th-TH", { dateStyle: "medium" });
  return `${name}: ${start ? th.format(start) : "-"} ถึง ${end ? th.format(end) : "-"}`;
}

function render() {
  const screenedList = filteredRecords();
  const popList = filteredPopulation();
  const { total, byVillage } = summarize(screenedList, popList);
  const [start, end, name] = currentRange();
  updateKpis(total);
  updateTable(byVillage);
  updateLeaderboard(byVillage);
  updateVhvLeaderboard(screenedList);
  updateCharts(screenedList, byVillage, total);
  updateRegistry(screenedList);
  els.activeRange.textContent = dateLabel(start, end, name);
  document.querySelectorAll(".custom-date").forEach((node) => {
    node.style.display = els.rangeMode.value === "custom" ? "grid" : "none";
  });
}

function init() {
  els.sourceInfo.textContent = `${meta.populationSourceFile} (${meta.populationExportedAt}) + ${meta.screenedSourceFile} (${meta.screenedExportedAt})`;
  els.totalRecords.textContent = fmt(records.length);
  els.populationTotal.textContent = fmt(population.length);
  for (const v of villages) els.villageFilter.insertAdjacentHTML("beforeend", `<option value="${v}">หมู่ ${v}</option>`);
  const [fyStart, fyEnd] = fiscalBounds();
  els.startDate.value = iso(fyStart);
  els.endDate.value = iso(fyEnd);
  render();
}

els.rangeMode.addEventListener("change", render);
els.startDate.addEventListener("change", render);
els.endDate.addEventListener("change", render);
els.screeningImport.addEventListener("change", () => {
  const file = els.screeningImport.files?.[0];
  if (file) importScreeningFile(file);
});
els.villageFilter.addEventListener("change", () => {
  els.registryVolunteerFilter.value = "all";
  render();
});
els.registryAddressFilter.addEventListener("input", render);
els.registryVolunteerFilter.addEventListener("change", render);
els.clearRegistryFilters.addEventListener("click", () => {
  els.registryAddressFilter.value = "";
  els.registryVolunteerFilter.value = "all";
  render();
});
els.registryRows.addEventListener("click", (event) => {
  const button = event.target.closest(".person-open");
  if (!button) return;
  openPersonDialog(idToKey(button.dataset.person));
});
els.saveMeasure.addEventListener("click", saveCurrentMeasure);
els.personDialog.addEventListener("close", () => {
  for (const chart of Object.values(personCharts)) chart.destroy();
  personCharts = {};
});
els.registryTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    activeRegistry = tab.dataset.registry;
    els.registryAddressFilter.value = "";
    els.registryVolunteerFilter.value = "all";
    els.registryTabs.forEach((item) => item.classList.toggle("active", item === tab));
    render();
  });
});
els.registryLoginForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const username = els.registryUsername.value.trim();
  const password = els.registryPassword.value;
  if (username === REGISTRY_AUTH.username && password === REGISTRY_AUTH.password) {
    registryUnlocked = true;
    sessionStorage.setItem("srr7-registry-auth", "ok");
    els.registryPassword.value = "";
    els.loginMessage.textContent = "";
    render();
    return;
  }
  els.loginMessage.textContent = "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง";
});
els.logoutRegistry.addEventListener("click", () => {
  registryUnlocked = false;
  sessionStorage.removeItem("srr7-registry-auth");
  els.registryUsername.value = "";
  els.registryPassword.value = "";
  render();
});

init();
