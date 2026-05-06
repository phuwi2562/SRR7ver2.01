const fs = require("fs");

const SCREENED_BASE = "xlsx_unpacked";
const POPULATION_BASE = "xlsx_source_unpacked";

function decodeXml(value = "") {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function sharedStrings(base) {
  const file = `${base}/xl/sharedStrings.xml`;
  if (!fs.existsSync(file)) return [];
  const xml = fs.readFileSync(file, "utf8");
  return xml
    .split("<si>")
    .slice(1)
    .map((part) => {
      const si = part.split("</si>")[0];
      const pieces = si.match(/<t[^>]*>[\s\S]*?<\/t>/g) || [];
      return decodeXml(
        pieces
          .map((piece) => piece.replace(/^<t[^>]*>/, "").replace(/<\/t>$/, ""))
          .join("")
      );
    });
}

function colIndex(ref) {
  const letters = (ref.match(/[A-Z]+/) || [""])[0];
  let value = 0;
  for (const letter of letters) value = value * 26 + letter.charCodeAt(0) - 64;
  return value - 1;
}

function readRows(base) {
  const strings = sharedStrings(base);
  const xml = fs.readFileSync(`${base}/xl/worksheets/sheet1.xml`, "utf8");
  const rows = [];
  for (const part of xml.split("<row").slice(1)) {
    const body = part.split("</row>")[0];
    const row = [];
    for (const cell of body.split("<c ").slice(1)) {
      const ref = (cell.match(/r="([A-Z]+\d+)"/) || [])[1];
      if (!ref) continue;
      const type = (cell.match(/t="([^"]+)"/) || [])[1];
      const raw = (cell.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
      const inline = (cell.match(/<t[^>]*>([\s\S]*?)<\/t>/) || [])[1];
      row[colIndex(ref)] = type === "s" ? strings[Number(raw)] || "" : decodeXml(inline ?? raw ?? "");
    }
    rows.push(row);
  }
  return rows;
}

function parseThaiDate(value) {
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

function personKey({ name, village, houseNo }) {
  return [cleanName(name), village, String(houseNo || "").replace(/,/g, "").trim()].join("|");
}

function yes(value) {
  return String(value || "").trim() === "ใช่";
}

function populationRecords() {
  const rows = readRows(POPULATION_BASE);
  return rows
    .slice(3)
    .filter((row) => row.some(Boolean))
    .map((row, index) => ({
      id: index + 1,
      name: cleanName(row[0]),
      sex: row[1] || "",
      houseNo: String(row[2] || "").replace(/,/g, "").trim(),
      village: village(row[3]),
      subdistrict: row[4] || "",
      volunteer: row[5] || "",
      diagnosedDm: yes(row[6]),
      diagnosedHt: yes(row[7]),
    }));
}

function riskFromScreening(row) {
  const sbp = row.sbp;
  const dbp = row.dbp;
  const dtx = row.dtx;
  const waist = row.waist;
  const bmi = row.bmi;
  const sex = row.sex;
  const preHt = ((sbp ?? 0) >= 130 && (sbp ?? 0) <= 139) || ((dbp ?? 0) >= 80 && (dbp ?? 0) <= 89);
  const preDm = (dtx ?? 0) >= 100 && (dtx ?? 0) <= 125;
  const waistRisk = waist != null && ((sex === "ชาย" && waist >= 36) || (sex === "หญิง" && waist >= 32));
  const bmiRisk = bmi != null && bmi >= 25;
  return preHt || preDm || waistRisk || bmiRisk;
}

function controlStatus(record) {
  if (!record.diagnosedDm && !record.diagnosedHt) return "";
  const dmOk = !record.diagnosedDm || (record.dtx != null && record.dtx < 130);
  const htOk = !record.diagnosedHt || (record.sbp != null && record.dbp != null && record.sbp < 140 && record.dbp < 90);
  return dmOk && htOk ? "ควบคุมได้" : "ควบคุมไม่ได้";
}

function group(record) {
  if (record.diagnosedDm && record.diagnosedHt) return "DM+HT";
  if (record.diagnosedDm) return "DM";
  if (record.diagnosedHt) return "HT";
  return riskFromScreening(record) ? "เสี่ยง" : "ปกติ";
}

function screenedRecords(popByKey) {
  const rows = readRows(SCREENED_BASE);
  return rows
    .slice(3)
    .filter((row) => row.some(Boolean))
    .map((row, index) => {
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
        screenedDate: parseThaiDate(row[13] || ""),
      };
      const pop = popByKey.get(personKey(base));
      const record = {
        ...base,
        populationId: pop?.id || null,
        diagnosedDm: pop?.diagnosedDm || false,
        diagnosedHt: pop?.diagnosedHt || false,
      };
      return { ...record, group: group(record), control: controlStatus(record) };
    });
}

const population = populationRecords();
const popByKey = new Map(population.map((record) => [personKey(record), record]));
const screened = screenedRecords(popByKey);
const screenedKeys = new Set(screened.map(personKey));
const unscreened = population
  .filter((record) => !screenedKeys.has(personKey(record)))
  .map((record) => {
    const withGroup = {
      ...record,
      sbp: null,
      dbp: null,
      dtx: null,
      bmi: null,
      waist: null,
      screenedDate: null,
      screenedDateText: "",
      screened: false,
    };
    return { ...withGroup, group: group(withGroup), control: "" };
  });

const meta = {
  title: "SRR7 NCD Dashboard",
  screenedSourceFile: "ไฟล์ส่งออกจาก SRR7 คัดกรอง.xlsx",
  populationSourceFile: "ไฟล์ส่งออกจาก SRR7.xlsx",
  screenedExportedAt: readRows(SCREENED_BASE)[1]?.[1] || "",
  populationExportedAt: readRows(POPULATION_BASE)[1]?.[1] || "",
  generatedAt: new Date().toISOString(),
  rules: [
    "ฐานประชากร ≥35 ปีมาจากไฟล์ส่งออกจาก SRR7.xlsx",
    "คัดกรองแล้วมาจากไฟล์ส่งออกจาก SRR7 คัดกรอง.xlsx และจับคู่กับฐานประชากรด้วยชื่อ หมู่ และเลขที่บ้าน",
    "กลุ่ม DM/HT/DM+HT ใช้สถานะโรคเดิมจากฐานประชากร",
    "ควบคุมได้: DM ใช้ DTX < 130, HT ใช้ SBP < 140 และ DBP < 90 เฉพาะผู้ที่มีผลคัดกรอง",
  ],
};

fs.writeFileSync(
  "srr7-data.js",
  `window.SRR7_DATA = ${JSON.stringify({ meta, records: screened, population, unscreened }, null, 2)};\n`,
  "utf8"
);

console.log(`Wrote srr7-data.js with ${population.length} population and ${screened.length} screened records.`);
