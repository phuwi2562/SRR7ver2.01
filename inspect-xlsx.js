const fs = require("fs");

function unesc(s) {
  return (s || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

const ss = fs.readFileSync("xlsx_unpacked/xl/sharedStrings.xml", "utf8");
const strings = ss
  .split("<si>")
  .slice(1)
  .map((part) =>
    unesc(
      (part.split("</si>")[0].match(/<t[^>]*>(.*?)<\/t>/gs) || [])
        .map((t) => t.replace(/^<t[^>]*>/, "").replace(/<\/t>$/, ""))
        .join("")
    )
  );

const xml = fs.readFileSync("xlsx_unpacked/xl/worksheets/sheet1.xml", "utf8");
function colNum(ref) {
  let letters = (ref.match(/[A-Z]+/) || [""])[0];
  let n = 0;
  for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64;
  return n - 1;
}

const rowParts = xml.split("<row").slice(1, 13);
const rows = [];
for (const part of rowParts) {
  const body = part.split("</row>")[0];
  const row = [];
  const cells = body.split("<c ").slice(1);
  for (const cell of cells) {
    const ref = (cell.match(/r="([A-Z]+\d+)"/) || [])[1];
    if (!ref) continue;
    const idx = colNum(ref);
    const t = (cell.match(/t="([^"]+)"/) || [])[1];
    const v = (cell.match(/<v>(.*?)<\/v>/s) || [])[1];
    const inline = (cell.match(/<t[^>]*>(.*?)<\/t>/s) || [])[1];
    row[idx] = t === "s" ? strings[Number(v)] || "" : unesc(inline ?? v ?? "");
  }
  rows.push(row);
}

console.log(JSON.stringify(rows, null, 2));
