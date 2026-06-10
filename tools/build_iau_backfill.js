/**
 * One-time builder for iau_backfill.json — IAU (iShares Gold Trust) daily
 * NAV per share + shares outstanding, recovered from a Wayback Machine
 * snapshot of the legacy iShares US fund download (removed from the live
 * site in the 2025 redesign).
 *
 * Run: node tools/build_iau_backfill.js
 */
const fs = require("fs");
const path = require("path");

const SNAPSHOT =
  "http://web.archive.org/web/20250216120056id_/https://www.ishares.com/us/products/239561/ishares-gold-trust-fund/1521942788811.ajax?fileType=xls&fileName=iShares-Gold-Trust_fund&dataType=fund";

const MONTHS = { Jan: "01", Feb: "02", Mar: "03", Apr: "04", May: "05", Jun: "06", Jul: "07", Aug: "08", Sep: "09", Oct: "10", Nov: "11", Dec: "12" };

(async () => {
  console.log("Downloading Wayback snapshot (Feb 16, 2025)...");
  const res = await fetch(SNAPSHOT, { headers: { "User-Agent": "Mozilla/5.0" } });
  if (!res.ok) throw new Error("HTTP " + res.status);
  const xml = await res.text();
  console.log(`Got ${(xml.length / 1024).toFixed(0)} KB (archive truncates at 1 MB)`);

  // Historical sheet rows: As Of | NAV per Share | Ex-Dividends | Shares Outstanding
  const histStart = xml.indexOf('<ss:Worksheet ss:Name="Historical"');
  if (histStart < 0) throw new Error("Historical sheet not found");
  const hist = xml.slice(histStart);

  const out = [];
  const rowRe = /<ss:Row>([\s\S]*?)<\/ss:Row>/g;
  let m;
  while ((m = rowRe.exec(hist))) {
    const cells = [];
    const cellRe = /<ss:Data ss:Type="(String|Number)">([\s\S]*?)<\/ss:Data>/g;
    let c;
    while ((c = cellRe.exec(m[1]))) cells.push(c[1] === "Number" ? parseFloat(c[2]) : c[2].trim());
    const dm = String(cells[0] || "").match(/^([A-Za-z]{3}) (\d{2}), (\d{4})$/);
    if (!dm || !MONTHS[dm[1]] || typeof cells[1] !== "number" || typeof cells[3] !== "number") continue;
    out.push({ date: `${dm[3]}-${MONTHS[dm[1]]}-${dm[2]}`, nav: cells[1], shares: cells[3] });
  }
  out.sort((a, b) => a.date.localeCompare(b.date));
  if (!out.length) throw new Error("No rows parsed");

  fs.writeFileSync(path.join(__dirname, "..", "iau_backfill.json"), JSON.stringify(out));
  console.log(`Parsed ${out.length} days: ${out[0].date} -> ${out[out.length - 1].date}`);
  console.log(`Latest: NAV $${out[out.length - 1].nav}, shares ${out[out.length - 1].shares.toLocaleString()}`);
})().catch((e) => { console.error("ERROR:", e.message); process.exit(1); });
