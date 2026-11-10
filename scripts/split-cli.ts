// Chạy thử bộ tách từ dòng lệnh:  npm run split -- "<file.xlsx>" [thư mục kết quả]
import fs from "node:fs";
import path from "node:path";
import { ORG_LABEL, Org, outputFileName } from "../lib/xlsx/classify";
import { inspectWorkbook, splitWorkbook } from "../lib/xlsx/split";

async function main() {
  const [input, outDir = "out"] = process.argv.slice(2);
  if (!input) {
    console.error('Cách dùng: npm run split -- "<file.xlsx>" [thư mục kết quả]');
    process.exit(1);
  }
  const buf = fs.readFileSync(input);
  let t = Date.now();
  const ins = await inspectWorkbook(buf);
  console.log(`inspect: ${Date.now() - t} ms, ${ins.positions.length} vị trí`);
  for (const p of ins.positions.filter((x) => x.suggested === "HDND" || !x.confident)) {
    console.log(`  ${p.code.padEnd(14)} ${ORG_LABEL[p.suggested]} ${p.confident ? "  " : "?!"} "${p.name}" — ${p.reason}`);
  }
  const mapping: Record<string, Org> = Object.fromEntries(ins.positions.map((p) => [p.code, p.suggested]));
  t = Date.now();
  const res = await splitWorkbook(buf, mapping);
  const mem = process.memoryUsage();
  console.log(`split: ${Date.now() - t} ms | rss ${(mem.rss / 1048576).toFixed(0)} MB | vị trí HĐND ${res.report.positions.hdnd}, UBND ${res.report.positions.ubnd}`);
  console.table(res.report.sheets);
  for (const w of res.report.warnings) console.log("  ⚠", w);
  fs.mkdirSync(outDir, { recursive: true });
  for (const org of ["HDND", "UBND"] as Org[]) {
    const name = outputFileName(path.basename(input), org);
    fs.writeFileSync(path.join(outDir, name), res.files[org]);
    console.log(`  → ${name} (${(res.files[org].length / 1048576).toFixed(1)} MB)`);
  }
}
main().catch((e) => {
  console.error("LỖI:", e.code ?? "", e.message);
  process.exit(1);
});
