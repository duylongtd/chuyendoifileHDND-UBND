// Tách bộ biểu VTVL "HĐND-UBND" thành 2 workbook: HĐND và UBND.
// Nội dung ô giữ nguyên; chỉ chọn dòng theo khối và dồn lên liền nhau. Công thức được Excel tính lại khi mở.
import { DATA_START_ROW, Org, RowOrg, SheetSpec, specForSheet, suggestOrg, Suggestion } from "./classify";
import { buildSheetModel, cellOf, mcellText, MRow, renderSheet, rowHasConstant, SheetFormatError } from "./sheet-model";
import { getAttr, parseSharedStrings, setAttr, xmlUnescape } from "./xml";
import { deflated, deflatedChunks, inflateEntry, OutEntry, passthrough, readZip, writeZip, ZipEntry, ZipError } from "./zip";

export type SplitErrorCode = "INVALID_FILE" | "NOT_VTVL" | "UNSUPPORTED" | "ONE_SIDE_EMPTY" | "TOO_LARGE";

export class SplitError extends Error {
  constructor(
    public code: SplitErrorCode,
    message: string,
  ) {
    super(message);
  }
}

export interface PositionInfo {
  code: string;
  name: string;
  row: number;
  suggested: Org;
  confident: boolean;
  reason: string;
  /** Trích mô tả ở biểu C2 (giúp người dùng xác nhận khi tên trống) */
  hint: string;
}

export interface InspectResult {
  positions: PositionInfo[];
  sheets: string[];
  missingSheets: string[];
}

interface SheetRef {
  name: string;
  path: string;
  spec?: SheetSpec;
}

interface Workbook {
  entries: ZipEntry[];
  byName: Map<string, ZipEntry>;
  sheets: SheetRef[];
  shared: string[];
}

function wrap<T>(fn: () => T): T {
  try {
    return fn();
  } catch (e) {
    if (e instanceof ZipError) throw new SplitError("INVALID_FILE", e.message);
    if (e instanceof SheetFormatError) throw new SplitError("UNSUPPORTED", e.message);
    throw e;
  }
}

async function readBuffer(wb: Workbook, path: string, into?: Buffer): Promise<Buffer> {
  const e = wb.byName.get(path);
  if (!e) throw new SplitError("INVALID_FILE", `File thiếu thành phần ${path}.`);
  try {
    return await inflateEntry(e, into);
  } catch (err) {
    if (err instanceof ZipError) throw new SplitError("INVALID_FILE", err.message);
    throw new SplitError("INVALID_FILE", `Không giải nén được ${path}: file bị hỏng.`);
  }
}

async function readText(wb: Workbook, path: string): Promise<string> {
  return (await readBuffer(wb, path)).toString("utf8");
}

async function openWorkbook(buf: Buffer): Promise<Workbook> {
  const entries = wrap(() => readZip(buf));
  const byName = new Map(entries.map((e) => [e.name, e]));
  const wb: Workbook = { entries, byName, sheets: [], shared: [] };
  if (!byName.has("xl/workbook.xml") || !byName.has("xl/_rels/workbook.xml.rels")) {
    throw new SplitError("INVALID_FILE", "Không phải file Excel (.xlsx) hợp lệ.");
  }
  const wbXml = await readText(wb, "xl/workbook.xml");
  const rels = await readText(wb, "xl/_rels/workbook.xml.rels");
  const target = new Map<string, string>();
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = getAttr(m[0], "Id");
    const t = getAttr(m[0], "Target");
    if (id && t) target.set(id, t.startsWith("/") ? t.slice(1) : `xl/${t}`);
  }
  for (const m of wbXml.matchAll(/<sheet\b[^>]*>/g)) {
    const name = xmlUnescape(getAttr(m[0], "name") ?? "");
    const rid = getAttr(m[0], "r:id") ?? "";
    const path = target.get(rid);
    if (name && path && byName.has(path)) wb.sheets.push({ name, path, spec: specForSheet(name) });
  }
  if (byName.has("xl/sharedStrings.xml")) wb.shared = parseSharedStrings(await readText(wb, "xl/sharedStrings.xml"));
  return wb;
}

function sheetByRole(wb: Workbook, role: string): SheetRef | undefined {
  return wb.sheets.find((s) => s.spec?.role === role);
}

async function loadRows(wb: Workbook, sheet: SheetRef): Promise<MRow[]> {
  const raw = await readBuffer(wb, sheet.path);
  const model = wrap(() => buildSheetModel(raw, sheet.spec!.keyCol));
  return model ? model.rows : [];
}

/** Đọc danh sách vị trí ở biểu C1 và gợi ý khối HĐND/UBND cho từng vị trí. */
export async function inspectWorkbook(buf: Buffer): Promise<InspectResult> {
  const wb = await openWorkbook(buf);
  const c1 = sheetByRole(wb, "C1");
  if (!c1) {
    throw new SplitError("NOT_VTVL", 'File không có biểu "6. C1_VTVL" nên không phải bộ biểu VTVL (hoặc tên trang tính đã bị đổi).');
  }
  const hints = new Map<string, string>();
  const c2 = sheetByRole(wb, "C2");
  if (c2) {
    for (const r of await loadRows(wb, c2)) {
      const code = mcellText(cellOf(r, "A"), wb.shared);
      if (code && !hints.has(code)) hints.set(code, mcellText(cellOf(r, "B"), wb.shared));
    }
  }
  const positions: PositionInfo[] = [];
  const seen = new Set<string>();
  for (const r of await loadRows(wb, c1)) {
    const code = mcellText(cellOf(r, "A"), wb.shared);
    if (!code || seen.has(code)) continue;
    seen.add(code);
    const name = mcellText(cellOf(r, "B"), wb.shared);
    const hint = hints.get(code) ?? "";
    const s: Suggestion = suggestOrg(name, hint);
    positions.push({ code, name, row: r.num, suggested: s.org, confident: s.confident, reason: s.reason, hint: hint.slice(0, 220) });
  }
  if (!positions.length) throw new SplitError("NOT_VTVL", 'Biểu "6. C1_VTVL" chưa có vị trí việc làm nào (từ dòng 3).');
  const have = new Set(wb.sheets.map((s) => s.spec?.role).filter(Boolean));
  const missing = ["A2", "B1", "C3", "D1", "D2"].filter((r) => !have.has(r as never));
  return { positions, sheets: wb.sheets.map((s) => s.name), missingSheets: missing };
}

export interface SheetReport {
  sheet: string;
  total: number;
  hdnd: number;
  ubnd: number;
  both: number;
}

export interface SplitReport {
  positions: { hdnd: number; ubnd: number };
  sheets: SheetReport[];
  warnings: string[];
}

export interface SplitResult {
  files: Record<Org, Buffer>;
  report: SplitReport;
}

interface Ctx {
  vtvl: Map<string, Org>;
  cv: Map<string, RowOrg>;
  nv: Map<string, RowOrg>;
  warnings: Map<string, number>;
}

function warn(ctx: Ctx, msg: string) {
  ctx.warnings.set(msg, (ctx.warnings.get(msg) ?? 0) + 1);
}

/** Khối của một dòng; null = dòng mẫu trống (không có mã, không có ô nhập tay). */
function classifyRow(row: MRow, spec: SheetSpec, shared: string[], ctx: Ctx): RowOrg | null {
  const key = mcellText(cellOf(row, spec.keyCol), shared);
  if (!key) {
    if (!rowHasConstant(row)) return null;
    warn(ctx, `Biểu ${spec.label}: có dòng chứa nội dung nhưng thiếu mã ở cột ${spec.keyCol} → giữ ở cả hai file`);
    return "BOTH";
  }
  if (spec.keyType === "VTVL") {
    const o = ctx.vtvl.get(key);
    if (o) return o;
    warn(ctx, `Biểu ${spec.label}: mã VTVL "${key}" không có trong biểu C1 → giữ ở cả hai file`);
    return "BOTH";
  }
  if (spec.keyType === "CV") {
    const o = ctx.cv.get(key);
    if (o) return o;
    const viaVtvl = spec.vtvlCol ? ctx.vtvl.get(mcellText(cellOf(row, spec.vtvlCol), shared)) : undefined;
    if (viaVtvl) return viaVtvl;
    warn(ctx, `Biểu ${spec.label}: mã công việc "${key}" không có trong biểu A2 → giữ ở cả hai file`);
    return "BOTH";
  }
  // Nhiệm vụ: chỉ thuộc một khối nếu mọi công việc của nó thuộc khối đó; dùng chung/chưa gắn -> cả hai
  return ctx.nv.get(key) ?? "BOTH";
}

function patchWorkbookXml(xml: string): string {
  const calc = xml.match(/<calcPr\b[^>]*\/>/);
  if (calc) return xml.replace(calc[0], setAttr(calc[0], "fullCalcOnLoad", "1"));
  const tag = '<calcPr fullCalcOnLoad="1"/>';
  for (const anchor of ["</definedNames>", "</externalReferences>", "</functionGroups>", "</sheets>"]) {
    const i = xml.indexOf(anchor);
    if (i >= 0) return xml.slice(0, i + anchor.length) + tag + xml.slice(i + anchor.length);
  }
  return xml;
}

/**
 * @param mapping mã VTVL -> khối (đã được người dùng xác nhận). Mã thiếu trong mapping dùng gợi ý tự động.
 */
export async function splitWorkbook(buf: Buffer, mapping: Record<string, Org>): Promise<SplitResult> {
  const wb = await openWorkbook(buf);
  const inspected = await inspectWorkbook(buf);
  const ctx: Ctx = { vtvl: new Map(), cv: new Map(), nv: new Map(), warnings: new Map() };
  for (const p of inspected.positions) {
    const o = mapping[p.code];
    ctx.vtvl.set(p.code, o === "HDND" || o === "UBND" ? o : p.suggested);
  }
  const nH = [...ctx.vtvl.values()].filter((o) => o === "HDND").length;
  const nU = ctx.vtvl.size - nH;
  if (nH === 0) throw new SplitError("ONE_SIDE_EMPTY", "Không có vị trí nào thuộc HĐND nên không có gì để tách.");
  if (nU === 0) throw new SplitError("ONE_SIDE_EMPTY", "Không có vị trí nào thuộc UBND nên không có gì để tách.");

  // A2: mã công việc -> khối; mã nhiệm vụ -> khối (hoặc cả hai)
  const a2 = sheetByRole(wb, "A2");
  if (a2) {
    for (const r of await loadRows(wb, a2)) {
      const vt = mcellText(cellOf(r, "C"), wb.shared);
      const cv = mcellText(cellOf(r, "A"), wb.shared);
      const nv = mcellText(cellOf(r, "B"), wb.shared);
      const o: RowOrg = ctx.vtvl.get(vt) ?? "BOTH";
      if (cv) ctx.cv.set(cv, ctx.cv.has(cv) && ctx.cv.get(cv) !== o ? "BOTH" : o);
      if (nv) ctx.nv.set(nv, ctx.nv.has(nv) && ctx.nv.get(nv) !== o ? "BOTH" : o);
    }
  }

  const out: Record<Org, OutEntry[]> = { HDND: [], UBND: [] };
  const report: SplitReport = { positions: { hdnd: nH, ubnd: nU }, sheets: [], warnings: [] };
  const sheetByPath = new Map(wb.sheets.filter((s) => s.spec).map((s) => [s.path, s]));
  const hasCalcChain = wb.byName.has("xl/calcChain.xml");
  // Một vùng đệm dùng chung để giải nén lần lượt từng sheet (sheet lớn nhất ~80 MB):
  // bộ nhớ dùng ổn định thay vì cấp phát mới cho mỗi sheet rồi chờ thu hồi.
  const scratch = Buffer.allocUnsafe(Math.max(1, ...wb.entries.filter((x) => sheetByPath.has(x.name)).map((x) => x.uncompressedSize)));

  for (const e of wb.entries) {
    if (e.name === "xl/calcChain.xml") continue;
    let shared: OutEntry | null = null;
    if (e.name === "xl/workbook.xml") {
      shared = await deflated(e, Buffer.from(patchWorkbookXml(await readText(wb, e.name)), "utf8"));
    } else if (hasCalcChain && e.name === "[Content_Types].xml") {
      const xml = (await readText(wb, e.name)).replace(/<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, "");
      shared = await deflated(e, Buffer.from(xml, "utf8"));
    } else if (hasCalcChain && e.name === "xl/_rels/workbook.xml.rels") {
      const xml = (await readText(wb, e.name)).replace(/<Relationship\b[^>]*Target="\/?(?:xl\/)?calcChain\.xml"[^>]*\/>/, "");
      shared = await deflated(e, Buffer.from(xml, "utf8"));
    }
    if (shared) {
      out.HDND.push(shared);
      out.UBND.push(shared);
      continue;
    }

    const sheet = sheetByPath.get(e.name);
    if (!sheet) {
      const p = passthrough(e);
      out.HDND.push(p);
      out.UBND.push(p);
      continue;
    }

    // LƯU Ý: model tham chiếu trực tiếp vào `scratch`, chỉ dùng trong vòng lặp này (trước khi giải nén sheet kế tiếp)
    const raw = await readBuffer(wb, e.name, scratch);
    const model = wrap(() => buildSheetModel(raw, sheet.spec!.keyCol));
    if (!model) {
      const p = passthrough(e);
      out.HDND.push(p);
      out.UBND.push(p);
      continue;
    }
    const orgs = new Map<number, RowOrg | null>();
    const rep: SheetReport = { sheet: sheet.name, total: 0, hdnd: 0, ubnd: 0, both: 0 };
    for (const r of model.rows) {
      const o = classifyRow(r, sheet.spec!, wb.shared, ctx);
      orgs.set(r.num, o);
      if (o) {
        rep.total++;
        if (o === "HDND") rep.hdnd++;
        else if (o === "UBND") rep.ubnd++;
        else rep.both++;
      }
    }
    report.sheets.push(rep);
    for (const org of ["HDND", "UBND"] as Org[]) {
      const rendered = renderSheet(model, (r) => {
        const o = orgs.get(r.num);
        return o === org || o === "BOTH";
      });
      out[org].push(await deflatedChunks(e, rendered.chunks));
    }
  }

  report.warnings = [...ctx.warnings.entries()].map(([m, n]) => (n > 1 ? `${m} (${n} dòng)` : m)).slice(0, 50);
  return { files: { HDND: wrap(() => writeZip(out.HDND)), UBND: wrap(() => writeZip(out.UBND)) }, report };
}

export { DATA_START_ROW };
