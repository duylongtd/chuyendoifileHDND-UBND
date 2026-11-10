// Mô hình hóa vùng dữ liệu của một sheet để dồn dòng mà vẫn giữ công thức/định dạng.
// LƯU Ý: chuỗi XML ở đây là dạng latin1 (1 byte = 1 ký tự = 1 vị trí byte), xem latin1ToUtf8 trong xml.ts.
//
// Bộ biểu VTVL điền sẵn công thức cho 30.000 dòng (sheet tới ~80 MB XML). Ta chỉ phân tích
// vùng có dữ liệu (dòng 3 .. dòng cuối có nội dung); phần "đuôi" (các dòng mẫu trống) giữ nguyên dạng chuỗi.
import { DATA_START_ROW } from "./classify";
import { translateFormula } from "./formula";
import { colToIndex, escapeFormula, getAttr, indexToCol, latin1ToUtf8, removeAttr, setAttr, unescapeFormula, xmlUnescape } from "./xml";

export class SheetFormatError extends Error {}

interface FInfo {
  kind: "normal" | "array" | "raw";
  /** Công thức đã unescape (với raw: nguyên thẻ <f>) */
  text: string;
  /** Thuộc tính còn lại của thẻ <f> (đã bỏ t/si/ref với công thức thường) */
  attrs: string;
}

export interface MCell {
  col: string;
  colIdx: number;
  /** Phần thuộc tính của thẻ <c ...> */
  attrs: string;
  f?: FInfo;
  /** Nội dung sau thẻ <f> (thường là <v>…</v>), hoặc toàn bộ nội dung nếu không có công thức */
  rest: string;
}

export interface MRow {
  num: number;
  attrs: string;
  cells: MCell[];
}

interface SharedGroup {
  si: string;
  row: number;
  colIdx: number;
  text: string;
  endRow: number;
  endCol: string;
  singleColumn: boolean;
  inRegion: boolean;
  remastered?: boolean;
}

export interface SheetModel {
  /** Phần đầu sheet tới hết 2 dòng tiêu đề: cắt thẳng từ vùng đệm gốc (không sao chép) */
  head: Buffer;
  rows: MRow[];
  /** Phần đuôi (các dòng mẫu sau vùng dữ liệu) dưới dạng byte, không dựng thành chuỗi */
  tail: Buffer[];
  lastData: number;
}

interface El {
  start: number;
  end: number;
  /** Phần thuộc tính của thẻ mở (kể cả khoảng trắng đầu) */
  attrs: string;
  inner: string;
}

/**
 * Duyệt lần lượt các phần tử <tag ...>...</tag> (hoặc <tag .../>) cùng cấp trong chuỗi XML.
 * Dùng indexOf nên thời gian luôn tuyến tính theo độ dài chuỗi, kể cả khi XML dị dạng
 * (regex kiểu <c ...>[\s\S]*?</c> sẽ quét lại từ đầu cho mỗi thẻ không đóng -> treo máy chủ).
 * strict = true: gặp thẻ không đóng / lồng nhau thì báo lỗi; false: dừng lại (dùng khi chuỗi là một đoạn cắt dở).
 */
function scanElements(xml: string, tag: "row" | "c", strict: boolean, fn: (el: El) => boolean | void): void {
  const open = `<${tag}`;
  const close = `</${tag}>`;
  const isOpenAt = (i: number) => {
    const ch = xml.charCodeAt(i + open.length);
    return ch === 32 || ch === 62 || ch === 47 || ch === 9 || ch === 10 || ch === 13; // khoảng trắng, ">", "/"
  };
  const nextOpen = (from: number, to: number) => {
    let i = xml.indexOf(open, from);
    while (i >= 0 && i < to && !isOpenAt(i)) i = xml.indexOf(open, i + open.length);
    return i >= 0 && i < to ? i : -1;
  };
  const malformed = () => new SheetFormatError("XML của trang tính bị lỗi (thẻ không đóng hoặc lồng nhau). File có thể đã hỏng; hãy mở bằng Excel, bấm Lưu rồi tải lên lại.");
  let pos = 0;
  for (;;) {
    const i = nextOpen(pos, xml.length);
    if (i < 0) return;
    const gt = xml.indexOf(">", i);
    if (gt < 0) {
      if (strict) throw malformed();
      return;
    }
    let el: El;
    if (xml.charCodeAt(gt - 1) === 47) {
      el = { start: i, end: gt + 1, attrs: xml.slice(i + open.length, gt - 1), inner: "" };
    } else {
      const c = xml.indexOf(close, gt);
      if (c < 0) {
        if (strict) throw malformed();
        return;
      }
      if (nextOpen(gt, c) >= 0) throw malformed();
      el = { start: i, end: c + close.length, attrs: xml.slice(i + open.length, gt), inner: xml.slice(gt + 1, c) };
    }
    if (fn(el) === false) return;
    pos = el.end;
  }
}

const F_RE = /^<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/;

function parseRef(ref: string): { endRow: number; endCol: string; single: boolean } {
  const parts = ref.split(":");
  const a = parts[0].match(/^([A-Z]+)(\d+)$/);
  const b = (parts[1] ?? parts[0]).match(/^([A-Z]+)(\d+)$/);
  if (!a || !b) throw new SheetFormatError(`Vùng công thức không hợp lệ: ${ref}`);
  return { endRow: Number(b[2]), endCol: b[1], single: a[1] === b[1] };
}

function parseRowCells(rowNum: number, inner: string, groups: Map<string, SharedGroup>, inRegion: boolean): MCell[] {
  const cells: MCell[] = [];
  scanElements(inner, "c", true, ({ attrs, inner: body }) => {
    const ref = getAttr(`<c${attrs}>`, "r");
    const rm = ref?.match(/^([A-Z]+)(\d+)$/);
    if (!rm) throw new SheetFormatError("Ô thiếu địa chỉ (thuộc tính r), cấu trúc file không được hỗ trợ.");
    const col = rm[1];
    const colIdx = colToIndex(col);
    const cell: MCell = { col, colIdx, attrs, rest: body };
    const fm = body.startsWith("<f") ? body.match(F_RE) : null;
    if (fm) {
      const fTag = `<f${fm[1]}>`;
      const fAttrs = fm[1];
      const fText = fm[2] ?? "";
      const t = getAttr(fTag, "t");
      cell.rest = body.slice(fm[0].length);
      if (t === "shared") {
        const si = getAttr(fTag, "si") ?? "";
        const sref = getAttr(fTag, "ref");
        let text: string;
        if (fText) {
          text = unescapeFormula(fText);
          if (sref && !groups.has(si)) {
            const r = parseRef(sref);
            groups.set(si, { si, row: rowNum, colIdx, text, endRow: r.endRow, endCol: r.endCol, singleColumn: r.single, inRegion });
          }
        } else {
          const g = groups.get(si);
          if (!g) throw new SheetFormatError("Công thức dùng chung thiếu ô gốc, cấu trúc file không được hỗ trợ.");
          text = translateFormula(g.text, rowNum - g.row, colIdx - g.colIdx);
        }
        let rest = removeAttr(removeAttr(removeAttr(fTag, "t"), "si"), "ref");
        rest = rest.slice(2, -1);
        cell.f = { kind: "normal", text, attrs: rest };
      } else if (t === "array") {
        cell.f = { kind: "array", text: unescapeFormula(fText), attrs: fAttrs };
      } else if (t === "dataTable") {
        cell.f = { kind: "raw", text: fm[0], attrs: "" };
      } else {
        cell.f = { kind: "normal", text: unescapeFormula(fText), attrs: fAttrs };
      }
    }
    cells.push(cell);
  });
  return cells;
}

interface Located {
  lastData: number;
  rowsStart: number;
  regionStart: number;
  regionEnd: number;
  tailStart: number;
}

const LT = 0x3c; // "<"

function rowOfCellAt(buf: Buffer, cellStart: number): { col: string; row: number } | null {
  // cellStart trỏ vào '<c r="'
  const end = buf.indexOf(0x22, cellStart + 6); // dấu " đóng
  if (end < 0 || end - cellStart > 20) return null;
  const m = buf.toString("latin1", cellStart + 6, end).match(/^([A-Z]+)(\d+)$/);
  return m ? { col: m[1], row: Number(m[2]) } : null;
}

/**
 * Dòng cuối cùng có nội dung: có ô nhập tay, hoặc ô mã (kể cả mã do công thức trả về).
 * Dò NGƯỢC từ cuối trên byte (không dựng chuỗi cho cả sheet ~80 MB).
 */
function findLastDataRow(buf: Buffer, keyCol: string, from: number, to: number): number {
  let last = 0;
  // (a) ô nhập tay: thẻ <c ...> theo ngay sau là <v>giá trị hoặc <is>
  for (const pat of ['"><v>', '"><is>']) {
    let pos = to;
    while (pos > from) {
      const hit = buf.lastIndexOf(pat, pos - 1, "latin1");
      if (hit < from) break;
      pos = hit;
      if (pat === '"><v>' && buf[hit + pat.length] === LT) continue; // <v></v> rỗng
      const cs = buf.lastIndexOf('<c r="', hit, "latin1");
      const ref = cs >= from ? rowOfCellAt(buf, cs) : null;
      if (ref) {
        if (ref.row > last) last = ref.row;
        break; // các dòng tăng dần nên lần gặp đầu tiên từ cuối là dòng lớn nhất
      }
    }
  }
  // (b) ô mã ở cột khóa có giá trị (mã do công thức trả về)
  const keyPat = `<c r="${keyCol}`;
  let pos = to;
  let closeLimit = to; // sau vị trí này chắc chắn không còn </c> nào (tránh quét lặp khi XML dị dạng)
  while (pos > from) {
    const cs = buf.lastIndexOf(keyPat, pos - 1, "latin1");
    if (cs < from) break;
    pos = cs;
    const ref = rowOfCellAt(buf, cs);
    if (!ref || ref.col !== keyCol) continue;
    if (ref.row <= last) break;
    const tagEnd = buf.indexOf(0x3e, cs); // ">"
    if (tagEnd < 0 || buf[tagEnd - 1] === 0x2f) continue; // <c .../> rỗng
    const rel = tagEnd < closeLimit ? buf.subarray(tagEnd, closeLimit).indexOf("</c>", 0, "latin1") : -1;
    if (rel < 0) {
      closeLimit = cs;
      continue;
    }
    const cellEnd = tagEnd + rel;
    // Chỉ tìm <v> TRONG phạm vi ô này (tìm không giới hạn sẽ quét tới tận ô có giá trị kế tiếp, rất chậm)
    const body = buf.toString("latin1", tagEnd + 1, cellEnd);
    const v = body.match(/<v>([^<]+)<\/v>/);
    if (v && v[1].trim()) {
      last = ref.row;
      break;
    }
  }
  return last;
}

function rowNumberAt(buf: Buffer, rowStart: number): number {
  const tagEnd = buf.indexOf(0x3e, rowStart);
  const r = tagEnd > 0 ? getAttr(buf.toString("latin1", rowStart, tagEnd + 1), "r") : undefined;
  const num = Number(r);
  if (!Number.isInteger(num) || num < 1) throw new SheetFormatError("Dòng thiếu số thứ tự (thuộc tính r), cấu trúc file không được hỗ trợ.");
  return num;
}

/** Dò vị trí (byte) của vùng dữ liệu trong XML của sheet. */
function locate(buf: Buffer, keyCol: string): Located | null {
  const sd = buf.indexOf("<sheetData", 0, "latin1");
  if (sd < 0) return null;
  const sdEnd = buf.indexOf("</sheetData>", sd, "latin1");
  if (sdEnd < 0) return null; // <sheetData/> rỗng
  // Bộ dò dựa vào dạng <c r="A1" ...> (Excel, openpyxl, LibreOffice đều ghi r trước).
  // Nếu file ghi thuộc tính theo thứ tự khác thì dừng, không âm thầm bỏ qua việc tách sheet này.
  const firstCell = buf.indexOf("<c ", sd, "latin1");
  if (firstCell >= 0 && firstCell < sdEnd && buf.toString("latin1", firstCell, firstCell + 6) !== '<c r="') {
    throw new SheetFormatError("File do phần mềm khác lưu theo định dạng chưa hỗ trợ. Hãy mở bằng Excel, bấm Lưu rồi tải lên lại.");
  }
  const lastData = findLastDataRow(buf, keyCol, sd, sdEnd);
  if (lastData < DATA_START_ROW) return null;

  const rowsStart = buf.indexOf("<row ", sd, "latin1");
  if (rowsStart < 0 || rowsStart > sdEnd) return null;
  let regionStart = -1;
  let tailStart = -1;
  let prev = 0;
  for (let pos = rowsStart; pos >= 0 && pos < sdEnd; pos = buf.indexOf("<row ", pos + 5, "latin1")) {
    const num = rowNumberAt(buf, pos);
    if (num <= prev) throw new SheetFormatError("Thứ tự dòng trong sheet không hợp lệ, cấu trúc file không được hỗ trợ.");
    prev = num;
    if (num > lastData) {
      tailStart = pos;
      break;
    }
    if (num >= DATA_START_ROW && regionStart < 0) regionStart = pos;
  }
  if (regionStart < 0) return null;
  if (tailStart < 0) tailStart = sdEnd;
  return { lastData, rowsStart, regionStart, regionEnd: tailStart, tailStart };
}

function eachRow(xml: string, fn: (num: number, attrs: string, inner: string) => void): void {
  scanElements(xml, "row", true, ({ attrs, inner }) => {
    const num = Number(getAttr(`<row${attrs}>`, "r"));
    if (!Number.isInteger(num)) throw new SheetFormatError("Dòng thiếu số thứ tự (thuộc tính r), cấu trúc file không được hỗ trợ.");
    fn(num, attrs, inner);
  });
}

/**
 * @param buf XML của sheet (đã giải nén). Trả về null nếu sheet không có dòng dữ liệu nào (không cần xử lý).
 */
export function buildSheetModel(buf: Buffer, keyCol: string): SheetModel | null {
  const loc = locate(buf, keyCol);
  if (!loc) return null;

  const groups = new Map<string, SharedGroup>();
  // Dòng tiêu đề (1-2): chỉ cần ghi nhận ô gốc của công thức dùng chung (nếu có)
  eachRow(buf.toString("latin1", loc.rowsStart, loc.regionStart), (num, _attrs, inner) => {
    if (inner.includes('t="shared"')) parseRowCells(num, inner, groups, false);
  });
  const rows: MRow[] = [];
  eachRow(buf.toString("latin1", loc.regionStart, loc.regionEnd), (num, attrs, inner) => {
    rows.push({ num, attrs, cells: parseRowCells(num, inner, groups, true) });
  });
  return { head: buf.subarray(0, loc.regionStart), rows, tail: fixTail(buf.subarray(loc.tailStart), groups, loc.lastData), lastData: loc.lastData };
}

/**
 * Các nhóm công thức dùng chung có ô gốc nằm trong vùng dữ liệu (đã bị viết lại thành công thức thường)
 * nhưng còn ô phụ thuộc ở phần đuôi: chuyển ô phụ thuộc đầu tiên ở đuôi thành ô gốc mới.
 * Chỉ giải mã một đoạn đầu của đuôi (thường ô cần sửa nằm ngay dòng đầu tiên); phần còn lại giữ nguyên byte.
 */
function fixTail(tail: Buffer, groups: Map<string, SharedGroup>, lastData: number): Buffer[] {
  const pending = new Map<string, SharedGroup>();
  let maxEnd = 0;
  for (const g of groups.values()) {
    if (g.inRegion && g.endRow > lastData) {
      pending.set(g.si, g);
      if (g.endRow > maxEnd) maxEnd = g.endRow;
    }
  }
  if (!pending.size || !tail.length) return [tail];
  const onlySingle = [...pending.values()].every((g) => g.singleColumn);

  for (let window = 1 << 20; ; window *= 4) {
    const whole = window >= tail.length;
    const text = tail.toString("latin1", 0, Math.min(window, tail.length));
    for (const g of pending.values()) g.remastered = false;
    const out: string[] = [];
    let consumed = 0;
    let open = pending.size;
    let done = false;
    // Trả về XML mới của ô nếu ô này là ô phụ thuộc của một nhóm đang chờ; null nếu giữ nguyên
    const remasterCell = (attrs: string, body: string): string | null => {
      if (!body.startsWith("<f")) return null;
      const fm = body.match(F_RE);
      if (!fm || (fm[2] ?? "") !== "") return null;
      const fTag = `<f${fm[1]}>`;
      if (getAttr(fTag, "t") !== "shared") return null;
      const g = pending.get(getAttr(fTag, "si") ?? "");
      if (!g) return null;
      const rm = (getAttr(`<c${attrs}>`, "r") ?? "").match(/^([A-Z]+)(\d+)$/);
      if (!rm) return null;
      const formula = escapeFormula(translateFormula(g.text, Number(rm[2]) - g.row, colToIndex(rm[1]) - g.colIdx));
      const rest = body.slice(fm[0].length);
      if (g.singleColumn) {
        if (g.remastered) return null;
        g.remastered = true;
        open--;
        const newF = setAttr(fTag, "ref", `${rm[1]}${rm[2]}:${g.endCol}${g.endRow}`).slice(0, -1);
        return `<c${attrs}>${newF}>${formula}</f>${rest}</c>`;
      }
      const plain = removeAttr(removeAttr(removeAttr(fTag, "t"), "si"), "ref").slice(0, -1);
      return `<c${attrs}>${plain}>${formula}</f>${rest}</c>`;
    };
    scanElements(text, "row", false, (row) => {
      const num = Number(getAttr(`<row${row.attrs}>`, "r"));
      if (num > maxEnd) {
        done = true;
        return false;
      }
      let rowStr = text.slice(row.start, row.end);
      if (rowStr.includes('t="shared"')) {
        const pieces: string[] = [];
        let last = 0;
        scanElements(rowStr, "c", true, (c) => {
          const fixed = remasterCell(c.attrs, c.inner);
          if (fixed !== null) {
            pieces.push(rowStr.slice(last, c.start), fixed);
            last = c.end;
          }
        });
        if (pieces.length) rowStr = pieces.join("") + rowStr.slice(last);
      }
      out.push(text.slice(consumed, row.start), rowStr);
      consumed = row.end;
      if (open === 0 && onlySingle) {
        done = true;
        return false;
      }
    });
    if (done || whole) return [Buffer.from(out.join(""), "latin1"), tail.subarray(consumed)];
  }
}

/** Nội dung chữ của ô (đã giải mã UTF-8). Dùng để đọc mã VTVL / mã công việc / tên vị trí. */
export function mcellText(cell: MCell | undefined, shared: string[]): string {
  if (!cell) return "";
  const t = getAttr(`<c${cell.attrs}>`, "t");
  if (t === "inlineStr") {
    const parts = [...cell.rest.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((x) => x[1]);
    return xmlUnescape(latin1ToUtf8(parts.join(""))).trim();
  }
  const v = cell.rest.match(/<v>([\s\S]*?)<\/v>/);
  if (!v) return "";
  if (t === "s") return (shared[Number(v[1])] ?? "").trim();
  return xmlUnescape(latin1ToUtf8(v[1])).trim();
}

export function cellOf(row: MRow, col: string): MCell | undefined {
  return row.cells.find((c) => c.col === col);
}

export function rowHasConstant(row: MRow): boolean {
  return row.cells.some((c) => !c.f && (/<v>[^<]/.test(c.rest) || c.rest.includes("<is>")));
}

function emitCell(c: MCell, srcRow: number, destRow: number, blank: boolean): string {
  let attrs = setAttr(`<c${c.attrs}>`, "r", `${c.col}${destRow}`).slice(2, -1);
  const delta = destRow - srcRow;
  if (c.f) {
    if (c.f.kind === "raw") return `<c${attrs}>${c.f.text}${blank ? "" : c.rest}</c>`;
    let fAttrs = c.f.attrs;
    if (c.f.kind === "array" && delta !== 0) {
      const ref = getAttr(`<f${fAttrs}>`, "ref");
      if (ref) fAttrs = setAttr(`<f${fAttrs}>`, "ref", translateFormula(ref, delta, 0)).slice(2, -1);
    }
    const text = escapeFormula(translateFormula(c.f.text, delta, 0));
    if (blank) attrs = removeAttr(`<c${attrs}>`, "t").slice(2, -1);
    return `<c${attrs}><f${fAttrs}>${text}</f>${blank ? "" : c.rest}</c>`;
  }
  if (blank) {
    if (!c.rest) return `<c${attrs}/>`;
    attrs = removeAttr(`<c${attrs}>`, "t").slice(2, -1);
    return `<c${attrs}/>`;
  }
  return c.rest ? `<c${attrs}>${c.rest}</c>` : `<c${attrs}/>`;
}

function emitRow(row: MRow, destRow: number, blank: boolean): string {
  let tag = setAttr(`<row${row.attrs}>`, "r", String(destRow));
  if (blank) tag = removeAttr(removeAttr(tag, "ht"), "customHeight");
  const cells = row.cells.map((c) => emitCell(c, row.num, destRow, blank)).join("");
  return cells ? `${tag}${cells}</row>` : `${tag.slice(0, -1)}/>`;
}

/**
 * Dựng lại XML của sheet (dạng các khúc byte nối tiếp): các dòng được giữ dồn lên liền nhau từ dòng 3
 * (giữ thứ tự gốc), các vị trí dòng còn lại trong vùng dữ liệu trở thành dòng trống
 * (giữ định dạng + công thức mẫu), phần đuôi giữ nguyên.
 */
export function renderSheet(model: SheetModel, keep: (row: MRow) => boolean): { chunks: Buffer[]; kept: number } {
  const parts: string[] = [];
  let dest = DATA_START_ROW;
  for (const row of model.rows) {
    if (keep(row)) parts.push(emitRow(row, dest++, false));
  }
  const kept = dest - DATA_START_ROW;
  for (const row of model.rows) {
    if (row.num >= dest) parts.push(emitRow(row, row.num, true));
  }
  return { chunks: [model.head, Buffer.from(parts.join(""), "latin1"), ...model.tail], kept };
}

export { indexToCol };
