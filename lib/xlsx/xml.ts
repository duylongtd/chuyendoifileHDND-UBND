// Tiện ích thao tác trực tiếp trên XML của file .xlsx (SpreadsheetML).
// Không dựng cây DOM: bộ biểu 30.000 dòng giải nén ra ~400 MB XML, chỉ xử lý bằng chuỗi/regex theo từng dòng.

export function xmlUnescape(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&amp;/g, "&");
}

export function xmlEscape(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function colToIndex(col: string): number {
  let n = 0;
  for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
  return n;
}

export function indexToCol(n: number): string {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function getAttr(tag: string, name: string): string | undefined {
  const m = tag.match(new RegExp(`\\s${name}="([^"]*)"`));
  return m ? m[1] : undefined;
}

export function setAttr(tag: string, name: string, value: string): string {
  const re = new RegExp(`\\s${name}="[^"]*"`);
  if (re.test(tag)) return tag.replace(re, ` ${name}="${value}"`);
  return tag.replace(/^(<\w+)/, `$1 ${name}="${value}"`);
}

export function removeAttr(tag: string, name: string): string {
  return tag.replace(new RegExp(`\\s${name}="[^"]*"`), "");
}

/** Danh sách chuỗi dùng chung (xl/sharedStrings.xml) -> mảng text thuần. */
export function parseSharedStrings(xml: string): string[] {
  const out: string[] = [];
  const re = /<si>([\s\S]*?)<\/si>|<si\/>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    const inner = m[1] ?? "";
    let text = "";
    const tre = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g;
    let t: RegExpExecArray | null;
    // Bỏ qua phần phiên âm (rPh) nếu có
    const clean = inner.replace(/<rPh\b[\s\S]*?<\/rPh>/g, "");
    while ((t = tre.exec(clean))) text += t[1] ?? "";
    out.push(xmlUnescape(text));
  }
  return out;
}

/**
 * XML của sheet được xử lý ở dạng chuỗi "latin1" (mỗi byte = 1 ký tự) để tiết kiệm RAM:
 * cấu trúc XML toàn ASCII, phần chữ tiếng Việt (UTF-8) đi qua nguyên vẹn. Khi cần ĐỌC nội dung chữ
 * (mã, tên vị trí) thì giải mã lại bằng hàm này.
 */
export function latin1ToUtf8(s: string): string {
  return /[\u0080-\u00ff]/.test(s) ? Buffer.from(s, "latin1").toString("utf8") : s;
}

/**
 * Bỏ escape cho công thức nhưng GIỮ NGUYÊN tham chiếu ký tự dạng số (&#10; ...) dưới dạng dấu giữ chỗ,
 * để khi ghi lại không làm hỏng (chuỗi latin1 không chứa được ký tự > 255).
 */
export function unescapeFormula(s: string): string {
  if (!s.includes("&")) return s;
  return s.replace(/&(lt|gt|quot|apos|amp|#\d+|#x[0-9a-fA-F]+);/g, (_, n: string) => {
    switch (n) {
      case "lt":
        return "<";
      case "gt":
        return ">";
      case "quot":
        return '"';
      case "apos":
        return "'";
      case "amp":
        return "&";
      default:
        return `\u0000${n}\u0000`;
    }
  });
}

export function escapeFormula(s: string): string {
  const e = s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  return e.includes("\u0000") ? e.replace(/\u0000(#[^\u0000]+)\u0000/g, "&$1;") : e;
}
