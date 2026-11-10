// Dịch tham chiếu tương đối trong công thức khi chuyển ô sang vị trí khác
// (giống Excel khi sao chép công thức): A5 -> A7 nếu dịch +2 dòng; $A$5 giữ nguyên.
import { colToIndex, indexToCol } from "./xml";

// Tham chiếu ô: tùy chọn tên sheet đã được bỏ qua ở bước tách chuỗi; ở đây chỉ bắt A1-style.
const REF_RE = /(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![\w(!])/g;

/**
 * Dịch các tham chiếu tương đối trong công thức đi dRow dòng, dCol cột.
 * Bỏ qua chuỗi trong ngoặc kép và tên sheet trong ngoặc đơn.
 * Lưu ý: công thức trong XML đã được escape (&quot; ...) -> gọi trên chuỗi đã unescape rồi escape lại.
 */
export function translateFormula(formula: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) return formula;
  let out = "";
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i];
    if (ch === '"' || ch === "'") {
      // chuỗi hoặc tên sheet: sao chép nguyên tới dấu đóng (xử lý "" / '' lồng)
      let j = i + 1;
      while (j < formula.length) {
        if (formula[j] === ch) {
          if (formula[j + 1] === ch) {
            j += 2;
            continue;
          }
          break;
        }
        j++;
      }
      out += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    // đoạn bình thường tới ký tự trích dẫn kế tiếp
    let j = i;
    while (j < formula.length && formula[j] !== '"' && formula[j] !== "'") j++;
    const seg = formula.slice(i, j);
    out += seg.replace(REF_RE, (m, cAbs: string, col: string, rAbs: string, row: string, offset: number, whole: string) => {
      const prev = offset > 0 ? whole[offset - 1] : "";
      // không phải tham chiếu nếu dính liền chữ/số/dấu chấm (tên hàm, tên định danh, số thập phân)
      if (prev && /[A-Za-z0-9_.]/.test(prev)) return m;
      let c = col;
      let r = Number(row);
      if (!cAbs && dCol) c = indexToCol(Math.max(1, colToIndex(col) + dCol));
      if (!rAbs && dRow) r = Math.max(1, r + dRow);
      return `${cAbs}${c}${rAbs}${r}`;
    });
    i = j;
  }
  return out;
}
