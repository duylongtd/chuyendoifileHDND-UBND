// Tự kiểm các quy tắc cốt lõi:  npm test
import assert from "node:assert/strict";
import { outputFileName, suggestOrg } from "../lib/xlsx/classify";
import { translateFormula } from "../lib/xlsx/formula";
import { escapeFormula, unescapeFormula } from "../lib/xlsx/xml";

let n = 0;
const eq = (a: unknown, b: unknown, msg = "") => {
  if (msg) assert.deepEqual(a, b, msg);
  else assert.deepEqual(a, b);
  n++;
};

// --- Dịch công thức khi dồn dòng (từ dòng 8 lên dòng 3 => -5) ---
eq(translateFormula("C8*P8/60", -5, 0), "C3*P3/60", "tham chiếu cùng dòng dịch theo, hằng số giữ nguyên");
eq(translateFormula("IF(A8=\"\",\"\",$A$3:$A$30002)", -5, 0), "IF(A3=\"\",\"\",$A$3:$A$30002)", "tham chiếu tuyệt đối giữ nguyên");
eq(translateFormula("IF(G8=\"\",\"C3 CHƯA ĐƯỢC XÁC NHẬN\",\"PP3: THIẾU\")", -5, 0), "IF(G3=\"\",\"C3 CHƯA ĐƯỢC XÁC NHẬN\",\"PP3: THIẾU\")", "chữ trong chuỗi không bị đổi");
eq(translateFormula("INDEX('2. A2_CongViec'!$C$3:$C$30002,MATCH(A8,'2. A2_CongViec'!$A$3:$A$30002,0))", -5, 0), "INDEX('2. A2_CongViec'!$C$3:$C$30002,MATCH(A3,'2. A2_CongViec'!$A$3:$A$30002,0))", "tên sheet trong nháy đơn không bị đổi");
eq(translateFormula("'4. B2_CT_TacDong'!A8", -5, 0), "'4. B2_CT_TacDong'!A3", "tham chiếu sang sheet khác cùng dòng dịch theo");
eq(translateFormula("DM_THAM_SO!$B$3+LOG10(B8)+QUY_GIO_PP2", -5, 0), "DM_THAM_SO!$B$3+LOG10(B3)+QUY_GIO_PP2", "tên hàm/tên vùng không bị coi là ô");
eq(translateFormula("DM1!A8+A$8+$A8", -5, 0), "DM1!A3+A$8+$A3", "tên sheet dạng chữ+số; $ khóa dòng");
eq(translateFormula("SUM(A8:C8)", 2, 0), "SUM(A10:C10)");
eq(translateFormula("A8", 0, 0), "A8");
eq(escapeFormula(unescapeFormula("IF(A3&lt;&gt;&quot;x&quot;,B3&amp;&#10;C3,1)")), "IF(A3&lt;&gt;\"x\",B3&amp;&#10;C3,1)", "escape giữ nguyên ký tự dạng số");

// --- Nhận diện 4 chức danh HĐND ---
const org = (name: string, hint = "") => {
  const s = suggestOrg(name, hint);
  return `${s.org}${s.confident ? "" : "?"}`;
};
eq(org("Chủ tịch HĐND phường"), "HDND");
eq(org("Phó Chủ tịch HĐND phường"), "HDND");
eq(org("Phó Chủ tịch Hội đồng nhân dân xã"), "HDND");
eq(org("Trưởng ban Hội đồng"), "HDND");
eq(org("Phó Trưởng ban Kinh tế - Ngân sách"), "HDND");
eq(org("Phó trưởng ban hội đồng"), "HDND");
eq(org("Chủ tịch UBND phường"), "UBND");
eq(org("Phó Chủ tịch UBND phường"), "UBND");
eq(org("Chánh Văn phòng HĐND và UBND phường"), "UBND");
eq(org("Phó Chánh Văn phòng HĐND và UBND phường"), "UBND");
eq(org("Chuyên viên tham mưu, giúp việc HĐND"), "UBND");
eq(org("Trưởng phòng"), "UBND");
eq(org("Trưởng Ban Chỉ huy quân sự"), "UBND");
eq(org(" ", "Giúp Trưởng ban tổ chức thực hiện nhiệm vụ"), "HDND?", "tên trống: gợi ý từ C2, cần xác nhận");
eq(org("", "Tham mưu công tác văn phòng"), "UBND?", "tên trống, không đủ căn cứ: cần xác nhận");
eq(org("", ""), "UBND?");

// --- Tên file kết quả ---
eq(outputFileName("01.HĐND-UBND Phường Hà Huy Tập.xlsx", "HDND"), "01.HĐND Phường Hà Huy Tập.xlsx");
eq(outputFileName("01.HĐND-UBND Phường Hà Huy Tập.xlsx", "UBND"), "01.UBND Phường Hà Huy Tập.xlsx");
eq(outputFileName("01.HĐND-UBND Phường Hà Huy Tập.xlsx".normalize("NFD"), "HDND"), "01.HĐND Phường Hà Huy Tập.xlsx", "tên file kiểu macOS (NFD)");
eq(outputFileName("HDND_UBND xa A.XLSX", "UBND"), "UBND xa A.xlsx");
eq(outputFileName("UBND và HĐND xã B.xlsx", "HDND"), "HĐND xã B.xlsx");
eq(outputFileName("Bo bieu xa C.xlsx", "HDND"), "Bo bieu xa C - HĐND.xlsx");
eq(outputFileName('a/b:c*"?.xlsx', "UBND"), "a_b_c___ - UBND.xlsx", "ký tự không hợp lệ trong tên file");

console.log(`Tự kiểm: ${n}/${n} đạt`);
