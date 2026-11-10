// Quy tắc nghiệp vụ: vị trí nào thuộc HĐND, vị trí nào thuộc UBND; biểu nào tách theo mã nào.

export type Org = "HDND" | "UBND";
export type RowOrg = Org | "BOTH";

export const ORG_LABEL: Record<Org, string> = { HDND: "HĐND", UBND: "UBND" };

export function normalizeText(s: string): string {
  return s
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

export interface Suggestion {
  org: Org;
  /** false = cần người dùng xác nhận (tên trống / không rõ) */
  confident: boolean;
  /** true = xếp UBND nhưng tên có nhắc HĐND: luôn hiện trong bảng xác nhận để người dùng soát */
  related?: boolean;
  reason: string;
}

const HDND_WORD = /\b(hdnd|hoi dong nhan dan)\b/;
const CT_HDND = /\b(pho )?chu tich (hdnd|hoi dong nhan dan)\b/;
const CT_UBND = /\b(pho )?chu tich (ubnd|uy ban nhan dan)\b/;
/** Vị trí tham mưu, giúp việc: dù tên có nhắc "Chủ tịch HĐND" cũng không phải chức danh HĐND */
const STAFF = /^(chuyen vien|can su|nhan vien|van thu|ke toan|thu ky|giup viec|tro ly)\b/;

/**
 * HĐND gồm đúng 4 chức danh: Chủ tịch HĐND, Phó Chủ tịch HĐND, Trưởng ban (của HĐND), Phó Trưởng ban (của HĐND).
 * Mọi vị trí khác (kể cả Văn phòng HĐND và UBND, chuyên viên giúp việc HĐND) thuộc UBND.
 */
export function suggestOrg(name: string, hint = ""): Suggestion {
  const n = normalizeText(name);
  if (n) {
    const staff = STAFF.test(n);
    if (!staff && CT_HDND.test(n)) {
      if (CT_UBND.test(n)) return { org: "UBND", confident: false, reason: "Tên vị trí ghi cả Chủ tịch HĐND và Chủ tịch UBND" };
      // gồm cả dạng kiêm nhiệm: "Bí thư Đảng ủy, Chủ tịch HĐND xã"
      return { org: "HDND", confident: true, reason: "Chức danh Chủ tịch/Phó Chủ tịch HĐND" };
    }
    if (/^(pho )?truong ban\b/.test(n) && !/\bban (chi huy|quan ly|chi dao)\b/.test(n)) {
      return { org: "HDND", confident: true, reason: "Chức danh Trưởng ban/Phó Trưởng ban của HĐND" };
    }
    if (HDND_WORD.test(n)) {
      return { org: "UBND", confident: true, related: true, reason: "Tên có nhắc HĐND nhưng không thuộc 4 chức danh HĐND nên xếp UBND" };
    }
    return { org: "UBND", confident: true, reason: "Không thuộc 4 chức danh HĐND" };
  }
  // Tên trống: gợi ý từ mô tả trách nhiệm (biểu C2), luôn cần xác nhận
  const h = normalizeText(hint);
  if (/\b(giup|pho) truong ban\b/.test(h) || /^truong ban\b/.test(h)) {
    return { org: "HDND", confident: false, reason: "Tên vị trí để trống; mô tả ở biểu C2 nhắc tới Trưởng ban" };
  }
  if (/^(pho )?chu tich (hdnd|hoi dong nhan dan)\b/.test(h)) {
    return { org: "HDND", confident: false, reason: "Tên vị trí để trống; mô tả ở biểu C2 nhắc tới Chủ tịch HĐND" };
  }
  return { org: "UBND", confident: false, reason: "Tên vị trí để trống, không đủ căn cứ nhận diện" };
}

export type SheetRole = "A1" | "A2" | "B1" | "B2CT" | "B2TD" | "C1" | "C2" | "C3" | "D1" | "D2" | "DUBAO";

export interface SheetSpec {
  role: SheetRole;
  label: string;
  match: RegExp;
  /** Cột chứa mã dùng để phân loại dòng */
  keyCol: string;
  /** Mã ở keyCol là loại gì */
  keyType: "VTVL" | "CV" | "NV";
  /** Cột mã VTVL dự phòng (khi mã công việc không có ở A2) */
  vtvlCol?: string;
  /** Cột "Cơ quan" (biểu A1): dùng khi nhiệm vụ chưa gắn công việc nào ở A2 */
  orgCol?: string;
}

/**
 * Khối ghi ở cột "Cơ quan": chỉ nhắc HĐND -> HĐND; còn lại (UBND, để trống, ghi cả hai) -> UBND,
 * vì UBND là khối nhận mọi thứ không thuộc 4 chức danh HĐND.
 */
export function orgFromAgency(text: string): Org {
  const n = normalizeText(text);
  const h = /\b(hdnd|hoi dong nhan dan)\b/.test(n);
  const u = /\b(ubnd|uy ban nhan dan)\b/.test(n);
  return h && !u ? "HDND" : "UBND";
}

/** Dữ liệu bắt đầu từ dòng 3 (dòng 1: tên biểu, dòng 2: tiêu đề cột). */
export const DATA_START_ROW = 3;

export const SHEET_SPECS: SheetSpec[] = [
  { role: "A1", label: "1. A1_NhiemVu", match: /A1_NhiemVu/i, keyCol: "A", keyType: "NV", orgCol: "C" },
  { role: "A2", label: "2. A2_CongViec", match: /A2_CongViec/i, keyCol: "C", keyType: "VTVL" },
  { role: "B1", label: "3. B1_KhoiLuong", match: /B1_KhoiLuong/i, keyCol: "A", keyType: "CV", vtvlCol: "B" },
  { role: "B2CT", label: "4. B2_CT_TacDong", match: /B2_CT/i, keyCol: "A", keyType: "CV", vtvlCol: "B" },
  { role: "B2TD", label: "5. B2_ThayDoi", match: /B2_ThayDoi/i, keyCol: "A", keyType: "CV", vtvlCol: "B" },
  { role: "C1", label: "6. C1_VTVL", match: /C1_VTVL/i, keyCol: "A", keyType: "VTVL" },
  { role: "C2", label: "7. C2_NangLuc", match: /C2_NangLuc/i, keyCol: "A", keyType: "VTVL" },
  { role: "C3", label: "8. C3_NhuCau", match: /C3_NhuCau/i, keyCol: "A", keyType: "VTVL" },
  { role: "D1", label: "9. D1_NhanLuc", match: /D1_NhanLuc/i, keyCol: "B", keyType: "VTVL" },
  { role: "D2", label: "10. D2_PhuongAn", match: /D2_PhuongAn/i, keyCol: "A", keyType: "VTVL" },
  { role: "DUBAO", label: "DU_BAO_5_NAM", match: /DU_BAO_5_NAM/i, keyCol: "A", keyType: "VTVL" },
];

export function specForSheet(sheetName: string): SheetSpec | undefined {
  return SHEET_SPECS.find((s) => s.match.test(sheetName));
}

/** Tên file kết quả: thay cụm "HĐND-UBND" bằng tên khối; không có cụm đó thì thêm hậu tố. */
export function outputFileName(inputName: string, org: Org): string {
  const nfc = inputName.normalize("NFC").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_");
  const dot = nfc.toLowerCase().endsWith(".xlsx") ? nfc.length - 5 : nfc.length;
  const stem = nfc.slice(0, dot);
  const label = ORG_LABEL[org];
  const H = "H[ĐD]ND";
  const SEP = "\\s*(?:[-–—_&+,/]|và|va)\\s*";
  const pair = new RegExp(`${H}${SEP}UBND|UBND${SEP}${H}`, "i");
  const replaced = pair.test(stem) ? stem.replace(pair, label) : `${stem} - ${label}`;
  return `${replaced}.xlsx`;
}
