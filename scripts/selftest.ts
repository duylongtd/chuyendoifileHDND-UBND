// Tự kiểm các quy tắc cốt lõi:  npm test
import assert from "node:assert/strict";
import { orgFromAgency, outputFileName, suggestOrg } from "../lib/xlsx/classify";
import { translateFormula } from "../lib/xlsx/formula";
import { escapeFormula, unescapeFormula } from "../lib/xlsx/xml";
import { ApiError, uploadFile, type SendHooks, type UploadOptions } from "../lib/client/upload";

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
eq(org("Bí thư Đảng ủy, Chủ tịch HĐND xã"), "HDND", "kiêm nhiệm: vẫn là chức danh Chủ tịch HĐND");
eq(org("Phó Bí thư Thường trực Đảng ủy, Chủ tịch Hội đồng nhân dân phường"), "HDND");
eq(org("Chủ tịch HĐND, Chủ tịch UBND xã"), "UBND?", "ghi cả hai chức danh: cần người dùng xác nhận");
eq(org("Chuyên viên giúp việc Chủ tịch HĐND"), "UBND", "vị trí giúp việc không phải chức danh HĐND");
eq(suggestOrg("Chánh Văn phòng HĐND và UBND phường").related, true, "có nhắc HĐND: luôn hiện để soát");
eq(suggestOrg("Chuyên viên về lĩnh vực y tế").related, undefined);
eq(org(" ", "Giúp Trưởng ban tổ chức thực hiện nhiệm vụ"), "HDND?", "tên trống: gợi ý từ C2, cần xác nhận");
eq(org("", "Tham mưu công tác văn phòng"), "UBND?", "tên trống, không đủ căn cứ: cần xác nhận");
eq(org("", ""), "UBND?");

// --- Cột "Cơ quan" của nhiệm vụ chưa gắn công việc ---
eq(orgFromAgency("HĐND phường Sông Trí"), "HDND");
eq(orgFromAgency("Hội đồng nhân dân xã A"), "HDND");
eq(orgFromAgency("UBND phường Sông Trí"), "UBND");
eq(orgFromAgency("Ủy ban nhân dân phường Hà Huy Tập"), "UBND");
eq(orgFromAgency("HĐND và UBND phường"), "UBND", "ghi cả hai: xếp UBND");
eq(orgFromAgency(""), "UBND", "để trống: xếp UBND");

// --- Tên file kết quả ---
eq(outputFileName("01.HĐND-UBND Phường Hà Huy Tập.xlsx", "HDND"), "01.HĐND Phường Hà Huy Tập.xlsx");
eq(outputFileName("01.HĐND-UBND Phường Hà Huy Tập.xlsx", "UBND"), "01.UBND Phường Hà Huy Tập.xlsx");
eq(outputFileName("01.HĐND-UBND Phường Hà Huy Tập.xlsx".normalize("NFD"), "HDND"), "01.HĐND Phường Hà Huy Tập.xlsx", "tên file kiểu macOS (NFD)");
eq(outputFileName("HDND_UBND xa A.XLSX", "UBND"), "UBND xa A.xlsx");
eq(outputFileName("UBND và HĐND xã B.xlsx", "HDND"), "HĐND xã B.xlsx");
eq(outputFileName("Bo bieu xa C.xlsx", "HDND"), "Bo bieu xa C - HĐND.xlsx");
eq(outputFileName('a/b:c*"?.xlsx', "UBND"), "a_b_c___ - UBND.xlsx", "ký tự không hợp lệ trong tên file");

// --- Tải lên: tự chuyển sang phần nhỏ hơn khi mạng chặn/rớt gói lớn, mạng chậm, mạng đứng ---
async function uploadTests() {
  const file = new File([new Uint8Array(9 * 1024 * 1024 + 123)], "a.xlsx");
  const MB = 1024 * 1024;
  const PART = 4 * MB;
  type Call = { part: number; total: number; size: number; ok: boolean };
  const run = async (send: (c: Call, hooks: SendHooks) => Promise<void>, extra: UploadOptions = {}) => {
    const calls: Call[] = [];
    const notes: string[] = [];
    let result: { parts: number } | null = null;
    let error: Error | null = null;
    try {
      result = await uploadFile("job", file, PART, () => {}, (x) => x && notes.push(x), {
        retryDelayMs: 1,
        ...extra,
        send: async (url, chunk, hooks) => {
          const q = new URL(url, "http://x").searchParams;
          const c: Call = { part: Number(q.get("part")), total: Number(q.get("total")), size: chunk.size, ok: false };
          calls.push(c);
          await send(c, hooks);
          c.ok = true;
        },
      });
    } catch (e) {
      error = e as Error;
    }
    return { calls, notes, result, error };
  };
  const failAbove = (maxOk: number, status: number) => async (c: Call) => {
    if (c.size > maxOk) throw new ApiError("lỗi", status);
  };
  // mạng tốt: 3 phần 4 MB, không ghi chú
  let r = await run(async () => {});
  eq(r.result?.parts, 3, "mạng tốt: 3 phần 4 MB");
  eq(r.calls.length, 3);
  eq(r.notes.length, 0);
  // mạng rớt mọi gói > 1,5 MB: phải tự hạ xuống 1 MB và thành công, gửi đủ mọi byte
  r = await run(failAbove(1.5 * MB, 0));
  eq(r.result?.parts, 10, "mạng chặn gói lớn: tự chuyển sang 10 phần 1 MB");
  const okParts = r.calls.filter((c) => c.ok && c.total === 10);
  eq(okParts.map((c) => c.part).sort((a, b) => a - b), [0, 1, 2, 3, 4, 5, 6, 7, 8, 9], "đủ 10 phần, không thiếu phần nào");
  eq(okParts.reduce((s, c) => s + c.size, 0), file.size, "tổng byte gửi đi bằng đúng kích thước file");
  eq(r.notes.length > 0, true, "có thông báo đang tải lại với phần nhỏ hơn");
  // máy chủ trả 503 với gói > 2,5 MB: hạ xuống 2 MB là đủ
  r = await run(failAbove(2.5 * MB, 503));
  eq(r.result?.parts, 5, "lỗi 5xx với gói lớn: chuyển sang 5 phần 2 MB");
  // mạng chết hẳn: báo lỗi rõ, không treo
  r = await run(failAbove(0, 0));
  eq(r.result, null);
  eq(/Tải lại file này/.test(r.error?.message ?? ""), true, "mạng chết: thông báo hướng dẫn bấm Tải lại");
  // lỗi do dữ liệu (4xx): dừng ngay, không thử lại vô ích
  r = await run(failAbove(0, 409));
  eq(r.calls.length <= 4, true, "lỗi 4xx: không thử lại (" + r.calls.length + " lần gọi)");

  // Bộ gửi giả lập đường truyền có tốc độ: đẩy byte dần, bị hủy thì báo lỗi mạng
  const paced = (bytesPerMs: (c: Call) => number) => (c: Call, h: SendHooks) =>
    new Promise<void>((resolve, reject) => {
      let loaded = 0;
      const t = setInterval(() => {
        loaded = Math.min(c.size, loaded + bytesPerMs(c) * 5);
        if (loaded > 0) h.onBytes(loaded);
        if (loaded >= c.size) {
          clearInterval(t);
          resolve();
        }
      }, 5);
      h.signal.addEventListener("abort", () => {
        clearInterval(t);
        reject(new ApiError("hủy", 0));
      });
    });
  const fast = { tickMs: 5, slowCheckMs: 40, stallMs: 80, responseWaitMs: 80 };
  // mạng đứng với gói > 1,5 MB (không đẩy được byte nào): không treo, tự xuống 1 MB
  let t0 = Date.now();
  r = await run(paced((c) => (c.size > 1.5 * MB ? 0 : MB)), fast);
  eq(r.result?.parts, 10, "đường truyền đứng với gói lớn: xuống 10 phần 1 MB");
  eq(Date.now() - t0 < 3000, true, "phát hiện đứng nhanh, không chờ hết thời gian của máy chủ");
  eq(r.notes.some((x) => /bị đứng/.test(x)), true, "có thông báo đường truyền bị đứng");
  eq(r.calls.filter((c) => c.total === 3).length <= 4, true, "mức 4 MB chỉ thử một lượt rồi bỏ");
  // mạng chậm đều: 4 MB × 4 ước tính quá lâu -> nhảy thẳng xuống mức đủ nhanh, bỏ qua mức 2 MB
  r = await run(paced((c) => (c.total === 3 ? 400 : 50000)), { ...fast, slowPartSec: 3, targetPartSec: 1, stallMs: 5000, responseWaitMs: 5000 });
  eq(r.notes.some((x) => /chậm/.test(x)), true, "mạng chậm: có thông báo kèm tốc độ");
  eq(r.calls.some((c) => c.total === 5), false, "mạng chậm: bỏ qua mức 2 MB, xuống thẳng mức phù hợp");
  eq(r.result !== null && r.calls.filter((c) => c.ok).reduce((s, c) => s + c.size, 0) >= file.size, true, "mạng chậm: vẫn tải xong đủ byte");
  // mức cuối cùng mà một phần bị đứng: chỉ thử lại phần đó, không bỏ cuộc
  let stuckOnce = true;
  r = await run(
    (c, h) => {
      if (c.total !== 19) return Promise.reject(new ApiError("lỗi", 0));
      if (c.part === 7 && stuckOnce) {
        stuckOnce = false;
        return paced(() => 0)(c, h);
      }
      return Promise.resolve();
    },
    { ...fast, attempts: 2 },
  );
  eq(r.result?.parts, 19, "mức 512 KB: phần bị đứng được thử lại và tải xong");
  eq(r.calls.filter((c) => c.total === 19 && c.part === 7).length, 2, "chỉ gửi lại đúng phần bị đứng");
}

uploadTests()
  .then(() => console.log(`Tự kiểm: ${n}/${n} đạt`))
  .catch((e) => {
    console.error(e);
    process.exit(1);
  });
