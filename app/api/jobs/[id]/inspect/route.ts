// Đọc danh sách vị trí ở biểu C1 và gợi ý HĐND/UBND để người dùng xác nhận trước khi tách.
import { LIMITS } from "@/lib/server/config";
import { fail, json, route } from "@/lib/server/http";
import { assembleInput, AssembleError, loadJob, saveJob } from "@/lib/server/jobs";
import { inspectWorkbook, SplitError } from "@/lib/xlsx/split";

export const maxDuration = 120;
type Ctx = { params: Promise<{ id: string }> };

export const POST = route(async (req: Request, ctx: Ctx) => {
  const job = await loadJob((await ctx.params).id);
  if (!job) return fail(404, "NOT_FOUND", "Không tìm thấy tác vụ (có thể đã hết hạn 24 giờ).");
  if (job.status !== "created" && job.status !== "ready" && job.status !== "error") return json({ job });

  const reject = async (code: string, message: string, status = 422) => {
    job.status = "error";
    job.error = { code, message };
    await saveJob(job);
    return json({ job, error: job.error }, status);
  };

  const body = (await req.json().catch(() => null)) as { parts?: unknown } | null;
  const expected = Number.isInteger(body?.parts) ? (body!.parts as number) : undefined;
  let input: Buffer | null;
  try {
    input = await assembleInput(job.id, expected, LIMITS.maxFileBytes);
  } catch (e) {
    if (e instanceof AssembleError) return reject("BAD_UPLOAD", e.message);
    throw e;
  }
  if (!input) return reject("NO_INPUT", "Chưa nhận được file. Hãy tải lại.");
  if (input.length > LIMITS.maxFileBytes) return reject("TOO_LARGE", `File vượt ${Math.floor(LIMITS.maxFileBytes / 1048576)} MB.`);
  if (input.length < 4 || input.readUInt32LE(0) !== 0x04034b50) {
    return reject("INVALID_FILE", "Đây không phải file Excel .xlsx hợp lệ (có thể là file .xls đổi đuôi, hoặc file đang đặt mật khẩu).");
  }
  try {
    const res = await inspectWorkbook(input);
    job.size = input.length;
    job.positions = res.positions;
    job.missingSheets = res.missingSheets;
    job.status = "ready";
    job.error = undefined;
    await saveJob(job);
    return json({ job });
  } catch (e) {
    if (e instanceof SplitError) return reject(e.code, e.message);
    throw e;
  }
});
