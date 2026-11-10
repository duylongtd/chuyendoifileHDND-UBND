import { LIMITS } from "@/lib/server/config";
import { clientIp, fail, json, rateLimit, route } from "@/lib/server/http";
import { createJob, isJobId, loadJob } from "@/lib/server/jobs";

function validateFile(f: unknown): { name: string; size: number } | string {
  const o = f as { name?: unknown; size?: unknown };
  const name = typeof o?.name === "string" ? o.name.normalize("NFC").trim() : "";
  const size = typeof o?.size === "number" ? o.size : NaN;
  if (!name || name.length > 200) return "Tên file không hợp lệ.";
  if (!/\.xlsx$/i.test(name)) return `"${name}": chỉ nhận file Excel .xlsx (file .xls cần lưu lại thành .xlsx).`;
  if (name.startsWith("~$")) return `"${name}" là file tạm của Excel, hãy chọn file gốc.`;
  if (!Number.isFinite(size) || size <= 0) return `"${name}": file rỗng.`;
  if (size > LIMITS.maxFileBytes) return `"${name}": vượt ${Math.floor(LIMITS.maxFileBytes / 1048576)} MB.`;
  return { name, size };
}

// Tạo tác vụ cho một lượt tải lên (tối đa 3 file)
export const POST = route(async (req: Request) => {
  const body = (await req.json().catch(() => null)) as { files?: unknown } | null;
  const files = Array.isArray(body?.files) ? body!.files : [];
  if (!files.length) return fail(400, "NO_FILES", "Chưa chọn file nào.");
  if (files.length > LIMITS.maxFilesPerBatch) return fail(400, "TOO_MANY_FILES", `Mỗi lượt chỉ tải tối đa ${LIMITS.maxFilesPerBatch} file.`);
  const valid: { name: string; size: number }[] = [];
  for (const f of files) {
    const v = validateFile(f);
    if (typeof v === "string") return fail(400, "INVALID_FILE", v);
    valid.push(v);
  }
  // Giới hạn theo IP: một cơ quan thường dùng chung một IP ra Internet nên để rộng (90 file / 10 phút)
  if (!rateLimit(`create:${clientIp(req)}`, LIMITS.maxFilesPer10MinPerIp, 10 * 60 * 1000, valid.length)) {
    return fail(429, "TOO_MANY", "Bạn đang tải lên quá nhiều file. Hãy chờ vài phút rồi thử lại.");
  }
  const jobs = [];
  for (const v of valid) jobs.push(await createJob(v.name, v.size));
  return json({ jobs });
});

// Trạng thái nhiều tác vụ (trang web hỏi định kỳ)
export const GET = route(async (req: Request) => {
  const ids = (new URL(req.url).searchParams.get("ids") ?? "").split(",").filter(isJobId).slice(0, 12);
  const jobs = await Promise.all(ids.map((id) => loadJob(id)));
  return json({ jobs: jobs.filter(Boolean), missing: ids.filter((_, i) => !jobs[i]) });
});
