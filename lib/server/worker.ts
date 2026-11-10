// Xử lý một tác vụ tách file. Dùng chung cho hàng đợi Vercel Queues và chế độ chạy nền (inline).
import { Org, outputFileName } from "../xlsx/classify";
import { SplitError, splitWorkbook } from "../xlsx/split";
import { LIMITS, XLSX_MIME } from "./config";
import { AssembleError, loadJob, outputFile, readInput, saveJob } from "./jobs";
import { putObject } from "./storage";

/** Ném ra khi cần hàng đợi giao lại thông điệp sau (không phải lỗi dữ liệu). */
export class RetryLater extends Error {}

// Mỗi instance chỉ tách 1 file tại một thời điểm: một file 37 MB cần ~1 GB RAM khi xử lý,
// chạy song song trong cùng instance dễ tràn bộ nhớ. Các tác vụ khác xếp hàng trong tiến trình.
let chain: Promise<unknown> = Promise.resolve();

export function processJob(id: string): Promise<void> {
  const run = chain.then(() => doProcess(id));
  chain = run.catch(() => {});
  return run;
}

async function doProcess(id: string): Promise<void> {
  const job = await loadJob(id);
  if (!job) return; // đã bị xóa/hết hạn -> bỏ qua
  if (job.status === "done") return; // giao trùng -> bỏ qua (idempotent)
  if (job.status === "processing" && Date.now() - job.updatedAt < 150_000) {
    throw new RetryLater("Tác vụ đang được xử lý ở tiến trình khác");
  }
  if (job.status !== "queued" && job.status !== "processing") return;

  job.attempts += 1;
  if (job.attempts > LIMITS.maxAttempts) {
    job.status = "error";
    job.error = { code: "TOO_MANY_ATTEMPTS", message: "Xử lý không thành công sau nhiều lần thử. Hãy tải lại file và thử lại." };
    await saveJob(job);
    return;
  }
  job.status = "processing";
  job.error = undefined;
  await saveJob(job);

  const started = Date.now();
  try {
    let input: Buffer | null;
    try {
      // Đối chiếu lại số phần, kích thước và SHA-256 trước khi tách
      input = await readInput(id, { parts: job.parts, size: job.size, sha256: job.sha256, maxBytes: LIMITS.maxFileBytes });
    } catch (e) {
      if (e instanceof AssembleError) throw new SplitError("INVALID_FILE", e.message);
      throw e;
    }
    if (!input) throw new SplitError("INVALID_FILE", "Không tìm thấy file đã tải lên (có thể đã quá 24 giờ và bị xóa). Hãy tải lại file.");
    const result = await splitWorkbook(input, job.mapping ?? {});
    const outputs = {} as Record<Org, { name: string; size: number }>;
    const orgs = ["HDND", "UBND"] as Org[];
    await Promise.all(orgs.map((org) => putObject(outputFile(id, org), result.files[org], XLSX_MIME)));
    for (const org of orgs) outputs[org] = { name: outputFileName(job.fileName, org), size: result.files[org].length };
    const latest = (await loadJob(id)) ?? job;
    latest.status = "done";
    latest.attempts = job.attempts;
    latest.outputs = outputs;
    latest.report = result.report;
    latest.durationMs = Date.now() - started;
    latest.error = undefined;
    await saveJob(latest);
  } catch (e) {
    if (e instanceof SplitError) {
      job.status = "error";
      job.error = { code: e.code, message: e.message };
      await saveJob(job);
      return; // lỗi do dữ liệu: không thử lại
    }
    console.error(`[worker] job ${id} lần ${job.attempts} lỗi:`, e);
    const final = job.attempts >= LIMITS.maxAttempts;
    job.status = final ? "error" : "queued";
    job.error = final ? { code: "INTERNAL", message: "Hệ thống gặp lỗi khi xử lý file này. Hãy thử lại sau ít phút." } : undefined;
    await saveJob(job);
    if (!final) throw e;
  }
}
