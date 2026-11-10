// Trạng thái tác vụ lưu dưới dạng JSON ngay trong kho file (jobs/<id>/job.json) – không cần cơ sở dữ liệu.
import crypto from "node:crypto";
import type { Org } from "../xlsx/classify";
import type { PositionInfo, SplitReport } from "../xlsx/split";
import { deletePrefix, getObject, listObjects, putObject } from "./storage";

export type JobStatus = "created" | "ready" | "queued" | "processing" | "done" | "error";

export interface Job {
  id: string;
  fileName: string;
  size: number;
  status: JobStatus;
  createdAt: number;
  updatedAt: number;
  attempts: number;
  positions?: PositionInfo[];
  missingSheets?: string[];
  /** Thông tin để đối chiếu toàn vẹn file tải lên */
  parts?: number;
  sha256?: string;
  mapping?: Record<string, Org>;
  error?: { code: string; message: string };
  report?: SplitReport;
  outputs?: Record<Org, { name: string; size: number }>;
  durationMs?: number;
  queuedVia?: string;
}

const ID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;
export function isJobId(id: string): boolean {
  return ID_RE.test(id);
}

export const jobPrefix = (id: string) => `jobs/${id}/`;
export const jobFile = (id: string) => `jobs/${id}/job.json`;
export const inputFile = (id: string) => `jobs/${id}/input.xlsx`;
export const partsPrefix = (id: string) => `jobs/${id}/parts/`;
export const partFile = (id: string, n: number) => `jobs/${id}/parts/${String(n).padStart(4, "0")}`;
export const outputFile = (id: string, org: Org) => `jobs/${id}/${org.toLowerCase()}.xlsx`;

export async function createJob(fileName: string, size: number): Promise<Job> {
  const now = Date.now();
  const job: Job = { id: crypto.randomUUID(), fileName, size, status: "created", createdAt: now, updatedAt: now, attempts: 0 };
  await saveJob(job);
  return job;
}

export async function loadJob(id: string): Promise<Job | null> {
  if (!isJobId(id)) return null;
  const buf = await getObject(jobFile(id));
  if (!buf) return null;
  try {
    return JSON.parse(buf.toString("utf8")) as Job;
  } catch {
    return null;
  }
}

export async function saveJob(job: Job): Promise<void> {
  job.updatedAt = Date.now();
  await putObject(jobFile(job.id), JSON.stringify(job), "application/json");
}

export async function deleteJob(id: string): Promise<void> {
  if (isJobId(id)) await deletePrefix(jobPrefix(id));
}

export class AssembleError extends Error {}

async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T, index: number) => Promise<R>): Promise<R[]> {
  const out = new Array<R>(items.length);
  let next = 0;
  await Promise.all(
    Array.from({ length: Math.min(limit, items.length) }, async () => {
      while (next < items.length) {
        const i = next++;
        out[i] = await fn(items[i], i);
      }
    }),
  );
  return out;
}

export interface InputCheck {
  /** Số phần trình duyệt đã gửi */
  parts?: number;
  /** Kích thước file gốc (byte) */
  size?: number;
  /** SHA-256 (hex) của file gốc, do trình duyệt tính trước khi cắt */
  sha256?: string;
  maxBytes: number;
}

/**
 * Đọc file gốc từ các phần đã tải lên (đọc song song, ghép trong bộ nhớ; không ghi lại thành file mới),
 * rồi đối chiếu số phần, kích thước và SHA-256 với file trên máy người dùng để bảo đảm toàn vẹn.
 * Trả về null nếu chưa có phần nào.
 */
export async function readInput(id: string, check: InputCheck): Promise<Buffer | null> {
  let n = check.parts;
  if (!n) {
    // Không biết số phần (tác vụ cũ): dò trong kho
    const legacy = await getObject(inputFile(id));
    if (legacy) return legacy;
    n = (await listObjects(partsPrefix(id))).length;
    if (!n) return null;
  }
  const bufs = await mapLimit(Array.from({ length: n }, (_, i) => i), 6, (i) => getObject(partFile(id, i)));
  const have = bufs.filter(Boolean).length;
  if (!have) return null;
  if (have !== n) throw new AssembleError(`File tải lên chưa đủ (nhận ${have}/${n} phần). Hãy tải lại.`);
  const total = bufs.reduce((sum, b) => sum + b!.length, 0);
  if (total > check.maxBytes) throw new AssembleError(`File vượt ${Math.floor(check.maxBytes / 1048576)} MB.`);
  const whole = Buffer.concat(bufs as Buffer[], total);
  if (check.size && whole.length !== check.size) {
    throw new AssembleError("File nhận được không đủ dung lượng so với file gốc (lỗi đường truyền). Hãy tải lại.");
  }
  if (check.sha256 && crypto.createHash("sha256").update(whole).digest("hex") !== check.sha256.toLowerCase()) {
    throw new AssembleError("File nhận được không khớp với file gốc (lỗi đường truyền). Hãy tải lại.");
  }
  return whole;
}
