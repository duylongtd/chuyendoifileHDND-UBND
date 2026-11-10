// Trạng thái tác vụ lưu dưới dạng JSON ngay trong kho file (jobs/<id>/job.json) – không cần cơ sở dữ liệu.
import crypto from "node:crypto";
import type { Org } from "../xlsx/classify";
import type { PositionInfo, SplitReport } from "../xlsx/split";
import { deleteObjects, deletePrefix, getObject, listObjects, putObject } from "./storage";

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

/**
 * Ghép các phần đã tải lên thành jobs/<id>/input.xlsx (một lần), rồi xóa các phần.
 * Trả về nội dung file; nếu đã ghép từ trước thì đọc lại file đã ghép.
 */
export async function assembleInput(id: string, expectedParts: number | undefined, maxBytes: number): Promise<Buffer | null> {
  const existing = await getObject(inputFile(id));
  if (existing) return existing;
  const parts = (await listObjects(partsPrefix(id))).sort((a, b) => a.path.localeCompare(b.path));
  if (!parts.length) return null;
  if (expectedParts && parts.length !== expectedParts) {
    throw new AssembleError(`File tải lên chưa đủ (nhận ${parts.length}/${expectedParts} phần). Hãy tải lại.`);
  }
  for (let i = 0; i < parts.length; i++) {
    if (parts[i].path !== partFile(id, i)) throw new AssembleError("File tải lên bị thiếu phần ở giữa. Hãy tải lại.");
  }
  if (parts.reduce((n, p) => n + p.size, 0) > maxBytes) throw new AssembleError(`File vượt ${Math.floor(maxBytes / 1048576)} MB.`);
  const bufs: Buffer[] = [];
  for (const p of parts) {
    const b = await getObject(p.path);
    if (!b) throw new AssembleError("Không đọc được một phần của file vừa tải lên. Hãy tải lại.");
    bufs.push(b);
  }
  const whole = Buffer.concat(bufs);
  await putObject(inputFile(id), whole, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  await deleteObjects(parts.map((p) => p.path)).catch(() => {});
  return whole;
}
