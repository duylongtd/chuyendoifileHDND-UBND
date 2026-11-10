// Trạng thái tác vụ lưu dưới dạng JSON ngay trong kho file (jobs/<id>/job.json) – không cần cơ sở dữ liệu.
import crypto from "node:crypto";
import type { Org } from "../xlsx/classify";
import type { PositionInfo, SplitReport } from "../xlsx/split";
import { deletePrefix, getObject, putObject } from "./storage";

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
