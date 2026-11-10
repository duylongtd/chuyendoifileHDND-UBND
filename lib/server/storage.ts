// Lớp lưu trữ file: Vercel Blob (private) khi deploy, thư mục .data/ khi chạy trên máy.
import fs from "node:fs";
import fsp from "node:fs/promises";
import path from "node:path";
import { Readable } from "node:stream";
import { blobAccess, storageDriver } from "./config";

export interface StoredObject {
  path: string;
  size: number;
  uploadedAt: number;
}

const LOCAL_ROOT = path.join(process.cwd(), ".data");

function localPath(p: string): string {
  const full = path.resolve(LOCAL_ROOT, p);
  if (!full.startsWith(LOCAL_ROOT + path.sep)) throw new Error("Đường dẫn không hợp lệ");
  return full;
}

export async function putObject(p: string, body: Buffer | string, contentType: string): Promise<void> {
  if (storageDriver() === "blob") {
    const { put } = await import("@vercel/blob");
    // File lớn (kết quả ~30 MB) tải theo nhiều phần cho ổn định
    const multipart = typeof body !== "string" && body.length > 8 * 1024 * 1024;
    await put(p, body, { access: blobAccess(), contentType, addRandomSuffix: false, allowOverwrite: true, multipart });
    return;
  }
  const full = localPath(p);
  await fsp.mkdir(path.dirname(full), { recursive: true });
  const tmp = `${full}.${process.pid}.${Date.now()}.tmp`;
  await fsp.writeFile(tmp, body);
  await fsp.rename(tmp, full);
}

export async function getObject(p: string): Promise<Buffer | null> {
  if (storageDriver() === "blob") {
    const { get } = await import("@vercel/blob");
    const r = await get(p, { access: blobAccess(), useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return Buffer.from(await new Response(r.stream).arrayBuffer());
  }
  try {
    return await fsp.readFile(localPath(p));
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === "ENOENT") return null;
    throw e;
  }
}

export async function getObjectStream(p: string): Promise<{ stream: ReadableStream<Uint8Array>; size?: number } | null> {
  if (storageDriver() === "blob") {
    const { get } = await import("@vercel/blob");
    const r = await get(p, { access: blobAccess(), useCache: false });
    if (!r || r.statusCode !== 200 || !r.stream) return null;
    return { stream: r.stream as ReadableStream<Uint8Array>, size: r.blob?.size };
  }
  try {
    const full = localPath(p);
    const st = await fsp.stat(full);
    return { stream: Readable.toWeb(fs.createReadStream(full)) as ReadableStream<Uint8Array>, size: st.size };
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === "ENOENT") return null;
    throw e;
  }
}

export async function listObjects(prefix: string): Promise<StoredObject[]> {
  if (storageDriver() === "blob") {
    const { list } = await import("@vercel/blob");
    const out: StoredObject[] = [];
    let cursor: string | undefined;
    do {
      const r = await list({ prefix, cursor, limit: 1000 });
      for (const b of r.blobs) out.push({ path: b.pathname, size: b.size, uploadedAt: new Date(b.uploadedAt).getTime() });
      cursor = r.hasMore ? r.cursor : undefined;
    } while (cursor && out.length < 20000);
    return out;
  }
  const out: StoredObject[] = [];
  const root = localPath(prefix.replace(/\/$/, "") || ".keep");
  const walk = async (dir: string) => {
    let items: fs.Dirent[];
    try {
      items = await fsp.readdir(dir, { withFileTypes: true });
    } catch {
      return;
    }
    for (const it of items) {
      const full = path.join(dir, it.name);
      if (it.isDirectory()) await walk(full);
      else {
        const st = await fsp.stat(full);
        out.push({ path: path.relative(LOCAL_ROOT, full).split(path.sep).join("/"), size: st.size, uploadedAt: st.mtimeMs });
      }
    }
  };
  await walk(root);
  return out;
}

export async function deleteObjects(paths: string[]): Promise<void> {
  if (!paths.length) return;
  if (storageDriver() === "blob") {
    const { del } = await import("@vercel/blob");
    for (let i = 0; i < paths.length; i += 100) await del(paths.slice(i, i + 100));
    return;
  }
  for (const p of paths) await fsp.rm(localPath(p), { force: true });
}

export async function deletePrefix(prefix: string): Promise<number> {
  const objs = await listObjects(prefix);
  await deleteObjects(objs.map((o) => o.path));
  if (storageDriver() === "local") await fsp.rm(localPath(prefix), { recursive: true, force: true }).catch(() => {});
  return objs.length;
}
