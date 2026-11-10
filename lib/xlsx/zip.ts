// Đọc/ghi file ZIP tối giản cho .xlsx.
// - Đọc từ Buffer qua central directory (không phụ thuộc data descriptor).
// - Ghi: entry không đổi được chép nguyên dữ liệu ĐÃ NÉN (không giải nén/nén lại) -> nhanh, ít RAM.
import zlib from "node:zlib";


export class ZipError extends Error {}

export interface ZipEntry {
  name: string;
  nameBytes: Buffer;
  flags: number;
  method: number; // 0 = store, 8 = deflate
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
  modTime: number;
  modDate: number;
  externalAttrs: number;
  /** Dữ liệu đã nén (cắt từ file gốc, không sao chép) */
  raw: Buffer;
}

const SIG_EOCD = 0x06054b50;
const SIG_CEN = 0x02014b50;
const SIG_LOC = 0x04034b50;

/** Giới hạn tổng dung lượng sau giải nén để chống "bom nén". */
export const MAX_TOTAL_UNCOMPRESSED = 1.2 * 1024 * 1024 * 1024;
/** Một thành phần (một sheet) sau giải nén không quá mức này; bộ biểu thật lớn nhất ~80 MB */
export const MAX_ENTRY_UNCOMPRESSED = 300 * 1024 * 1024;
export const MAX_ENTRIES = 2000;

export function readZip(buf: Buffer): ZipEntry[] {
  if (buf.length < 22) throw new ZipError("File quá nhỏ, không phải file .xlsx hợp lệ.");
  // Tìm End Of Central Directory từ cuối file (comment tối đa 65535 byte)
  let eocd = -1;
  const min = Math.max(0, buf.length - 22 - 65535);
  for (let i = buf.length - 22; i >= min; i--) {
    if (buf.readUInt32LE(i) === SIG_EOCD) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new ZipError("Không đọc được cấu trúc file (không phải .xlsx hoặc file bị hỏng).");
  const count = buf.readUInt16LE(eocd + 10);
  const cenSize = buf.readUInt32LE(eocd + 12);
  const cenOffset = buf.readUInt32LE(eocd + 16);
  if (count === 0xffff || cenOffset === 0xffffffff) throw new ZipError("File dùng định dạng ZIP64, chưa được hỗ trợ.");
  if (count > MAX_ENTRIES) throw new ZipError("File chứa quá nhiều thành phần.");
  if (cenOffset + cenSize > buf.length) throw new ZipError("File bị cắt cụt hoặc hỏng.");

  const entries: ZipEntry[] = [];
  let p = cenOffset;
  let total = 0;
  for (let i = 0; i < count; i++) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== SIG_CEN) throw new ZipError("Mục lục file bị hỏng.");
    const flags = buf.readUInt16LE(p + 8);
    const method = buf.readUInt16LE(p + 10);
    const modTime = buf.readUInt16LE(p + 12);
    const modDate = buf.readUInt16LE(p + 14);
    const crc32 = buf.readUInt32LE(p + 16);
    const compressedSize = buf.readUInt32LE(p + 20);
    const uncompressedSize = buf.readUInt32LE(p + 24);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    const externalAttrs = buf.readUInt32LE(p + 38);
    const locOffset = buf.readUInt32LE(p + 42);
    const nameBytes = buf.subarray(p + 46, p + 46 + nameLen);
    const name = nameBytes.toString("utf8");
    p += 46 + nameLen + extraLen + commentLen;

    if (flags & 0x1) throw new ZipError("File được đặt mật khẩu/mã hóa, không xử lý được. Hãy bỏ mật khẩu rồi tải lại.");
    if (method !== 0 && method !== 8) throw new ZipError(`Kiểu nén không hỗ trợ (${method}).`);
    if (locOffset + 30 > buf.length || buf.readUInt32LE(locOffset) !== SIG_LOC) throw new ZipError("File bị hỏng (sai vị trí dữ liệu).");
    const locNameLen = buf.readUInt16LE(locOffset + 26);
    const locExtraLen = buf.readUInt16LE(locOffset + 28);
    const dataStart = locOffset + 30 + locNameLen + locExtraLen;
    if (dataStart + compressedSize > buf.length) throw new ZipError("File bị cắt cụt hoặc hỏng.");
    if (uncompressedSize > MAX_ENTRY_UNCOMPRESSED) throw new ZipError("File sau giải nén quá lớn, vượt giới hạn cho phép.");
    total += uncompressedSize;
    if (total > MAX_TOTAL_UNCOMPRESSED) throw new ZipError("File sau giải nén quá lớn, vượt giới hạn cho phép.");

    entries.push({
      name,
      nameBytes,
      flags,
      method,
      crc32,
      compressedSize,
      uncompressedSize,
      modTime,
      modDate,
      externalAttrs,
      raw: buf.subarray(dataStart, dataStart + compressedSize),
    });
  }
  return entries;
}

/** Giải nén thẳng vào vùng đệm cấp phát sẵn đúng kích thước (không tạo bản sao trung gian). */
export function inflateEntry(e: ZipEntry, into?: Buffer): Promise<Buffer> {
  if (e.method === 0) return Promise.resolve(e.raw);
  return new Promise((resolve, reject) => {
    // Dùng lại vùng đệm được truyền vào (nếu đủ lớn) để không cấp phát mới cho từng sheet
    const out = into && into.length >= e.uncompressedSize ? into.subarray(0, e.uncompressedSize) : Buffer.allocUnsafe(e.uncompressedSize);
    let off = 0;
    const z = zlib.createInflateRaw({ chunkSize: 1 << 20 });
    z.on("data", (d: Buffer) => {
      if (off + d.length > out.length) {
        z.destroy();
        reject(new ZipError(`Thành phần ${e.name} bị hỏng (sai kích thước).`));
        return;
      }
      d.copy(out, off);
      off += d.length;
    });
    z.on("error", () => reject(new ZipError(`Không giải nén được ${e.name}: file bị hỏng.`)));
    z.on("end", () => (off === out.length ? resolve(out) : reject(new ZipError(`Thành phần ${e.name} bị hỏng (sai kích thước).`))));
    z.end(e.raw);
  });
}

export interface OutEntry {
  name: string;
  nameBytes: Buffer;
  flags: number;
  method: number;
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
  modTime: number;
  modDate: number;
  externalAttrs: number;
  raw: Buffer;
}

/** Entry chép nguyên từ file gốc. */
export function passthrough(e: ZipEntry): OutEntry {
  return { ...e, flags: e.flags & ~0x8 };
}

type EntryMeta = Pick<ZipEntry, "name" | "nameBytes" | "flags" | "modTime" | "modDate" | "externalAttrs">;

/**
 * Entry có nội dung mới, nén theo luồng từ nhiều khúc dữ liệu (không cần ghép thành một vùng đệm lớn).
 * Ghi từng khúc 4 MB và chờ "drain" để không giữ event loop quá lâu.
 */
export function deflatedChunks(e: EntryMeta, chunks: Buffer[], level = 4): Promise<OutEntry> {
  return new Promise((resolve, reject) => {
    const z = zlib.createDeflateRaw({ level, chunkSize: 256 * 1024 });
    const parts: Buffer[] = [];
    let crc = 0;
    let size = 0;
    z.on("data", (d: Buffer) => parts.push(d));
    z.on("error", reject);
    z.on("end", () => {
      const raw = Buffer.concat(parts);
      resolve({
        name: e.name,
        nameBytes: e.nameBytes,
        flags: e.flags & ~0x8,
        method: 8,
        crc32: crc >>> 0,
        compressedSize: raw.length,
        uncompressedSize: size,
        modTime: e.modTime,
        modDate: e.modDate,
        externalAttrs: e.externalAttrs,
        raw,
      });
    });
    (async () => {
      const STEP = 4 * 1024 * 1024;
      for (const c of chunks) {
        crc = zlib.crc32(c, crc);
        size += c.length;
        for (let o = 0; o < c.length; o += STEP) {
          if (!z.write(c.subarray(o, o + STEP))) await new Promise((r) => z.once("drain", r));
        }
      }
      z.end();
    })().catch(reject);
  });
}

/** Entry có nội dung mới (nén lại). */
export function deflated(e: EntryMeta, data: Buffer, level = 4): Promise<OutEntry> {
  return deflatedChunks(e, [data], level);
}

export function writeZip(entries: OutEntry[]): Buffer {
  const parts: Buffer[] = [];
  const central: Buffer[] = [];
  let offset = 0;
  for (const e of entries) {
    const loc = Buffer.alloc(30);
    loc.writeUInt32LE(SIG_LOC, 0);
    loc.writeUInt16LE(20, 4);
    loc.writeUInt16LE(e.flags, 6);
    loc.writeUInt16LE(e.method, 8);
    loc.writeUInt16LE(e.modTime, 10);
    loc.writeUInt16LE(e.modDate, 12);
    loc.writeUInt32LE(e.crc32, 14);
    loc.writeUInt32LE(e.compressedSize, 18);
    loc.writeUInt32LE(e.uncompressedSize, 22);
    loc.writeUInt16LE(e.nameBytes.length, 26);
    loc.writeUInt16LE(0, 28);
    parts.push(loc, e.nameBytes, e.raw);

    const cen = Buffer.alloc(46);
    cen.writeUInt32LE(SIG_CEN, 0);
    cen.writeUInt16LE(20, 4);
    cen.writeUInt16LE(20, 6);
    cen.writeUInt16LE(e.flags, 8);
    cen.writeUInt16LE(e.method, 10);
    cen.writeUInt16LE(e.modTime, 12);
    cen.writeUInt16LE(e.modDate, 14);
    cen.writeUInt32LE(e.crc32, 16);
    cen.writeUInt32LE(e.compressedSize, 20);
    cen.writeUInt32LE(e.uncompressedSize, 24);
    cen.writeUInt16LE(e.nameBytes.length, 28);
    cen.writeUInt32LE(e.externalAttrs, 38);
    cen.writeUInt32LE(offset, 42);
    central.push(cen, e.nameBytes);
    offset += 30 + e.nameBytes.length + e.raw.length;
  }
  const cenSize = central.reduce((n, b) => n + b.length, 0);
  if (offset + cenSize > 0xfffffff0) throw new ZipError("File kết quả quá lớn.");
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(SIG_EOCD, 0);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(cenSize, 12);
  eocd.writeUInt32LE(offset, 16);
  return Buffer.concat([...parts, ...central, eocd]);
}
