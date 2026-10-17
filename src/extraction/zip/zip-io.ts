import { getNodeRequire } from "../xml/runtime-env.js";
type ExportEntry = {
    name: string;
    data: Uint8Array;
};

type ZipEntryTimestamp = {
    dosTime: number;
    dosDate: number;
};

type ZipCentralEntry = {
    path: string;
    method: number;
    compressedSize: number;
    uncompressedSize: number;
    localHeaderOffset: number;
};

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
const LOCAL_FILE_SIGNATURE = 0x04034b50;
const textDecoder = new TextDecoder();
const textEncoder = new TextEncoder();
const crcTable = buildCrc32Table();
const fixedZipEntryTimestamp = toDosDateTime(2025, 1, 1, 0, 0, 0);
const utf8FileNameFlag = 0x0800;

function buildCrc32Table(): Uint32Array {
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i += 1) {
        let value = i;
        for (let bit = 0; bit < 8; bit += 1) {
            value = (value & 1) === 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
        }
        table[i] = value >>> 0;
    }
    return table;
}

function crc32(bytes: Uint8Array): number {
    let crc = 0xffffffff;
    for (const byte of bytes) {
        crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
}

function hasNonAsciiCharacters(value: string): boolean {
    return /[^\x00-\x7f]/.test(value);
}

function toDosDateTime(
    year: number,
    month: number,
    day: number,
    hour: number,
    minute: number,
    second: number
): ZipEntryTimestamp {
    const clampedYear = Math.max(1980, Math.min(2107, year));
    const dosTime = ((hour & 0x1f) << 11) | ((minute & 0x3f) << 5) | (Math.floor(second / 2) & 0x1f);
    const dosDate = (((clampedYear - 1980) & 0x7f) << 9) | ((month & 0x0f) << 5) | (day & 0x1f);
    return {
        dosTime,
        dosDate,
    };
}

async function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
    // Use synchronous zlib on Node. It skips stream allocation per entry.
    const nodeRequire = getNodeRequire();
    if (nodeRequire) {
        const zlib = nodeRequire("node:zlib") as { inflateRawSync: (input: Uint8Array) => Uint8Array };
        // Copy into a plain, unpooled Uint8Array. Bulk copy beats per-element copy.
        return new Uint8Array(zlib.inflateRawSync(data));
    }
    if (typeof DecompressionStream === "function") {
        try {
            const stream = new Blob([data as unknown as BlobPart])
                .stream()
                .pipeThrough(new DecompressionStream("deflate-raw"));
            const buffer = await new Response(stream).arrayBuffer();
            return new Uint8Array(buffer);
        } catch (_error) {
            // Fall through when the runtime lacks deflate-raw support. The throw below reports it.
        }
    }
    throw new Error("This environment does not support ZIP deflate decompression.");
}

function readUint16(data: Uint8Array, offset: number): number {
    return data[offset] | (data[offset + 1] << 8);
}

function readUint32(data: Uint8Array, offset: number): number {
    return (data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24)) >>> 0;
}

/** Normalizes a ZIP entry name to an OPC part path. Rejects paths that escape the package root. */
function normalizePartPath(entryName: string): string {
    const raw = (entryName.split("#", 1)[0] || "").replace(/\\/g, "/").replace(/^\/+/, "");
    const parts: string[] = [];
    for (const part of raw.split("/")) {
        if (part === "" || part === ".") continue;
        if (part === "..") {
            if (parts.length === 0) {
                throw new Error(`OPC part path escapes package root: ${entryName}`);
            }
            parts.pop();
            continue;
        }
        parts.push(part);
    }
    if (parts.length === 0) {
        throw new Error(`OPC part path is empty: ${entryName}`);
    }
    return parts.join("/");
}

/** Finds the end-of-central-directory record, scanning back over a trailing comment of up to 64 KiB. */
function findEndOfCentralDirectory(data: Uint8Array): number {
    const minOffset = Math.max(0, data.length - 0xffff - 22);
    for (let offset = data.length - 22; offset >= minOffset; offset -= 1) {
        if (readUint32(data, offset) === EOCD_SIGNATURE) {
            return offset;
        }
    }
    return -1;
}

/** Reads the central directory. ZIP64 and encrypted entries are out of scope for .xlsx input. */
function readCentralDirectory(data: Uint8Array): ZipCentralEntry[] {
    const eocdOffset = findEndOfCentralDirectory(data);
    if (eocdOffset < 0) {
        throw new Error("ZIP end-of-central-directory record was not found.");
    }
    const entryCount = readUint16(data, eocdOffset + 10);
    const entries: ZipCentralEntry[] = [];
    let offset = readUint32(data, eocdOffset + 16);
    for (let index = 0; index < entryCount; index += 1) {
        if (readUint32(data, offset) !== CENTRAL_DIRECTORY_SIGNATURE) {
            throw new Error("ZIP central directory entry signature is invalid.");
        }
        const method = readUint16(data, offset + 10);
        const nameLength = readUint16(data, offset + 28);
        const extraLength = readUint16(data, offset + 30);
        const commentLength = readUint16(data, offset + 32);
        const nameStart = offset + 46;
        const name = textDecoder.decode(data.subarray(nameStart, nameStart + nameLength));
        if (method !== 0 && method !== 8) {
            throw new Error(`${name}: Unsupported ZIP compression method: ${method}`);
        }
        entries.push({
            path: normalizePartPath(name),
            method,
            compressedSize: readUint32(data, offset + 20),
            uncompressedSize: readUint32(data, offset + 24),
            localHeaderOffset: readUint32(data, offset + 42),
        });
        offset = nameStart + nameLength + extraLength + commentLength;
    }
    return entries;
}

/** Reads one entry body. Sizes come from the central directory, so data-descriptor entries work too. */
async function readEntryData(data: Uint8Array, entry: ZipCentralEntry): Promise<Uint8Array> {
    if (readUint32(data, entry.localHeaderOffset) !== LOCAL_FILE_SIGNATURE) {
        throw new Error(`${entry.path}: ZIP local file header signature is invalid.`);
    }
    const nameLength = readUint16(data, entry.localHeaderOffset + 26);
    const extraLength = readUint16(data, entry.localHeaderOffset + 28);
    const dataStart = entry.localHeaderOffset + 30 + nameLength + extraLength;
    const compressed = data.slice(dataStart, dataStart + entry.compressedSize);
    if (entry.method === 0) {
        return compressed;
    }
    const inflated = await inflateRaw(compressed);
    if (inflated.length !== entry.uncompressedSize) {
        throw new Error(`${entry.path}: Deflated ZIP entry size mismatch.`);
    }
    return inflated;
}

// Rejects archives that no spreadsheet can produce. Sizes come from the
// central directory, so oversized packages fail before any inflation.
const MAX_ZIP_ENTRIES = 20000;
const MAX_ZIP_UNCOMPRESSED_BYTES = 512 * 1024 * 1024;

/** Unzips every entry into a map from OPC part path to raw bytes. Rejects packages over entry or size limits. */
async function unzipEntries(arrayBuffer: ArrayBuffer): Promise<Map<string, Uint8Array>> {
    const data = new Uint8Array(arrayBuffer);
    const entries = readCentralDirectory(data);
    if (entries.length > MAX_ZIP_ENTRIES) {
        throw new Error(`Too many ZIP entries: ${entries.length}. Limit is ${MAX_ZIP_ENTRIES}.`);
    }
    const declaredUncompressed = entries.reduce((sum, entry) => sum + entry.uncompressedSize, 0);
    if (declaredUncompressed > MAX_ZIP_UNCOMPRESSED_BYTES) {
        throw new Error("Uncompressed ZIP size is over 512 MiB.");
    }
    const files = new Map<string, Uint8Array>();
    let totalInflated = 0;
    for (const entry of entries) {
        const bytes = await readEntryData(data, entry);
        totalInflated += bytes.length;
        if (totalInflated > MAX_ZIP_UNCOMPRESSED_BYTES) {
            throw new Error("Uncompressed ZIP size is over 512 MiB.");
        }
        files.set(entry.path, bytes);
    }
    return files;
}

/**
 * Packs entries into an uncompressed ZIP archive.
 * Uses a fixed entry timestamp, so repeated exports stay byte-identical.
 */
function createStoredZip(entries: ExportEntry[]): Uint8Array {
    const localChunks: Uint8Array[] = [];
    const centralChunks: Uint8Array[] = [];
    let offset = 0;

    for (const entry of entries) {
        const nameBytes = textEncoder.encode(entry.name);
        const dataBytes = entry.data;
        const entryCrc32 = crc32(dataBytes);
        const generalPurposeBitFlag = hasNonAsciiCharacters(entry.name) ? utf8FileNameFlag : 0;

        const localHeader = new Uint8Array(30 + nameBytes.length);
        const localView = new DataView(localHeader.buffer);
        localView.setUint32(0, 0x04034b50, true);
        localView.setUint16(4, 20, true);
        localView.setUint16(6, generalPurposeBitFlag, true);
        localView.setUint16(8, 0, true);
        localView.setUint16(10, fixedZipEntryTimestamp.dosTime, true);
        localView.setUint16(12, fixedZipEntryTimestamp.dosDate, true);
        localView.setUint32(14, entryCrc32, true);
        localView.setUint32(18, dataBytes.length, true);
        localView.setUint32(22, dataBytes.length, true);
        localView.setUint16(26, nameBytes.length, true);
        localView.setUint16(28, 0, true);
        localHeader.set(nameBytes, 30);
        localChunks.push(localHeader, dataBytes);

        const centralHeader = new Uint8Array(46 + nameBytes.length);
        const centralView = new DataView(centralHeader.buffer);
        centralView.setUint32(0, 0x02014b50, true);
        centralView.setUint16(4, 20, true);
        centralView.setUint16(6, 20, true);
        centralView.setUint16(8, generalPurposeBitFlag, true);
        centralView.setUint16(10, 0, true);
        centralView.setUint16(12, fixedZipEntryTimestamp.dosTime, true);
        centralView.setUint16(14, fixedZipEntryTimestamp.dosDate, true);
        centralView.setUint32(16, entryCrc32, true);
        centralView.setUint32(20, dataBytes.length, true);
        centralView.setUint32(24, dataBytes.length, true);
        centralView.setUint16(28, nameBytes.length, true);
        centralView.setUint16(30, 0, true);
        centralView.setUint16(32, 0, true);
        centralView.setUint16(34, 0, true);
        centralView.setUint16(36, 0, true);
        centralView.setUint32(38, 0, true);
        centralView.setUint32(42, offset, true);
        centralHeader.set(nameBytes, 46);
        centralChunks.push(centralHeader);

        offset += localHeader.length + dataBytes.length;
    }

    const centralDirectoryStart = offset;
    const centralDirectorySize = centralChunks.reduce((sum, chunk) => sum + chunk.length, 0);
    const eocd = new Uint8Array(22);
    const eocdView = new DataView(eocd.buffer);
    eocdView.setUint32(0, 0x06054b50, true);
    eocdView.setUint16(4, 0, true);
    eocdView.setUint16(6, 0, true);
    eocdView.setUint16(8, entries.length, true);
    eocdView.setUint16(10, entries.length, true);
    eocdView.setUint32(12, centralDirectorySize, true);
    eocdView.setUint32(16, centralDirectoryStart, true);
    eocdView.setUint16(20, 0, true);

    const totalLength = localChunks.reduce((sum, chunk) => sum + chunk.length, 0) + centralDirectorySize + eocd.length;
    const output = new Uint8Array(totalLength);
    let cursor = 0;
    for (const chunk of localChunks) {
        output.set(chunk, cursor);
        cursor += chunk.length;
    }
    for (const chunk of centralChunks) {
        output.set(chunk, cursor);
        cursor += chunk.length;
    }
    output.set(eocd, cursor);
    return output;
}

/** Reads ZIP packages. Writes stored export archives. */
export { unzipEntries, createStoredZip, fixedZipEntryTimestamp };
