/*
 * Minimal ZIP reader/writer for DOCX packaging in the browser.
 *
 * Exposes `window.AutoBerichtWordDocxZip`:
 * - `unzipAllEntries(arrayBuffer)` -> [{ name, data, flags }]
 * - `buildZipStore(entries)` -> Uint8Array
 *
 * Notes:
 * - Supports STORE (0) and DEFLATE (8) for input.
 * - Emits STORE-only output for deterministic export behavior.
 */
(() => {
  const textDecoder = new TextDecoder();
  const textEncoder = new TextEncoder();

  const ZIP_LOCAL_FILE_SIG = 0x04034b50;
  const ZIP_CENTRAL_FILE_SIG = 0x02014b50;
  const ZIP_EOCD_SIG = 0x06054b50;

  const crcTable = (() => {
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i += 1) {
      let c = i;
      for (let j = 0; j < 8; j += 1) {
        c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
      }
      table[i] = c >>> 0;
    }
    return table;
  })();

  const crc32 = (bytes) => {
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i += 1) {
      c = crcTable[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    }
    return (c ^ 0xffffffff) >>> 0;
  };

  const concatUint8 = (chunks) => {
    const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
    const out = new Uint8Array(total);
    let offset = 0;
    chunks.forEach((chunk) => {
      out.set(chunk, offset);
      offset += chunk.length;
    });
    return out;
  };

  const readU16 = (view, offset) => view.getUint16(offset, true);
  const readU32 = (view, offset) => view.getUint32(offset, true);

  const writeU16 = (arr, offset, value) => {
    arr[offset] = value & 0xff;
    arr[offset + 1] = (value >>> 8) & 0xff;
  };

  const writeU32 = (arr, offset, value) => {
    arr[offset] = value & 0xff;
    arr[offset + 1] = (value >>> 8) & 0xff;
    arr[offset + 2] = (value >>> 16) & 0xff;
    arr[offset + 3] = (value >>> 24) & 0xff;
  };

  const findEocdOffset = (bytes) => {
    const minEocd = 22;
    const maxComment = 0xffff;
    const start = Math.max(0, bytes.length - minEocd - maxComment);
    for (let i = bytes.length - minEocd; i >= start; i -= 1) {
      if (
        bytes[i] === 0x50
        && bytes[i + 1] === 0x4b
        && bytes[i + 2] === 0x05
        && bytes[i + 3] === 0x06
      ) {
        return i;
      }
    }
    return -1;
  };

  const inflateRaw = async (compressed) => {
    if (typeof DecompressionStream !== "function") {
      throw new Error("This browser does not support ZIP deflate decompression.");
    }
    const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    const buffer = await new Response(stream).arrayBuffer();
    return new Uint8Array(buffer);
  };

  const unzipAllEntries = async (arrayBuffer) => {
    const bytes = new Uint8Array(arrayBuffer);
    const view = new DataView(arrayBuffer);
    const eocd = findEocdOffset(bytes);
    if (eocd < 0) throw new Error("Invalid zip: EOCD not found");

    const centralDirSize = readU32(view, eocd + 12);
    const centralDirOffset = readU32(view, eocd + 16);
    const centralDirEnd = centralDirOffset + centralDirSize;

    const entries = [];
    let cursor = centralDirOffset;
    while (cursor < centralDirEnd) {
      const sig = readU32(view, cursor);
      if (sig !== ZIP_CENTRAL_FILE_SIG) throw new Error("Invalid zip: central directory signature mismatch");

      const flags = readU16(view, cursor + 8);
      const method = readU16(view, cursor + 10);
      const compressedSize = readU32(view, cursor + 20);
      const nameLen = readU16(view, cursor + 28);
      const extraLen = readU16(view, cursor + 30);
      const commentLen = readU16(view, cursor + 32);
      const localHeaderOffset = readU32(view, cursor + 42);

      const nameBytes = bytes.slice(cursor + 46, cursor + 46 + nameLen);
      const name = textDecoder.decode(nameBytes);

      const localSig = readU32(view, localHeaderOffset);
      if (localSig !== ZIP_LOCAL_FILE_SIG) throw new Error(`Invalid zip: local header mismatch for ${name}`);
      const localNameLen = readU16(view, localHeaderOffset + 26);
      const localExtraLen = readU16(view, localHeaderOffset + 28);
      const dataStart = localHeaderOffset + 30 + localNameLen + localExtraLen;
      const compressed = bytes.slice(dataStart, dataStart + compressedSize);

      let data;
      if (method === 0) {
        data = compressed;
      } else if (method === 8) {
        // eslint-disable-next-line no-await-in-loop
        data = await inflateRaw(compressed);
      } else {
        throw new Error(`Unsupported zip method ${method} for ${name}`);
      }

      entries.push({ name, data, flags });
      cursor += 46 + nameLen + extraLen + commentLen;
    }

    return entries;
  };

  const buildZipStore = (entries) => {
    const localParts = [];
    const centralParts = [];
    let localOffset = 0;

    const sorted = [...entries].sort((a, b) => a.name.localeCompare(b.name, "en", { numeric: true }));

    sorted.forEach((entry) => {
      const nameBytes = textEncoder.encode(entry.name);
      const data = entry.data instanceof Uint8Array ? entry.data : new Uint8Array(entry.data);
      const crc = crc32(data);

      const localHeader = new Uint8Array(30 + nameBytes.length);
      writeU32(localHeader, 0, ZIP_LOCAL_FILE_SIG);
      writeU16(localHeader, 4, 20);
      writeU16(localHeader, 6, (entry.flags || 0) & 0x0800); // Keep UTF-8 bit only.
      writeU16(localHeader, 8, 0); // Store
      writeU16(localHeader, 10, 0);
      writeU16(localHeader, 12, 0);
      writeU32(localHeader, 14, crc);
      writeU32(localHeader, 18, data.length);
      writeU32(localHeader, 22, data.length);
      writeU16(localHeader, 26, nameBytes.length);
      writeU16(localHeader, 28, 0);
      localHeader.set(nameBytes, 30);

      localParts.push(localHeader, data);

      const centralHeader = new Uint8Array(46 + nameBytes.length);
      writeU32(centralHeader, 0, ZIP_CENTRAL_FILE_SIG);
      writeU16(centralHeader, 4, 20);
      writeU16(centralHeader, 6, 20);
      writeU16(centralHeader, 8, (entry.flags || 0) & 0x0800);
      writeU16(centralHeader, 10, 0);
      writeU16(centralHeader, 12, 0);
      writeU16(centralHeader, 14, 0);
      writeU32(centralHeader, 16, crc);
      writeU32(centralHeader, 20, data.length);
      writeU32(centralHeader, 24, data.length);
      writeU16(centralHeader, 28, nameBytes.length);
      writeU16(centralHeader, 30, 0);
      writeU16(centralHeader, 32, 0);
      writeU16(centralHeader, 34, 0);
      writeU16(centralHeader, 36, 0);
      writeU32(centralHeader, 38, 0);
      writeU32(centralHeader, 42, localOffset);
      centralHeader.set(nameBytes, 46);
      centralParts.push(centralHeader);

      localOffset += localHeader.length + data.length;
    });

    const localBlob = concatUint8(localParts);
    const centralBlob = concatUint8(centralParts);

    const eocd = new Uint8Array(22);
    writeU32(eocd, 0, ZIP_EOCD_SIG);
    writeU16(eocd, 4, 0);
    writeU16(eocd, 6, 0);
    writeU16(eocd, 8, sorted.length);
    writeU16(eocd, 10, sorted.length);
    writeU32(eocd, 12, centralBlob.length);
    writeU32(eocd, 16, localBlob.length);
    writeU16(eocd, 20, 0);

    return concatUint8([localBlob, centralBlob, eocd]);
  };

  window.AutoBerichtWordDocxZip = {
    unzipAllEntries,
    buildZipStore,
  };
})();
