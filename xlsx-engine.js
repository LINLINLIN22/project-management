/**
 * Pure Vanilla JavaScript XLSX (Excel OpenXML) & ZIP Engine
 * Zero dependencies, no third-party libraries.
 * Uses native Web APIs: DataView, Uint8Array, CompressionStream/DecompressionStream, DOMParser.
 */

// CRC-32 table and computation
const CRC32_TABLE = new Uint32Array(256);
for (let i = 0; i < 256; i++) {
  let c = i;
  for (let k = 0; k < 8; k++) {
    c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
  }
  CRC32_TABLE[i] = c;
}

export function computeCRC32(bytes) {
  let crc = 0 ^ (-1);
  for (let i = 0; i < bytes.length; i++) {
    crc = (crc >>> 8) ^ CRC32_TABLE[(crc ^ bytes[i]) & 0xFF];
  }
  return (crc ^ (-1)) >>> 0;
}

// Convert column index (0-based) to Excel column letters (A, B, ... Z, AA, AB...)
function colToLetter(colIndex) {
  let temp = colIndex;
  let letter = '';
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

// Escape XML special characters
function escapeXml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Creates a valid ZIP archive as a Uint8Array from a map of { fileName: Uint8Array | string }
 * Uses compression method 0 (Store) or deflate if supported, fully standard and recognized by Excel.
 */
export function createZipArchive(filesMap) {
  const fileEntries = [];
  let localOffset = 0;

  const encoder = new TextEncoder();

  // First pass: generate local file records
  for (const [filename, content] of Object.entries(filesMap)) {
    const fileBytes = typeof content === 'string' ? encoder.encode(content) : content;
    const nameBytes = encoder.encode(filename);
    const crc = computeCRC32(fileBytes);
    const size = fileBytes.length;

    // Local file header (30 bytes + name length + data)
    const header = new Uint8Array(30 + nameBytes.length);
    const view = new DataView(header.buffer);

    view.setUint32(0, 0x04034b50, true); // Local header signature
    view.setUint16(4, 20, true); // Version needed (2.0)
    view.setUint16(6, 0, true);  // General purpose flags
    view.setUint16(8, 0, true);  // Compression method (0 = Store / uncompressed)
    view.setUint16(10, 0x5421, true); // Last mod time
    view.setUint16(12, 0x5421, true); // Last mod date
    view.setUint32(14, crc, true); // CRC-32
    view.setUint32(18, size, true); // Compressed size
    view.setUint32(22, size, true); // Uncompressed size
    view.setUint16(26, nameBytes.length, true); // File name length
    view.setUint16(28, 0, true); // Extra field length

    header.set(nameBytes, 30);

    fileEntries.push({
      filename,
      nameBytes,
      header,
      fileBytes,
      crc,
      size,
      offset: localOffset,
    });

    localOffset += header.length + fileBytes.length;
  }

  // Second pass: Central Directory records
  const cdRecords = [];
  let cdSize = 0;

  for (const entry of fileEntries) {
    const cdHeader = new Uint8Array(46 + entry.nameBytes.length);
    const view = new DataView(cdHeader.buffer);

    view.setUint32(0, 0x02014b50, true); // Central directory signature
    view.setUint16(4, 20, true); // Version made by
    view.setUint16(6, 20, true); // Version needed
    view.setUint16(8, 0, true);  // Flags
    view.setUint16(10, 0, true); // Compression method (Store)
    view.setUint16(12, 0x5421, true); // Mod time
    view.setUint16(14, 0x5421, true); // Mod date
    view.setUint32(16, entry.crc, true); // CRC-32
    view.setUint32(20, entry.size, true); // Compressed size
    view.setUint32(24, entry.size, true); // Uncompressed size
    view.setUint16(28, entry.nameBytes.length, true); // Name length
    view.setUint16(30, 0, true); // Extra field length
    view.setUint16(32, 0, true); // File comment length
    view.setUint16(34, 0, true); // Disk number
    view.setUint16(36, 0, true); // Internal attributes
    view.setUint32(38, 0, true); // External attributes
    view.setUint32(42, entry.offset, true); // Local header offset

    cdHeader.set(entry.nameBytes, 46);
    cdRecords.push(cdHeader);
    cdSize += cdHeader.length;
  }

  // End of Central Directory Record (22 bytes)
  const eocd = new Uint8Array(22);
  const eocdView = new DataView(eocd.buffer);
  eocdView.setUint32(0, 0x06054b50, true); // EOCD signature
  eocdView.setUint16(4, 0, true); // Disk number
  eocdView.setUint16(6, 0, true); // Disk with CD
  eocdView.setUint16(8, fileEntries.length, true); // Number of entries on disk
  eocdView.setUint16(10, fileEntries.length, true); // Total entries
  eocdView.setUint32(12, cdSize, true); // Size of central directory
  eocdView.setUint32(16, localOffset, true); // Offset of central directory
  eocdView.setUint16(20, 0, true); // Comment length

  // Combine everything into a single Uint8Array
  const totalLength = localOffset + cdSize + 22;
  const zipBuffer = new Uint8Array(totalLength);
  let cur = 0;

  for (const entry of fileEntries) {
    zipBuffer.set(entry.header, cur);
    cur += entry.header.length;
    zipBuffer.set(entry.fileBytes, cur);
    cur += entry.fileBytes.length;
  }

  for (const cd of cdRecords) {
    zipBuffer.set(cd, cur);
    cur += cd.length;
  }

  zipBuffer.set(eocd, cur);

  return zipBuffer;
}

/**
 * Builds standard OpenXML XLSX file content from sheets data.
 * @param {Array<{ name: string, rows: Array<Array<any>> }>} sheetsData
 * @returns {Uint8Array}
 */
export function buildXlsxBuffer(sheetsData) {
  const files = {};

  // 1. [Content_Types].xml
  let contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`;

  sheetsData.forEach((_, idx) => {
    contentTypesXml += `\n  <Override PartName="/xl/worksheets/sheet${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  });
  contentTypesXml += `\n</Types>`;
  files['[Content_Types].xml'] = contentTypesXml;

  // 2. _rels/.rels
  files['_rels/.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  // 3. xl/_rels/workbook.xml.rels
  let wbRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
  sheetsData.forEach((_, idx) => {
    wbRelsXml += `\n  <Relationship Id="rId${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${idx + 1}.xml"/>`;
  });
  wbRelsXml += `\n</Relationships>`;
  files['xl/_rels/workbook.xml.rels'] = wbRelsXml;

  // 4. xl/workbook.xml
  let wbXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>`;
  sheetsData.forEach((sheet, idx) => {
    const cleanSheetName = escapeXml(sheet.name.replace(/[/\\?*:[\]]/g, '_').substring(0, 31) || `Sheet${idx + 1}`);
    wbXml += `\n    <sheet name="${cleanSheetName}" sheetId="${idx + 1}" r:id="rId${idx + 1}"/>`;
  });
  wbXml += `\n  </sheets>
</workbook>`;
  files['xl/workbook.xml'] = wbXml;

  // 5. xl/worksheets/sheetN.xml
  sheetsData.forEach((sheet, sheetIdx) => {
    let sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>`;

    sheet.rows.forEach((row, rowIdx) => {
      const rowNum = rowIdx + 1;
      sheetXml += `\n    <row r="${rowNum}">`;
      row.forEach((cellVal, colIdx) => {
        const cellRef = `${colToLetter(colIdx)}${rowNum}`;
        if (cellVal === null || cellVal === undefined || cellVal === '') {
          // empty cell
        } else if (typeof cellVal === 'number') {
          sheetXml += `<c r="${cellRef}"><v>${cellVal}</v></c>`;
        } else {
          sheetXml += `<c r="${cellRef}" t="inlineStr"><is><t>${escapeXml(String(cellVal))}</t></is></c>`;
        }
      });
      sheetXml += `</row>`;
    });

    sheetXml += `\n  </sheetData>
</worksheet>`;
    files[`xl/worksheets/sheet${sheetIdx + 1}.xml`] = sheetXml;
  });

  return createZipArchive(files);
}

/**
 * Triggers a client-side file download of a Uint8Array or Blob
 */
export function downloadFile(buffer, filename) {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/**
 * Decompress deflate-raw chunk using native DecompressionStream
 */
async function decompressDeflateRaw(bytes) {
  if (typeof DecompressionStream !== 'undefined') {
    try {
      const ds = new DecompressionStream('deflate-raw');
      const writer = ds.writable.getWriter();
      writer.write(bytes);
      writer.close();
      const reader = ds.readable.getReader();
      const chunks = [];
      let total = 0;
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        chunks.push(value);
        total += value.length;
      }
      const out = new Uint8Array(total);
      let offset = 0;
      for (const chunk of chunks) {
        out.set(chunk, offset);
        offset += chunk.length;
      }
      return out;
    } catch (e) {
      // Fallback
    }
  }
  return bytes;
}

/**
 * Parse a ZIP archive and return a map of fileName -> Uint8Array
 */
export async function readZipArchive(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const view = new DataView(arrayBuffer);
  const files = {};

  // Locate End of Central Directory Record (EOCD) from the end
  let eocdOffset = -1;
  for (let i = bytes.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocdOffset = i;
      break;
    }
  }

  const decoder = new TextDecoder();

  if (eocdOffset !== -1) {
    // Read via Central Directory
    const totalEntries = view.getUint16(eocdOffset + 10, true);
    const cdOffset = view.getUint32(eocdOffset + 16, true);
    let curCd = cdOffset;

    for (let i = 0; i < totalEntries; i++) {
      if (curCd + 46 > bytes.length) break;
      if (view.getUint32(curCd, true) !== 0x02014b50) break;

      const compMethod = view.getUint16(curCd + 10, true);
      const compSize = view.getUint32(curCd + 20, true);
      const nameLen = view.getUint16(curCd + 28, true);
      const extraLen = view.getUint16(curCd + 30, true);
      const commentLen = view.getUint16(curCd + 32, true);
      const localHeaderOffset = view.getUint32(curCd + 42, true);

      const fileName = decoder.decode(bytes.subarray(curCd + 46, curCd + 46 + nameLen));

      // Go to local header to find data offset
      if (localHeaderOffset + 30 <= bytes.length) {
        const localNameLen = view.getUint16(localHeaderOffset + 26, true);
        const localExtraLen = view.getUint16(localHeaderOffset + 28, true);
        const dataStart = localHeaderOffset + 30 + localNameLen + localExtraLen;
        const compBytes = bytes.subarray(dataStart, dataStart + compSize);

        if (compMethod === 0) {
          files[fileName] = compBytes;
        } else if (compMethod === 8) {
          const decompressed = await decompressDeflateRaw(compBytes);
          files[fileName] = decompressed;
        }
      }

      curCd += 46 + nameLen + extraLen + commentLen;
    }
  } else {
    // Fallback: scan local headers sequentially
    let offset = 0;
    while (offset + 30 < bytes.length) {
      if (view.getUint32(offset, true) !== 0x04034b50) {
        offset++;
        continue;
      }
      const compMethod = view.getUint16(offset + 8, true);
      const compSize = view.getUint32(offset + 18, true);
      const nameLen = view.getUint16(offset + 26, true);
      const extraLen = view.getUint16(offset + 28, true);

      const fileName = decoder.decode(bytes.subarray(offset + 30, offset + 30 + nameLen));
      const dataStart = offset + 30 + nameLen + extraLen;

      if (compSize > 0 && dataStart + compSize <= bytes.length) {
        const compBytes = bytes.subarray(dataStart, dataStart + compSize);
        if (compMethod === 0) {
          files[fileName] = compBytes;
        } else if (compMethod === 8) {
          files[fileName] = await decompressDeflateRaw(compBytes);
        }
        offset = dataStart + compSize;
      } else {
        offset += 30 + nameLen + extraLen;
      }
    }
  }

  return files;
}

/**
 * Parses XLSX arrayBuffer into sheets with rows of cells
 */
export async function parseXlsx(arrayBuffer) {
  const files = await readZipArchive(arrayBuffer);
  const decoder = new TextDecoder();
  const parser = new DOMParser();

  // 1. Shared Strings Table (optional)
  const sharedStrings = [];
  if (files['xl/sharedStrings.xml']) {
    const sstXml = decoder.decode(files['xl/sharedStrings.xml']);
    const sstDoc = parser.parseFromString(sstXml, 'application/xml');
    const siNodes = sstDoc.getElementsByTagName('si');
    for (let i = 0; i < siNodes.length; i++) {
      const tNodes = siNodes[i].getElementsByTagName('t');
      let str = '';
      for (let j = 0; j < tNodes.length; j++) {
        str += tNodes[j].textContent || '';
      }
      sharedStrings.push(str);
    }
  }

  // 2. Workbook sheet names and relation IDs
  const sheetsMeta = [];
  if (files['xl/workbook.xml']) {
    const wbXml = decoder.decode(files['xl/workbook.xml']);
    const wbDoc = parser.parseFromString(wbXml, 'application/xml');
    const sheetNodes = wbDoc.getElementsByTagName('sheet');

    // Parse relations to map r:id to actual file target
    const relsMap = {};
    if (files['xl/_rels/workbook.xml.rels']) {
      const relsXml = decoder.decode(files['xl/_rels/workbook.xml.rels']);
      const relsDoc = parser.parseFromString(relsXml, 'application/xml');
      const relNodes = relsDoc.getElementsByTagName('Relationship');
      for (let i = 0; i < relNodes.length; i++) {
        const id = relNodes[i].getAttribute('Id');
        const target = relNodes[i].getAttribute('Target');
        if (id && target) {
          relsMap[id] = target.replace(/^\/?xl\//, '').replace(/^\.\//, '');
        }
      }
    }

    for (let i = 0; i < sheetNodes.length; i++) {
      const name = sheetNodes[i].getAttribute('name') || `Sheet${i + 1}`;
      const rId = sheetNodes[i].getAttribute('r:id') || sheetNodes[i].getAttribute('id');
      const targetPath = relsMap[rId] || `worksheets/sheet${i + 1}.xml`;
      sheetsMeta.push({
        name,
        filePath: targetPath.startsWith('xl/') ? targetPath : `xl/${targetPath}`,
      });
    }
  }

  // 3. Parse each sheet
  const sheets = [];
  for (const meta of sheetsMeta) {
    const fileBytes = files[meta.filePath];
    if (!fileBytes) continue;

    const sheetXml = decoder.decode(fileBytes);
    const sheetDoc = parser.parseFromString(sheetXml, 'application/xml');
    const rowNodes = sheetDoc.getElementsByTagName('row');
    const rows = [];

    for (let r = 0; r < rowNodes.length; r++) {
      const rowNode = rowNodes[r];
      const cellNodes = rowNode.getElementsByTagName('c');
      const rowCells = [];

      for (let c = 0; c < cellNodes.length; c++) {
        const cellNode = cellNodes[c];
        const cellRef = cellNode.getAttribute('r') || '';
        const cellType = cellNode.getAttribute('t');

        // Extract column index from ref (e.g., "B3" -> col 1)
        const colLetters = cellRef.replace(/[0-9]/g, '');
        let colIndex = 0;
        for (let k = 0; k < colLetters.length; k++) {
          colIndex = colIndex * 26 + (colLetters.charCodeAt(k) - 64);
        }
        colIndex = Math.max(0, colIndex - 1);

        let cellValue = '';
        if (cellType === 's') {
          // Shared string index
          const vNode = cellNode.getElementsByTagName('v')[0];
          if (vNode) {
            const idx = parseInt(vNode.textContent || '', 10);
            cellValue = sharedStrings[idx] !== undefined ? sharedStrings[idx] : '';
          }
        } else if (cellType === 'inlineStr') {
          // Inline string
          const isNode = cellNode.getElementsByTagName('is')[0];
          if (isNode) {
            const tNodes = isNode.getElementsByTagName('t');
            for (let t = 0; t < tNodes.length; t++) {
              cellValue += tNodes[t].textContent || '';
            }
          }
        } else {
          // Number or raw value
          const vNode = cellNode.getElementsByTagName('v')[0];
          if (vNode) {
            cellValue = vNode.textContent || '';
          }
        }

        // Fill sparse arrays if columns were skipped
        while (rowCells.length < colIndex) {
          rowCells.push('');
        }
        rowCells[colIndex] = cellValue.trim();
      }

      if (rowCells.some((val) => val !== '')) {
        rows.push(rowCells);
      }
    }

    sheets.push({
      name: meta.name,
      rows,
    });
  }

  return sheets;
}
