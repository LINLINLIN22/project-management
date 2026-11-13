/**
 * IssueDesk - Pure Vanilla JavaScript XLSX (Excel OpenXML) & ZIP Engine
 * Zero dependencies, no frameworks, no ES module imports required.
 * Works seamlessly both on static HTTP servers (GitHub Pages) and local file:// opening.
 */

(function (global) {
  'use strict';

  // CRC-32 table
  const CRC32_TABLE = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    }
    CRC32_TABLE[i] = c;
  }

  function computeCRC32(bytes) {
    let crc = 0 ^ (-1);
    for (let i = 0; i < bytes.length; i++) {
      crc = (crc >>> 8) ^ CRC32_TABLE[(crc ^ bytes[i]) & 0xFF];
    }
    return (crc ^ (-1)) >>> 0;
  }

  function colToLetter(colIndex) {
    let temp = colIndex;
    let letter = '';
    while (temp >= 0) {
      letter = String.fromCharCode((temp % 26) + 65) + letter;
      temp = Math.floor(temp / 26) - 1;
    }
    return letter;
  }

  function escapeXml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function getElementsByTag(parent, localName) {
    if (!parent) return [];
    const byLocal = parent.getElementsByTagNameNS ? parent.getElementsByTagNameNS('*', localName) : null;
    if (byLocal && byLocal.length > 0) return byLocal;
    return parent.getElementsByTagName(localName);
  }

  /**
   * Convert Excel serial date number (e.g. 46337.5) to ISO string
   */
  function excelSerialToIso(serial) {
    const num = Number(serial);
    if (!isFinite(num) || num < 1000) return null;
    const utcDays = Math.floor(num - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);
    const fractionalDay = num - Math.floor(num) + 0.0000001;
    let totalSeconds = Math.floor(86400 * fractionalDay);
    const seconds = totalSeconds % 60;
    totalSeconds -= seconds;
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor(totalSeconds / 60) % 60;
    const d = new Date(Date.UTC(
      dateInfo.getUTCFullYear(),
      dateInfo.getUTCMonth(),
      dateInfo.getUTCDate(),
      hours,
      minutes,
      seconds
    ));
    return isNaN(d.getTime()) ? null : d.toISOString();
  }

  /**
   * Creates a valid ZIP archive Uint8Array from { fileName: Uint8Array | string }
   */
  function createZipArchive(filesMap) {
    const fileEntries = [];
    let localOffset = 0;
    const encoder = new TextEncoder();

    for (const [filename, content] of Object.entries(filesMap)) {
      const fileBytes = typeof content === 'string' ? encoder.encode(content) : content;
      const nameBytes = encoder.encode(filename);
      const crc = computeCRC32(fileBytes);
      const size = fileBytes.length;

      const header = new Uint8Array(30 + nameBytes.length);
      const view = new DataView(header.buffer);

      view.setUint32(0, 0x04034b50, true);
      view.setUint16(4, 20, true);
      view.setUint16(6, 0x0800, true); // UTF-8 filename flag
      view.setUint16(8, 0, true);      // Store (uncompressed)
      view.setUint16(10, 0x5421, true);
      view.setUint16(12, 0x5421, true);
      view.setUint32(14, crc, true);
      view.setUint32(18, size, true);
      view.setUint32(22, size, true);
      view.setUint16(26, nameBytes.length, true);
      view.setUint16(28, 0, true);

      header.set(nameBytes, 30);

      fileEntries.push({
        nameBytes,
        header,
        fileBytes,
        crc,
        size,
        offset: localOffset,
      });

      localOffset += header.length + fileBytes.length;
    }

    const cdRecords = [];
    let cdSize = 0;

    for (const entry of fileEntries) {
      const cdHeader = new Uint8Array(46 + entry.nameBytes.length);
      const view = new DataView(cdHeader.buffer);

      view.setUint32(0, 0x02014b50, true);
      view.setUint16(4, 20, true);
      view.setUint16(6, 20, true);
      view.setUint16(8, 0x0800, true);
      view.setUint16(10, 0, true);
      view.setUint16(12, 0x5421, true);
      view.setUint16(14, 0x5421, true);
      view.setUint32(16, entry.crc, true);
      view.setUint32(20, entry.size, true);
      view.setUint32(24, entry.size, true);
      view.setUint16(28, entry.nameBytes.length, true);
      view.setUint16(30, 0, true);
      view.setUint16(32, 0, true);
      view.setUint16(34, 0, true);
      view.setUint16(36, 0, true);
      view.setUint32(38, 0, true);
      view.setUint32(42, entry.offset, true);

      cdHeader.set(entry.nameBytes, 46);
      cdRecords.push(cdHeader);
      cdSize += cdHeader.length;
    }

    const eocd = new Uint8Array(22);
    const eocdView = new DataView(eocd.buffer);
    eocdView.setUint32(0, 0x06054b50, true);
    eocdView.setUint16(4, 0, true);
    eocdView.setUint16(6, 0, true);
    eocdView.setUint16(8, fileEntries.length, true);
    eocdView.setUint16(10, fileEntries.length, true);
    eocdView.setUint32(12, cdSize, true);
    eocdView.setUint32(16, localOffset, true);
    eocdView.setUint16(20, 0, true);

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
   * Builds standard OpenXML XLSX file Uint8Array from sheets data
   */
  function buildXlsxBuffer(sheetsData) {
    const files = {};

    let contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-officedocument.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`;

    sheetsData.forEach((_, idx) => {
      contentTypesXml += `\n  <Override PartName="/xl/worksheets/sheet${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    });
    contentTypesXml += `\n</Types>`;
    files['[Content_Types].xml'] = contentTypesXml;

    files['_rels/.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

    let wbRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    sheetsData.forEach((_, idx) => {
      wbRelsXml += `\n  <Relationship Id="rId${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${idx + 1}.xml"/>`;
    });
    wbRelsXml += `\n</Relationships>`;
    files['xl/_rels/workbook.xml.rels'] = wbRelsXml;

    let wbXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>`;
    const usedSheetNames = new Set();
    sheetsData.forEach((sheet, idx) => {
      let baseName = (sheet.name || `Sheet${idx + 1}`).replace(/[/\\?*:[\]]/g, '_').trim().substring(0, 28) || `Sheet${idx + 1}`;
      let uniqueName = baseName;
      let suffix = 2;
      while (usedSheetNames.has(uniqueName.toLowerCase())) {
        uniqueName = `${baseName.substring(0, 25)}_${suffix++}`;
      }
      usedSheetNames.add(uniqueName.toLowerCase());
      wbXml += `\n    <sheet name="${escapeXml(uniqueName)}" sheetId="${idx + 1}" r:id="rId${idx + 1}"/>`;
    });
    wbXml += `\n  </sheets>
</workbook>`;
    files['xl/workbook.xml'] = wbXml;

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
            // Skip empty cells
          } else if (typeof cellVal === 'number' && isFinite(cellVal)) {
            sheetXml += `<c r="${cellRef}"><v>${cellVal}</v></c>`;
          } else {
            sheetXml += `<c r="${cellRef}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(String(cellVal))}</t></is></c>`;
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

  function downloadFile(buffer, filename) {
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
   * Deadlock-free deflate-raw decompression using native DecompressionStream
   */
  async function decompressDeflateRaw(bytes) {
    if (typeof DecompressionStream !== 'undefined') {
      try {
        const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
        const ab = await new Response(stream).arrayBuffer();
        return new Uint8Array(ab);
      } catch (e) {
        // Fallback if stream fails
      }
    }
    return bytes;
  }

  /**
   * Parse ZIP archive into { normalizedPath: Uint8Array }
   */
  async function readZipArchive(arrayBuffer) {
    const bytes = new Uint8Array(arrayBuffer);
    const view = new DataView(arrayBuffer);
    const files = {};

    let eocdOffset = -1;
    const minOffset = Math.max(0, bytes.length - 65557);
    for (let i = bytes.length - 22; i >= minOffset; i--) {
      if (view.getUint32(i, true) === 0x06054b50) {
        eocdOffset = i;
        break;
      }
    }

    const decoder = new TextDecoder();

    if (eocdOffset !== -1) {
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

        const fileName = decoder.decode(bytes.subarray(curCd + 46, curCd + 46 + nameLen)).replace(/\\/g, '/');

        if (localHeaderOffset + 30 <= bytes.length) {
          const localNameLen = view.getUint16(localHeaderOffset + 26, true);
          const localExtraLen = view.getUint16(localHeaderOffset + 28, true);
          const dataStart = localHeaderOffset + 30 + localNameLen + localExtraLen;
          if (dataStart + compSize <= bytes.length) {
            const compBytes = bytes.subarray(dataStart, dataStart + compSize);
            if (compMethod === 0) {
              files[fileName] = compBytes;
            } else if (compMethod === 8) {
              files[fileName] = await decompressDeflateRaw(compBytes);
            }
          }
        }

        curCd += 46 + nameLen + extraLen + commentLen;
      }
    } else {
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

        const fileName = decoder.decode(bytes.subarray(offset + 30, offset + 30 + nameLen)).replace(/\\/g, '/');
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

  function extractSharedStringText(siNode) {
    let text = '';
    const children = siNode.childNodes;
    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      const localName = child.localName || child.nodeName;
      if (localName === 't') {
        text += child.textContent || '';
      } else if (localName === 'r') {
        const tNodes = getElementsByTag(child, 't');
        for (let j = 0; j < tNodes.length; j++) {
          text += tNodes[j].textContent || '';
        }
      }
      // Ignore phonetic guide <rPh> and <phoneticPr>
    }
    return text;
  }

  /**
   * Parses XLSX ArrayBuffer into [{ name, rows }]
   */
  async function parseXlsx(arrayBuffer) {
    const files = await readZipArchive(arrayBuffer);
    const decoder = new TextDecoder();
    const parser = new DOMParser();

    // Case-insensitive file lookup helper
    const getFileBytes = (targetPath) => {
      if (files[targetPath]) return files[targetPath];
      const lower = targetPath.toLowerCase();
      for (const k of Object.keys(files)) {
        if (k.toLowerCase() === lower) return files[k];
      }
      return null;
    };

    // 1. Shared Strings Table
    const sharedStrings = [];
    const sstBytes = getFileBytes('xl/sharedStrings.xml');
    if (sstBytes) {
      const sstXml = decoder.decode(sstBytes);
      const sstDoc = parser.parseFromString(sstXml, 'application/xml');
      const siNodes = getElementsByTag(sstDoc, 'si');
      for (let i = 0; i < siNodes.length; i++) {
        sharedStrings.push(extractSharedStringText(siNodes[i]));
      }
    }

    // 2. Workbook sheets & relations
    const sheetsMeta = [];
    const wbBytes = getFileBytes('xl/workbook.xml');
    if (wbBytes) {
      const wbXml = decoder.decode(wbBytes);
      const wbDoc = parser.parseFromString(wbXml, 'application/xml');
      const sheetNodes = getElementsByTag(wbDoc, 'sheet');

      const relsMap = {};
      const relsBytes = getFileBytes('xl/_rels/workbook.xml.rels');
      if (relsBytes) {
        const relsXml = decoder.decode(relsBytes);
        const relsDoc = parser.parseFromString(relsXml, 'application/xml');
        const relNodes = getElementsByTag(relsDoc, 'Relationship');
        for (let i = 0; i < relNodes.length; i++) {
          const id = relNodes[i].getAttribute('Id');
          const target = relNodes[i].getAttribute('Target');
          if (id && target) {
            relsMap[id] = target.replace(/^\/?xl\//i, '').replace(/^\.\//, '');
          }
        }
      }

      for (let i = 0; i < sheetNodes.length; i++) {
        const name = sheetNodes[i].getAttribute('name') || `Sheet${i + 1}`;
        const rId =
          sheetNodes[i].getAttribute('r:id') ||
          sheetNodes[i].getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') ||
          sheetNodes[i].getAttribute('id');
        const targetPath = relsMap[rId] || `worksheets/sheet${i + 1}.xml`;
        sheetsMeta.push({
          name,
          filePath: targetPath.toLowerCase().startsWith('xl/') ? targetPath : `xl/${targetPath}`,
        });
      }
    }

    // 3. Parse worksheets
    const sheets = [];
    for (const meta of sheetsMeta) {
      const fileBytes = getFileBytes(meta.filePath);
      if (!fileBytes) continue;

      const sheetXml = decoder.decode(fileBytes);
      const sheetDoc = parser.parseFromString(sheetXml, 'application/xml');
      const rowNodes = getElementsByTag(sheetDoc, 'row');
      const rows = [];

      for (let r = 0; r < rowNodes.length; r++) {
        const rowNode = rowNodes[r];
        const cellNodes = getElementsByTag(rowNode, 'c');
        const rowCells = [];

        for (let c = 0; c < cellNodes.length; c++) {
          const cellNode = cellNodes[c];
          const cellRef = cellNode.getAttribute('r') || '';
          const cellType = cellNode.getAttribute('t');

          let colIndex = c;
          if (cellRef) {
            const colLetters = cellRef.replace(/[0-9]/g, '');
            if (colLetters) {
              colIndex = 0;
              for (let k = 0; k < colLetters.length; k++) {
                colIndex = colIndex * 26 + (colLetters.charCodeAt(k) - 64);
              }
              colIndex = Math.max(0, colIndex - 1);
            }
          }

          let cellValue = '';
          if (cellType === 's') {
            const vNode = getElementsByTag(cellNode, 'v')[0];
            if (vNode) {
              const idx = parseInt(vNode.textContent || '', 10);
              cellValue = sharedStrings[idx] !== undefined ? sharedStrings[idx] : '';
            }
          } else if (cellType === 'inlineStr') {
            const isNode = getElementsByTag(cellNode, 'is')[0];
            if (isNode) {
              const tNodes = getElementsByTag(isNode, 't');
              for (let t = 0; t < tNodes.length; t++) {
                cellValue += tNodes[t].textContent || '';
              }
            }
          } else {
            const vNode = getElementsByTag(cellNode, 'v')[0];
            if (vNode) {
              cellValue = vNode.textContent || '';
            }
          }

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

  global.XlsxEngine = {
    computeCRC32,
    createZipArchive,
    buildXlsxBuffer,
    downloadFile,
    readZipArchive,
    parseXlsx,
    excelSerialToIso,
  };
})(window);
