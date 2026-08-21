/* A very small .xlsx writer.
   An xlsx file is a zip of XML parts. We build the few parts Excel needs,
   store them uncompressed, and zip them by hand - so the page needs no
   libraries and works as a static site. */
(function (global) {
  "use strict";

  const TABLE = (() => {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[i] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++)
      c = TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  const enc = (s) => new TextEncoder().encode(s);

  function zip(files) {
    const chunks = [], central = [];
    let offset = 0;
    for (const f of files) {
      const name = enc(f.name), data = f.data;
      const crc = crc32(data);
      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true);
      local.setUint16(4, 20, true);
      local.setUint16(6, 0, true);
      local.setUint16(8, 0, true);          // stored, no compression
      local.setUint16(10, 0, true);
      local.setUint16(12, 0, true);
      local.setUint32(14, crc, true);
      local.setUint32(18, data.length, true);
      local.setUint32(22, data.length, true);
      local.setUint16(26, name.length, true);
      local.setUint16(28, 0, true);
      chunks.push(new Uint8Array(local.buffer), name, data);

      const cen = new DataView(new ArrayBuffer(46));
      cen.setUint32(0, 0x02014b50, true);
      cen.setUint16(4, 20, true);
      cen.setUint16(6, 20, true);
      cen.setUint16(10, 0, true);
      cen.setUint32(16, crc, true);
      cen.setUint32(20, data.length, true);
      cen.setUint32(24, data.length, true);
      cen.setUint16(28, name.length, true);
      cen.setUint32(42, offset, true);
      central.push(new Uint8Array(cen.buffer), name);
      offset += 30 + name.length + data.length;
    }
    let censize = 0;
    for (const c of central) censize += c.length;
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true);
    end.setUint32(12, censize, true);
    end.setUint32(16, offset, true);

    const all = chunks.concat(central, [new Uint8Array(end.buffer)]);
    let total = 0;
    for (const a of all) total += a.length;
    const out = new Uint8Array(total);
    let p = 0;
    for (const a of all) { out.set(a, p); p += a.length; }
    return out;
  }

  const esc = (s) => String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");

  function colName(n) {
    let s = "";
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    return s;
  }

  /* opts: {sheet, columns:[{header,width}], rows:[[..]], tickColumns:[i,..]} */
  function build(opts) {
    const cols = opts.columns, rows = opts.rows;
    const last = rows.length + 1;

    const colsXml = cols.map((c, i) =>
      `<col min="${i + 1}" max="${i + 1}" width="${c.width || 18}" customWidth="1"/>`).join("");

    const rowXml = (cells, r, style) => `<row r="${r}">` + cells.map((v, i) => {
      const ref = colName(i + 1) + r;
      return `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
    }).join("") + `</row>`;

    let sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetPr/><dimension ref="A1:${colName(cols.length)}${last}"/>
<sheetViews><sheetView workbookViewId="0" tabSelected="1">
<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
</sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols>${colsXml}</cols><sheetData>`;
    sheet += rowXml(cols.map(c => c.header), 1, 1);
    rows.forEach((r, i) => { sheet += rowXml(r, i + 2, 2); });
    sheet += `</sheetData>`;
    sheet += `<autoFilter ref="A1:${colName(cols.length)}${last}"/>`;

    (opts.tickColumns || []).forEach((idx, n) => {
      const L = colName(idx + 1);
      sheet += `<conditionalFormatting sqref="${L}2:${L}${last}">` +
        `<cfRule type="cellIs" dxfId="0" priority="${n + 2}" operator="equal">` +
        `<formula>"Yes"</formula></cfRule></conditionalFormatting>`;
    });
    if ((opts.tickColumns || []).length) {
      const L = colName(opts.tickColumns[0] + 1);
      sheet += `<conditionalFormatting sqref="A2:${colName(cols.length)}${last}">` +
        `<cfRule type="expression" dxfId="1" priority="1">` +
        `<formula>EXACT($${L}2,"Yes")</formula></cfRule></conditionalFormatting>`;
    }
    if ((opts.tickColumns || []).length) {
      sheet += `<dataValidations count="${opts.tickColumns.length}">` +
        opts.tickColumns.map(idx => {
          const L = colName(idx + 1);
          return `<dataValidation type="list" allowBlank="1" showInputMessage="1"` +
            ` showErrorMessage="1" sqref="${L}2:${L}${last}">` +
            `<formula1>"Yes,No"</formula1></dataValidation>`;
        }).join("") + `</dataValidations>`;
    }
    sheet += `</worksheet>`;

    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2">
<font><sz val="11"/><name val="Iowan Old Style"/></font>
<font><b/><sz val="11"/><color rgb="FFFDFDFB"/><name val="Iowan Old Style"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF3B3833"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFAF9F5"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2"><border/>
<border><left/><right/><top/><bottom><color rgb="FFE6E3DB"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3">
<xf xfId="0"/>
<xf xfId="0" fontId="1" fillId="2" applyFont="1" applyFill="1"><alignment vertical="center" wrapText="1"/></xf>
<xf xfId="0" fontId="0" fillId="3" borderId="1" applyFont="1" applyFill="1" applyBorder="1"><alignment vertical="top"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="2">
<dxf><font><b/></font><fill><patternFill><bgColor rgb="FFCFE3D4"/></patternFill></fill></dxf>
<dxf><fill><patternFill><bgColor rgb="FFEDF4EE"/></patternFill></fill></dxf>
</dxfs>
</styleSheet>`;

    const files = [
      { name: "[Content_Types].xml", data: enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`) },
      { name: "_rels/.rels", data: enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`) },
      { name: "xl/workbook.xml", data: enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="${esc(opts.sheet || "Sheet1")}" sheetId="1" r:id="rId1"/></sheets>
</workbook>`) },
      { name: "xl/_rels/workbook.xml.rels", data: enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`) },
      { name: "xl/styles.xml", data: enc(styles) },
      { name: "xl/worksheets/sheet1.xml", data: enc(sheet) },
    ];
    return new Blob([zip(files)], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
  }

  global.miniXlsx = { build };
})(window);
