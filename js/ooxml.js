(function (global) {
  'use strict';

  const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const WORKSHEET_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
  const WORKSHEET_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

  function assertRuntime() {
    if (!global.JSZip) throw new Error('JSZip no está disponible. Recarga la página.');
  }

  function decodeXml(value) {
    return String(value == null ? '' : value)
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&');
  }

  function escapeXml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function parseAttributes(tag) {
    const attrs = {};
    const regex = /([:\w-]+)\s*=\s*(["'])(.*?)\2/g;
    let match;
    while ((match = regex.exec(tag))) attrs[match[1]] = decodeXml(match[3]);
    return attrs;
  }

  function normalizeZipPath(target) {
    if (!target) return '';
    if (target.startsWith('/')) return target.slice(1);
    return target.startsWith('xl/') ? target : `xl/${target.replace(/^\.\//, '')}`;
  }

  function workbookSheetRecords(workbookXml, relsXml) {
    const relMap = new Map();
    const relRegex = /<Relationship\b[^>]*\/>/g;
    for (const tag of relsXml.match(relRegex) || []) {
      const attrs = parseAttributes(tag);
      if (attrs.Id) relMap.set(attrs.Id, attrs);
    }

    const records = [];
    const sheetRegex = /<(?:\w+:)?sheet\b[^>]*\/>/g;
    for (const tag of workbookXml.match(sheetRegex) || []) {
      const attrs = parseAttributes(tag);
      const relId = attrs['r:id'] || attrs.id || attrs['x:id'];
      const rel = relMap.get(relId);
      records.push({
        name: attrs.name,
        sheetId: Number(attrs.sheetId || 0),
        relId,
        path: rel ? normalizeZipPath(rel.Target) : '',
        tag
      });
    }
    return records;
  }


  function extractTextNodes(xml) {
    const parts = [];
    const regex = /<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g;
    let match;
    while ((match = regex.exec(xml || ''))) parts.push(decodeXml(match[1]));
    return parts.join('');
  }

  function parseSharedStrings(xml) {
    const values = [];
    const regex = /<(?:\w+:)?si\b[^>]*>([\s\S]*?)<\/(?:\w+:)?si>/g;
    let match;
    while ((match = regex.exec(xml || ''))) values.push(extractTextNodes(match[1]));
    return values;
  }

  function columnIndexFromRef(ref) {
    const letters = String(ref || '').match(/[A-Z]+/i);
    if (!letters) return 0;
    return letters[0].toUpperCase().split('').reduce((value, letter) => value * 26 + (letter.charCodeAt(0) - 64), 0) - 1;
  }

  function parseWorksheetMatrix(sheetXml, sharedStrings) {
    const matrix = [];
    const rowRegex = /<(?:\w+:)?row\b[^>]*\/\s*>|<(?:\w+:)?row\b[^>]*>[\s\S]*?<\/(?:\w+:)?row>/g;
    let rowMatch;
    let inferredRow = 0;
    while ((rowMatch = rowRegex.exec(sheetXml || ''))) {
      const rowTag = rowMatch[0];
      const rowOpenTag = (rowTag.match(/^<[^>]+>/) || [''])[0];
      const rowAttrs = parseAttributes(rowOpenTag);
      const rowIndex = Math.max(0, Number(rowAttrs.r || inferredRow + 1) - 1);
      inferredRow = rowIndex + 1;
      const row = matrix[rowIndex] || [];
      const selfClosingRow = /\/\s*>$/.test(rowTag);
      const body = selfClosingRow ? '' : rowTag.slice(rowOpenTag.length, rowTag.lastIndexOf('</'));
      const cellRegex = /<(?:\w+:)?c\b[^>]*\/\s*>|<(?:\w+:)?c\b[^>]*>[\s\S]*?<\/(?:\w+:)?c>/g;
      let cellMatch;
      let inferredColumn = 0;
      while ((cellMatch = cellRegex.exec(body))) {
        const cellTag = cellMatch[0];
        const openTag = (cellTag.match(/^<[^>]+>/) || [''])[0];
        const attrs = parseAttributes(openTag);
        const columnIndex = attrs.r ? columnIndexFromRef(attrs.r) : inferredColumn;
        inferredColumn = columnIndex + 1;
        let value = null;
        if (!/\/\s*>$/.test(cellTag)) {
          if (attrs.t === 'inlineStr') {
            value = extractTextNodes(cellTag);
          } else {
            const vMatch = cellTag.match(/<(?:\w+:)?v\b[^>]*>([\s\S]*?)<\/(?:\w+:)?v>/);
            const raw = vMatch ? decodeXml(vMatch[1]) : '';
            if (attrs.t === 's') value = sharedStrings[Number(raw)] == null ? '' : sharedStrings[Number(raw)];
            else if (attrs.t === 'str' || attrs.t === 'e') value = raw;
            else if (attrs.t === 'b') value = raw === '1';
            else if (raw !== '' && Number.isFinite(Number(raw))) value = Number(raw);
            else if (raw !== '') value = raw;
          }
        }
        row[columnIndex] = value;
      }
      matrix[rowIndex] = row;
    }
    const maxColumns = matrix.reduce((max, row) => Math.max(max, row ? row.length : 0), 0);
    return Array.from({ length: matrix.length }, (_, rowIndex) => {
      const row = matrix[rowIndex] || [];
      return Array.from({ length: maxColumns }, (_, index) => {
        const value = row[index];
        return value === undefined ? null : value;
      });
    });
  }

  async function readFirstSheet(arrayBuffer) {
    assertRuntime();
    let zip;
    try {
      zip = await global.JSZip.loadAsync(arrayBuffer);
    } catch (error) {
      throw new Error('El reporte no es un archivo XLSX válido.');
    }
    const workbookPart = zip.file('xl/workbook.xml');
    const relsPart = zip.file('xl/_rels/workbook.xml.rels');
    if (!workbookPart || !relsPart) throw new Error('El reporte Excel está incompleto.');
    const workbookXml = await workbookPart.async('string');
    const relsXml = await relsPart.async('string');
    const records = workbookSheetRecords(workbookXml, relsXml);
    if (!records.length) throw new Error('El reporte Excel no contiene hojas.');
    const first = records[0];
    const sheetPart = zip.file(first.path);
    if (!sheetPart) throw new Error('No se encontró la primera hoja del reporte.');
    const sharedPart = zip.file('xl/sharedStrings.xml');
    const sharedStrings = sharedPart ? parseSharedStrings(await sharedPart.async('string')) : [];
    return parseWorksheetMatrix(await sheetPart.async('string'), sharedStrings);
  }

  async function inspectTemplate(arrayBuffer, requiredSheets) {
    assertRuntime();
    let zip;
    try {
      zip = await global.JSZip.loadAsync(arrayBuffer);
    } catch (error) {
      throw new Error('La plantilla integrada no es un archivo XLSX válido.');
    }

    const workbookPart = zip.file('xl/workbook.xml');
    const relsPart = zip.file('xl/_rels/workbook.xml.rels');
    if (!workbookPart || !relsPart) throw new Error('La plantilla integrada está incompleta.');

    const workbookXml = await workbookPart.async('string');
    const relsXml = await relsPart.async('string');
    const records = workbookSheetRecords(workbookXml, relsXml);
    const names = records.map((item) => item.name);
    const missing = (requiredSheets || []).filter((name) => !names.includes(name));
    if (missing.length) throw new Error(`La plantilla no contiene las hojas requeridas: ${missing.join(', ')}.`);
    const unresolved = records.filter((item) => !item.path || !zip.file(item.path));
    if (unresolved.length) throw new Error(`La plantilla contiene hojas sin archivo interno: ${unresolved.map((item) => item.name).join(', ')}.`);
    return { names, records };
  }

  function cellColumnNumber(ref) {
    const letters = String(ref).match(/[A-Z]+/i);
    if (!letters) return 0;
    return letters[0].toUpperCase().split('').reduce((value, letter) => value * 26 + (letter.charCodeAt(0) - 64), 0);
  }

  function cellXml(ref, value, styleId) {
    const style = styleId == null || styleId === '' ? '' : ` s="${escapeXml(styleId)}"`;
    if (value == null || value === '') return `<x:c r="${ref}"${style}/>`;
    if (typeof value === 'number' && Number.isFinite(value)) {
      return `<x:c r="${ref}"${style}><x:v>${Number(value)}</x:v></x:c>`;
    }
    if (typeof value === 'boolean') {
      return `<x:c r="${ref}"${style} t="b"><x:v>${value ? 1 : 0}</x:v></x:c>`;
    }
    const text = String(value);
    const preserve = /^\s|\s$|\n/.test(text) ? ' xml:space="preserve"' : '';
    return `<x:c r="${ref}"${style} t="inlineStr"><x:is><x:t${preserve}>${escapeXml(text)}</x:t></x:is></x:c>`;
  }

  function replaceCell(sheetXml, ref, value) {
    const escaped = ref.replace(/[$()*+.?[\\\]^{|}-]/g, '\\$&');
    const cellPattern = new RegExp(`<(?:\\w+:)?c\\b(?=[^>]*\\br=["']${escaped}["'])[^>]*\\/\\s*>|<(?:\\w+:)?c\\b(?=[^>]*\\br=["']${escaped}["'])[^>]*>[\\s\\S]*?<\\/(?:\\w+:)?c>`);
    const existing = sheetXml.match(cellPattern);
    if (!existing) throw new Error(`La plantilla no contiene la celda ${ref}.`);
    const attrs = parseAttributes(existing[0]);
    return sheetXml.replace(cellPattern, cellXml(ref, value, attrs.s));
  }

  function writeCompanySheet(sheetXml, metrics, rows, periodLabel) {
    const changes = new Map();
    changes.set('A1', 'Valores monetarios expresados en miles');
    changes.set('B2', periodLabel);

    function metric(row, bucket) {
      changes.set(`B${row}`, bucket && bucket.value != null ? bucket.value / 1000 : null);
      changes.set(`C${row}`, bucket && bucket.operations != null ? bucket.operations : null);
    }

    function ratio(row, valueRatio, operationsRatio) {
      changes.set(`B${row}`, valueRatio == null ? null : valueRatio);
      changes.set(`C${row}`, operationsRatio == null ? null : operationsRatio);
    }

    metric(rows.TOTAL, metrics.total);
    metric(rows.ORIGINAL, metrics.original);
    metric(rows.OVERDUE, metrics.overdue);
    ratio(rows.OVERDUE_RATIO, metrics.ratios.overdueTotal, metrics.ratios.overdueOperations);
    metric(rows.ND_TOTAL, metrics.noDevenga);
    metric(rows.ND_NORMAL, metrics.noDevengaNormal);
    metric(rows.ND_RESTRUCTURED, metrics.noDevengaRestructured);
    metric(rows.ND_60_90, metrics.noDevenga6090);
    metric(rows.ND_60_90_NORMAL, metrics.noDevenga6090Normal);
    metric(rows.ND_60_90_RESTRUCTURED, metrics.noDevenga6090Restructured);
    ratio(rows.ND_RATIO_TOTAL, metrics.ratios.noDevengaTotal, metrics.ratios.noDevengaOperations);
    ratio(rows.ND_RATIO_ORIGINAL, metrics.ratios.noDevengaOriginal, metrics.ratios.noDevengaOperations);
    metric(rows.ND_OVER_90, metrics.noDevengaOver90);
    metric(rows.ND_OVER_90_NORMAL, metrics.noDevengaOver90Normal);
    metric(rows.ND_OVER_90_RESTRUCTURED, metrics.noDevengaOver90Restructured);
    ratio(rows.ND_OVER_90_RATIO, metrics.ratios.noDevengaOver90Total, metrics.ratios.noDevengaOver90Operations);

    for (let row = rows.BEP; row <= rows.BEP_NET; row += 1) {
      changes.set(`B${row}`, null);
      changes.set(`C${row}`, null);
    }
    metric(rows.CHARGED_OFF, metrics.chargedOff);

    let result = sheetXml;
    for (const [ref, value] of changes.entries()) result = replaceCell(result, ref, value);
    return result;
  }

  function clearCompanySheet(sheetXml, rows) {
    let result = replaceCell(sheetXml, 'A1', null);
    for (let row = 2; row <= rows.CHARGED_OFF; row += 1) {
      result = replaceCell(result, `B${row}`, null);
      result = replaceCell(result, `C${row}`, null);
    }
    return result;
  }

  function columnLetter(number) {
    let value = Number(number) || 1;
    let result = '';
    while (value > 0) {
      const remainder = (value - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      value = Math.floor((value - 1) / 26);
    }
    return result;
  }

  function displayDate(value) {
    if (!(value instanceof Date) || Number.isNaN(value.getTime())) return value;
    const day = String(value.getDate()).padStart(2, '0');
    const month = String(value.getMonth() + 1).padStart(2, '0');
    return `${day}/${month}/${value.getFullYear()}`;
  }

  function rowXml(rowNumber, values, options) {
    const headerRows = options && options.headerRows ? options.headerRows : [];
    const header = headerRows.includes(rowNumber);
    const title = options && options.titleRows ? options.titleRows.includes(rowNumber) : false;
    const cells = [];
    values.forEach((rawValue, index) => {
      const value = displayDate(rawValue);
      if (value == null || value === '') return;
      const ref = `${columnLetter(index + 1)}${rowNumber}`;
      const style = title ? 46 : header ? 72 : null;
      cells.push(cellXml(ref, value, style));
    });
    const height = header ? ' ht="42" customHeight="1"' : title ? ' ht="28" customHeight="1"' : '';
    return `<x:row r="${rowNumber}"${height}>${cells.join('')}</x:row>`;
  }

  function estimateWidths(matrix, maxColumns) {
    const widths = [];
    const sample = matrix.slice(0, 45);
    for (let column = 0; column < maxColumns; column += 1) {
      let maxLength = 0;
      for (const row of sample) {
        const value = displayDate((row || [])[column]);
        maxLength = Math.max(maxLength, String(value == null ? '' : value).length);
      }
      widths.push(Math.min(34, Math.max(10, maxLength + 2)));
    }
    return widths;
  }

  function worksheetXml(matrix, options) {
    const normalized = Array.isArray(matrix) ? matrix : [];
    const maxColumns = Math.max(1, ...normalized.map((row) => Array.isArray(row) ? row.length : 0));
    const widths = options && options.widths ? options.widths : estimateWidths(normalized, maxColumns);
    const cols = widths.map((width, index) => `<x:col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`).join('');
    const rows = normalized.map((row, index) => rowXml(index + 1, Array.isArray(row) ? row : [row], options || {})).join('');
    const freezeRows = options && options.freezeRows ? Number(options.freezeRows) : 0;
    const freezeColumns = options && options.freezeColumns ? Number(options.freezeColumns) : 0;
    const topLeft = `${columnLetter(freezeColumns + 1)}${freezeRows + 1}`;
    const paneAttrs = [];
    if (freezeColumns) paneAttrs.push(`xSplit="${freezeColumns}"`);
    if (freezeRows) paneAttrs.push(`ySplit="${freezeRows}"`);
    if (freezeColumns || freezeRows) paneAttrs.push(`topLeftCell="${topLeft}"`, 'state="frozen"');
    const views = paneAttrs.length
      ? `<x:sheetViews><x:sheetView workbookViewId="0"><x:pane ${paneAttrs.join(' ')}/></x:sheetView></x:sheetViews>`
      : '<x:sheetViews><x:sheetView workbookViewId="0"/></x:sheetViews>';
    const merge = options && options.merge ? `<x:mergeCells count="1"><x:mergeCell ref="${escapeXml(options.merge)}"/></x:mergeCells>` : '';
    const autoFilter = options && options.autoFilter ? `<x:autoFilter ref="${escapeXml(options.autoFilter)}"/>` : '';
    const lastCell = `${columnLetter(maxColumns)}${Math.max(1, normalized.length)}`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<x:worksheet xmlns:x="${MAIN_NS}">` +
      `<x:dimension ref="A1:${lastCell}"/>${views}` +
      `<x:sheetFormatPr defaultRowHeight="15"/><x:cols>${cols}</x:cols>` +
      `<x:sheetData>${rows}</x:sheetData>${autoFilter}${merge}` +
      `<x:pageMargins left="0.35" right="0.35" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>` +
      `</x:worksheet>`;
  }

  function insertBeforeClosing(xml, localName, fragment) {
    const regex = new RegExp(`</(?:\\w+:)?${localName}>`);
    if (!regex.test(xml)) throw new Error(`No se encontró el cierre XML ${localName}.`);
    return xml.replace(regex, `${fragment}$&`);
  }

  function removeGeneratedSheets(workbookXml, relsXml, contentTypesXml, zip) {
    const records = workbookSheetRecords(workbookXml, relsXml);
    const generated = records.filter((item) => item.name === 'CONTROL_EJECUCION' || String(item.name).startsWith('ORIGEN_'));
    for (const record of generated) {
      workbookXml = workbookXml.replace(record.tag, '');
      const relRegex = new RegExp(`<Relationship\\b(?=[^>]*\\bId=["']${record.relId.replace(/[$()*+.?[\\\]^{|}-]/g, '\\$&')}["'])[^>]*\\/>`);
      relsXml = relsXml.replace(relRegex, '');
      if (record.path) {
        zip.remove(record.path);
        const overrideRegex = new RegExp(`<Override\\b(?=[^>]*\\bPartName=["']\/${record.path.replace(/[$()*+.?[\\\]^{|}-]/g, '\\$&')}["'])[^>]*\\/>`);
        contentTypesXml = contentTypesXml.replace(overrideRegex, '');
      }
    }
    return { workbookXml, relsXml, contentTypesXml };
  }

  function addWorksheet(zip, state, name, xml) {
    state.nextSheetIndex += 1;
    state.nextSheetId += 1;
    state.nextRelIndex += 1;
    const path = `xl/worksheets/sheet${state.nextSheetIndex}.xml`;
    const relId = `rIdGenerated${state.nextRelIndex}`;
    const sheetTag = `<x:sheet name="${escapeXml(name)}" sheetId="${state.nextSheetId}" r:id="${relId}" xmlns:r="${REL_NS}"/>`;
    state.workbookXml = insertBeforeClosing(state.workbookXml, 'sheets', sheetTag);
    const relTag = `<Relationship Id="${relId}" Type="${WORKSHEET_REL_TYPE}" Target="/xl/worksheets/sheet${state.nextSheetIndex}.xml"/>`;
    state.relsXml = insertBeforeClosing(state.relsXml, 'Relationships', relTag);
    const override = `<Override PartName="/${path}" ContentType="${WORKSHEET_CONTENT_TYPE}"/>`;
    state.contentTypesXml = insertBeforeClosing(state.contentTypesXml, 'Types', override);
    zip.file(path, xml);
  }

  function controlMatrix(analyses, periodLabel, options) {
    const config = options || {};
    const scenario = config.scenario || 'NORMAL';
    const baseDate = config.baseDate instanceof Date ? config.baseDate : analyses[0].metrics.cutDate;
    const targetDate = config.targetDate instanceof Date ? config.targetDate : analyses[0].metrics.cutDate;
    const projection = config.projection || {};
    const headers = [
      'EMPRESA', 'FECHA ESCENARIO', 'ARCHIVO FUENTE', 'TOTAL CARTERA', 'OPERACIONES',
      'CARTERA VENCIDA', 'OP. VENCIDA', 'NO DEVENGA', 'OP. NO DEVENGA',
      'NO DEVENGA 60-90', 'OP. 60-90', 'NO DEVENGA +90', 'OP. +90',
      'CASTIGADA', 'OP. CASTIGADA', 'RECLASIFICADAS', 'DUPLICADOS EXCLUIDOS', 'ESTADO', 'OBSERVACIONES'
    ];
    const ruleText = scenario === 'PROYECTADO'
      ? `Dias Morosidad +${projection.horizonDays || 0}; reclasificar Por Vencer cuando el resultado sea > ${projection.thresholdDays == null ? 60 : projection.thresholdDays} días y no exista saldo previo No Devenga/Vencido.`
      : 'Lógica vigente con datos originales, sin proyección.';
    const rows = [
      [`CONTROL DE EJECUCIÓN – ESCENARIO ${scenario}`],
      ['Fecha de procesamiento', new Date().toLocaleString('es-EC')],
      ['Escenario', scenario],
      ['Fecha base', baseDate.toLocaleDateString('es-EC')],
      ['Fecha objetivo', targetDate.toLocaleDateString('es-EC')],
      ['Horizonte (días)', scenario === 'PROYECTADO' ? projection.horizonDays : 0],
      ['Umbral de migración', scenario === 'PROYECTADO' ? projection.thresholdDays : null],
      ['Regla aplicada', ruleText],
      ['Período', periodLabel],
      ['Unidad monetaria', 'Miles'],
      [],
      headers
    ];
    analyses.forEach(({ metrics }) => {
      rows.push([
        metrics.company,
        metrics.cutDate.toLocaleDateString('es-EC'),
        metrics.fileName,
        metrics.total.value / 1000,
        metrics.total.operations,
        metrics.overdue.value / 1000,
        metrics.overdue.operations,
        metrics.noDevenga.value / 1000,
        metrics.noDevenga.operations,
        metrics.noDevenga6090.value / 1000,
        metrics.noDevenga6090.operations,
        metrics.noDevengaOver90.value / 1000,
        metrics.noDevengaOver90.operations,
        metrics.chargedOff.value == null ? null : metrics.chargedOff.value / 1000,
        metrics.chargedOff.operations,
        metrics.projection ? metrics.projection.reclassifiedOperations : 0,
        metrics.duplicatesExcluded,
        metrics.status,
        metrics.warnings.join(' | ')
      ]);
    });
    return { rows, headerRow: 12 };
  }

  async function build(templateBuffer, analyses, options) {
    assertRuntime();
    const config = options || {};
    const rows = config.templateRows;
    if (!rows) throw new Error('No se recibió la estructura de filas de la plantilla.');

    const zip = await global.JSZip.loadAsync(templateBuffer.slice(0));
    const workbookPart = zip.file('xl/workbook.xml');
    const relsPart = zip.file('xl/_rels/workbook.xml.rels');
    const typesPart = zip.file('[Content_Types].xml');
    if (!workbookPart || !relsPart || !typesPart) throw new Error('La plantilla está incompleta.');

    let workbookXml = await workbookPart.async('string');
    let relsXml = await relsPart.async('string');
    let contentTypesXml = await typesPart.async('string');
    ({ workbookXml, relsXml, contentTypesXml } = removeGeneratedSheets(workbookXml, relsXml, contentTypesXml, zip));

    const records = workbookSheetRecords(workbookXml, relsXml);
    const recordMap = new Map(records.map((record) => [record.name, record]));
    const required = (config.requiredSheets || []).filter((name) => !recordMap.has(name));
    if (required.length) throw new Error(`La plantilla no contiene las hojas requeridas: ${required.join(', ')}.`);

    for (const name of config.requiredSheets || []) {
      const record = recordMap.get(name);
      const part = zip.file(record.path);
      if (!part) throw new Error(`No se encontró la hoja interna ${name}.`);
      let xml = await part.async('string');
      xml = clearCompanySheet(xml, rows);
      zip.file(record.path, xml);
    }

    for (const analysis of analyses) {
      const record = recordMap.get(analysis.company.reportSheet);
      if (!record) throw new Error(`La plantilla no contiene la hoja ${analysis.company.reportSheet}.`);
      const part = zip.file(record.path);
      let xml = await part.async('string');
      xml = writeCompanySheet(xml, analysis.metrics, rows, config.periodLabel(analysis.metrics.cutDate));
      zip.file(record.path, xml);
    }

    const worksheetIndexes = records.map((item) => {
      const match = String(item.path).match(/sheet(\d+)\.xml$/i);
      return match ? Number(match[1]) : 0;
    });
    const relIndexes = (relsXml.match(/Id=["']rIdGenerated(\d+)["']/g) || []).map((value) => Number(value.match(/(\d+)/)[1]));
    const state = {
      workbookXml,
      relsXml,
      contentTypesXml,
      nextSheetIndex: Math.max(0, ...worksheetIndexes),
      nextSheetId: Math.max(0, ...records.map((item) => item.sheetId || 0)),
      nextRelIndex: Math.max(0, ...relIndexes)
    };

    const controlInfo = controlMatrix(analyses, config.periodLabel(analyses[0].metrics.cutDate), config.control || {});
    const control = controlInfo.rows;
    const headerRow = controlInfo.headerRow;
    addWorksheet(zip, state, 'CONTROL_EJECUCION', worksheetXml(control, {
      widths: [14, 17, 38, 17, 12, 17, 12, 17, 14, 18, 12, 17, 12, 17, 13, 16, 19, 12, 52],
      titleRows: [1],
      headerRows: [headerRow],
      freezeRows: headerRow,
      freezeColumns: 1,
      merge: 'A1:S1',
      autoFilter: `A${headerRow}:S${Math.max(headerRow, control.length)}`
    }));

    for (const analysis of analyses) {
      const matrix = analysis.sourceValues || [];
      const maxCols = Math.max(1, ...matrix.map((row) => Array.isArray(row) ? row.length : 0));
      addWorksheet(zip, state, `ORIGEN_${analysis.company.code}`, worksheetXml(matrix, {
        headerRows: [5, 6],
        freezeRows: 6,
        freezeColumns: 2,
        autoFilter: matrix.length >= 6 ? `A6:${columnLetter(maxCols)}${Math.max(6, matrix.length)}` : null
      }));
    }

    zip.file('xl/workbook.xml', state.workbookXml);
    zip.file('xl/_rels/workbook.xml.rels', state.relsXml);
    zip.file('[Content_Types].xml', state.contentTypesXml);
    if (zip.file('xl/calcChain.xml')) zip.remove('xl/calcChain.xml');

    return zip.generateAsync({
      type: 'arraybuffer',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
  }

  global.OOXMLWorkbook = Object.freeze({ inspectTemplate, readFirstSheet, build });
  if (typeof module !== 'undefined' && module.exports) module.exports = global.OOXMLWorkbook;
})(typeof window !== 'undefined' ? window : globalThis);
