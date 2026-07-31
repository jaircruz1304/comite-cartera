(function (global) {
  'use strict';

  const NAVY = [14, 48, 70];
  const BLUE = [25, 77, 112];
  const GREEN = [52, 132, 101];
  const GOLD = [190, 134, 34];
  const RED = [180, 54, 67];
  const SLATE = [87, 107, 121];
  const LIGHT = [244, 247, 249];
  const LINE = [218, 228, 235];
  const WHITE = [255, 255, 255];

  function available() {
    return Boolean(global.jspdf && global.jspdf.jsPDF);
  }

  function money(value) {
    return Math.round((Number(value) || 0) / 1000).toLocaleString('es-EC', { maximumFractionDigits: 0 });
  }

  function integer(value) {
    return Math.round(Number(value) || 0).toLocaleString('es-EC', { maximumFractionDigits: 0 });
  }

  function percent(value, total) {
    return total ? `${Math.round((Number(value) || 0) / total * 100)}%` : '0%';
  }

  function aggregate(rows) {
    return rows.reduce((acc, row) => {
      Object.keys(acc).forEach((key) => { acc[key] += Number(row[key] || 0); });
      return acc;
    }, {
      total: 0,
      operations: 0,
      overdue: 0,
      overdueOperations: 0,
      noDevenga: 0,
      noDevengaOperations: 0,
      noDevenga6090: 0,
      noDevengaOver90: 0,
      chargedOff: 0,
      chargedOffOperations: 0,
      reclassified: 0
    });
  }

  function roundedRect(doc, x, y, w, h, fill, radius = 3) {
    doc.setFillColor(...fill);
    doc.roundedRect(x, y, w, h, radius, radius, 'F');
  }

  function drawHeader(doc, payload) {
    const pageWidth = doc.internal.pageSize.getWidth();
    doc.setFillColor(...NAVY);
    doc.rect(0, 0, pageWidth, 29, 'F');
    if (payload.logoDataUrl) {
      try { doc.addImage(payload.logoDataUrl, 'PNG', 13, 6, 38, 15, undefined, 'FAST'); } catch (error) {}
    } else {
      doc.setTextColor(...WHITE); doc.setFontSize(18); doc.setFont('helvetica', 'bold'); doc.text('CTH', 15, 18);
    }
    doc.setTextColor(...WHITE);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.text(`Comité de Cartera · ${payload.scenarioLabel}`, 58, 12);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.5);
    doc.text(payload.subtitle, 58, 18);
    doc.setFontSize(7.5);
    doc.text(`${payload.dateLabel}: ${payload.dateValue}`, 58, 23.5);
    doc.text('Valores monetarios expresados en miles · sin decimales', pageWidth - 14, 12, { align: 'right' });
    doc.text('Desarrolladora: Lizbeth Sanipatín', pageWidth - 14, 18, { align: 'right' });
    doc.text(`Generado: ${payload.generatedAt}`, pageWidth - 14, 23.5, { align: 'right' });
  }

  function drawFooter(doc, payload, pageNumber, totalPages) {
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    doc.setDrawColor(...LINE);
    doc.line(12, pageHeight - 10, pageWidth - 12, pageHeight - 10);
    doc.setTextColor(...SLATE);
    doc.setFontSize(7);
    doc.setFont('helvetica', 'normal');
    doc.text(`CTH · Comité de Cartera · ${payload.scenarioLabel}`, 12, pageHeight - 5.5);
    doc.text(`Página ${pageNumber} de ${totalPages}`, pageWidth - 12, pageHeight - 5.5, { align: 'right' });
  }

  function drawKpis(doc, payload, totals) {
    const cards = [
      ['TOTAL CARTERA', money(totals.total), `${integer(totals.operations)} operaciones`, NAVY],
      ['CARTERA VENCIDA', money(totals.overdue), `${percent(totals.overdue, totals.total)} del total`, RED],
      ['NO DEVENGA', money(totals.noDevenga), `${percent(totals.noDevenga, totals.total)} del total`, GOLD],
      ['NO DEVENGA +90', money(totals.noDevengaOver90), `${percent(totals.noDevengaOver90, totals.total)} del total`, BLUE],
      payload.scenario === 'PROYECTADO'
        ? ['RECLASIFICADAS', integer(totals.reclassified), 'operaciones proyectadas', GREEN]
        : ['CARTERA CASTIGADA', totals.chargedOff ? money(totals.chargedOff) : '—', totals.chargedOff ? `${integer(totals.chargedOffOperations)} operaciones` : 'Sin información', SLATE]
    ];
    const startX = 12;
    const gap = 4;
    const width = (273 - gap * 4) / 5;
    cards.forEach((card, index) => {
      const x = startX + index * (width + gap);
      roundedRect(doc, x, 35, width, 27, WHITE, 2.5);
      doc.setFillColor(...card[3]); doc.rect(x, 35, 2.5, 27, 'F');
      doc.setTextColor(...SLATE); doc.setFontSize(7); doc.setFont('helvetica', 'bold'); doc.text(card[0], x + 6, 42);
      doc.setTextColor(...NAVY); doc.setFontSize(15); doc.text(card[1], x + 6, 51);
      doc.setTextColor(...SLATE); doc.setFontSize(7); doc.setFont('helvetica', 'normal'); doc.text(card[2], x + 6, 57.5);
    });
  }

  function drawBarChart(doc, rows, y) {
    roundedRect(doc, 12, y, 171, 74, WHITE, 3);
    doc.setTextColor(...NAVY); doc.setFontSize(10); doc.setFont('helvetica', 'bold'); doc.text('Exposición por empresa', 19, y + 9);
    doc.setTextColor(...SLATE); doc.setFontSize(7.5); doc.setFont('helvetica', 'normal'); doc.text('Total cartera, No Devenga y Vencida en miles', 19, y + 14);

    const max = Math.max(1, ...rows.map((row) => Number(row.total) || 0));
    const chartX = 32;
    const chartWidth = 135;
    rows.forEach((row, index) => {
      const rowY = y + 24 + index * 11;
      doc.setTextColor(...NAVY); doc.setFontSize(8); doc.setFont('helvetica', 'bold'); doc.text(row.company, 19, rowY + 3);
      doc.setFillColor(231, 237, 241); doc.roundedRect(chartX, rowY, chartWidth, 4, 2, 2, 'F');
      doc.setFillColor(...BLUE); doc.roundedRect(chartX, rowY, Math.max(2, chartWidth * row.total / max), 4, 2, 2, 'F');
      doc.setFillColor(...GOLD); doc.roundedRect(chartX, rowY + 5, Math.max(1, chartWidth * row.noDevenga / max), 2.3, 1, 1, 'F');
      doc.setFillColor(...RED); doc.roundedRect(chartX, rowY + 8.2, Math.max(1, chartWidth * row.overdue / max), 2.3, 1, 1, 'F');
      doc.setTextColor(...SLATE); doc.setFontSize(7.5); doc.setFont('helvetica', 'normal'); doc.text(money(row.total), 176, rowY + 3, { align: 'right' });
    });
    doc.setFontSize(6.8);
    doc.setTextColor(...SLATE);
    doc.text('■ Total', 19, y + 69); doc.setTextColor(...GOLD); doc.text('■ No Devenga', 49, y + 69); doc.setTextColor(...RED); doc.text('■ Vencida', 91, y + 69);
  }

  function drawRiskPanel(doc, payload, rows, totals, y) {
    roundedRect(doc, 188, y, 97, 74, WHITE, 3);
    doc.setTextColor(...NAVY); doc.setFontSize(10); doc.setFont('helvetica', 'bold'); doc.text('Composición de riesgo', 195, y + 9);
    const metrics = [
      ['No Devenga', totals.noDevenga, GOLD],
      ['Vencida', totals.overdue, RED],
      ['ND +90', totals.noDevengaOver90, BLUE]
    ];
    metrics.forEach((metric, index) => {
      const cy = y + 24 + index * 14;
      const p = totals.total ? Math.round(metric[1] / totals.total * 100) : 0;
      doc.setDrawColor(226, 233, 238); doc.setLineWidth(2.7); doc.circle(202, cy, 5.2, 'S');
      doc.setDrawColor(...metric[2]); doc.setLineWidth(2.7);
      const segments = Math.max(1, Math.round(p / 10));
      for (let s = 0; s < Math.min(10, segments); s += 1) {
        const angle = (Math.PI * 2 * s / 10) - Math.PI / 2;
        const next = (Math.PI * 2 * (s + 0.7) / 10) - Math.PI / 2;
        doc.line(202 + Math.cos(angle) * 5.2, cy + Math.sin(angle) * 5.2, 202 + Math.cos(next) * 5.2, cy + Math.sin(next) * 5.2);
      }
      doc.setTextColor(...NAVY); doc.setFontSize(9); doc.setFont('helvetica', 'bold'); doc.text(`${p}%`, 213, cy + 1);
      doc.setTextColor(...SLATE); doc.setFontSize(7); doc.setFont('helvetica', 'normal'); doc.text(metric[0], 228, cy + 1);
    });

    const highest = [...rows].sort((a, b) => (b.noDevenga / Math.max(1, b.total)) - (a.noDevenga / Math.max(1, a.total)))[0];
    doc.setFillColor(240, 247, 244); doc.roundedRect(195, y + 62, 83, 8, 2, 2, 'F');
    doc.setTextColor(...GREEN); doc.setFontSize(6.8); doc.setFont('helvetica', 'bold');
    const summary = highest
      ? `${highest.company}: mayor participación ND (${percent(highest.noDevenga, highest.total)}).`
      : 'Sin información disponible.';
    doc.text(summary, 198, y + 67, { maxWidth: 77 });
  }

  function drawExecutiveReading(doc, payload, totals, y) {
    roundedRect(doc, 12, y, 273, 18, LIGHT, 3);
    doc.setTextColor(...NAVY); doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.text('LECTURA EJECUTIVA', 18, y + 6);
    doc.setTextColor(...SLATE); doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5);
    const sentence = payload.scenario === 'PROYECTADO'
      ? `El escenario proyectado conserva el Total Cartera y reclasifica ${integer(totals.reclassified)} operaciones según un horizonte de ${payload.horizonDays} días y umbral mayor a ${payload.thresholdDays}.`
      : `El escenario Normal refleja el corte real procesado: No Devenga representa ${percent(totals.noDevenga, totals.total)} y Cartera Vencida ${percent(totals.overdue, totals.total)} del Total Cartera.`;
    doc.text(sentence, 18, y + 12, { maxWidth: 260 });
  }

  function drawTable(doc, payload, rows) {
    const head = [['Empresa', 'Total cartera', 'Operaciones', 'Vencida', 'No Devenga', 'ND 60–90', 'ND +90', 'Castigada']];
    if (payload.scenario === 'PROYECTADO') head[0].push('Reclasificadas');
    head[0].push('Estado');
    const body = rows.map((row) => {
      const values = [
        row.company, money(row.total), integer(row.operations), money(row.overdue), money(row.noDevenga),
        money(row.noDevenga6090), money(row.noDevengaOver90), row.chargedOff ? money(row.chargedOff) : '—'
      ];
      if (payload.scenario === 'PROYECTADO') values.push(integer(row.reclassified));
      values.push(row.status || 'OK');
      return values;
    });

    doc.setTextColor(...NAVY); doc.setFontSize(12); doc.setFont('helvetica', 'bold'); doc.text('Matriz ejecutiva por empresa', 12, 39);
    doc.setTextColor(...SLATE); doc.setFontSize(8); doc.setFont('helvetica', 'normal');
    doc.text('Resultados oficiales del mismo motor que genera el Excel estándar.', 12, 45);

    if (typeof doc.autoTable === 'function') {
      doc.autoTable({
        startY: 51,
        head,
        body,
        theme: 'grid',
        margin: { left: 12, right: 12, bottom: 18 },
        styles: { font: 'helvetica', fontSize: 7.5, cellPadding: 3.2, lineColor: LINE, lineWidth: 0.15, textColor: NAVY },
        headStyles: { fillColor: NAVY, textColor: WHITE, fontStyle: 'bold', halign: 'center' },
        alternateRowStyles: { fillColor: LIGHT },
        columnStyles: { 0: { fontStyle: 'bold', halign: 'left' } },
        didParseCell(data) { if (data.section === 'body' && data.column.index > 0 && data.column.index < head[0].length - 1) data.cell.styles.halign = 'right'; }
      });
    } else {
      let y = 54;
      body.forEach((row) => { doc.text(row.join(' | '), 12, y, { maxWidth: 270 }); y += 7; });
    }

    const endY = doc.lastAutoTable ? doc.lastAutoTable.finalY : 95;
    roundedRect(doc, 12, endY + 8, 273, 30, LIGHT, 3);
    doc.setTextColor(...NAVY); doc.setFontSize(9); doc.setFont('helvetica', 'bold'); doc.text('Notas de presentación', 18, endY + 16);
    doc.setTextColor(...SLATE); doc.setFontSize(7.5); doc.setFont('helvetica', 'normal');
    const notes = [
      '• Las cifras monetarias del PDF están expresadas en miles y redondeadas sin decimales.',
      '• El Excel estándar conserva precisión contable y trazabilidad mediante hojas ORIGEN.',
      '• Esta exportación no recalcula ni modifica la lógica de negocio.'
    ];
    notes.forEach((note, index) => doc.text(note, 18, endY + 22 + index * 5));
  }

  async function generate(payload) {
    if (!available()) throw new Error('El motor PDF no está disponible. Verifica la conexión o actualiza la página.');
    const { jsPDF } = global.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true });
    const rows = payload.rows || [];
    const totals = aggregate(rows);

    doc.setProperties({
      title: `CTH - Comité de Cartera - ${payload.scenarioLabel}`,
      subject: 'Reporte Enterprise de Comité de Cartera',
      author: 'Lizbeth Sanipatín',
      creator: 'CTH Comité de Cartera'
    });

    doc.setFillColor(237, 242, 246); doc.rect(0, 0, 297, 210, 'F');
    drawHeader(doc, payload);
    drawKpis(doc, payload, totals);
    drawBarChart(doc, rows, 68);
    drawRiskPanel(doc, payload, rows, totals, 68);
    drawExecutiveReading(doc, payload, totals, 148);

    doc.addPage('a4', 'landscape');
    doc.setFillColor(250, 251, 252); doc.rect(0, 0, 297, 210, 'F');
    drawHeader(doc, payload);
    drawTable(doc, payload, rows);

    const totalPages = doc.getNumberOfPages();
    for (let page = 1; page <= totalPages; page += 1) {
      doc.setPage(page);
      drawFooter(doc, payload, page, totalPages);
    }

    const blob = doc.output('blob');
    const signature = new Uint8Array(await blob.slice(0, 5).arrayBuffer());
    if (String.fromCharCode(...signature) !== '%PDF-') throw new Error('El archivo PDF generado no superó la validación de firma.');
    return blob;
  }

  global.PDFReport = Object.freeze({ available, generate });
})(typeof window !== 'undefined' ? window : globalThis);
