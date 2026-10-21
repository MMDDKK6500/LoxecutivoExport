require('dotenv').config();

const axios = require('axios');
const cors = require('cors');
const ExcelJS = require('exceljs');
const express = require('express');
const PDFDocument = require('pdfkit');
const { aggregateCompanyReport, createCompanyWorkbook, createCompanyPdf } = require('./companyReport');

const app = express();
const port = Number(process.env.PORT || 8083);
const coreUrl = (process.env.CORE_URL || 'http://127.0.0.1:8081').replace(/\/$/, '');
const requestTimeout = Number(process.env.REQUEST_TIMEOUT_MS || 15000);
const allowedOrigins = (process.env.CORS_ORIGINS || 'https://loxecutivo.mmddkk.dev.br,http://localhost:5173')
  .split(',')
  .map((origin) => origin.trim())
  .filter(Boolean);

const core = axios.create({
  baseURL: coreUrl,
  timeout: requestTimeout,
  headers: { Accept: 'application/json' },
});

const COLORS = {
  navy: '1F4E78',
  blue: 'D9EAF7',
  gray: '64748B',
  text: '1F2937',
  border: 'D7E1EA',
  white: 'FFFFFF',
};

const PDF_COLORS = {
  navy: '#1F4E78',
  blue: '#D9EAF7',
  lightBlue: '#EEF5FB',
  gray: '#64748B',
  text: '#1F2937',
  border: '#D7E1EA',
  white: '#FFFFFF',
};

app.use(cors({
  credentials: true,
  methods: ['GET', 'OPTIONS'],
  allowedHeaders: ['Authorization', 'Content-Type', 'Accept'],
  origin(origin, callback) {
    if (!origin || allowedOrigins.includes(origin)) return callback(null, true);
    return callback(new Error('Origem não permitida.'));
  },
}));
app.use(express.json({ limit: '1mb' }));

class ServiceError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}

function authorizationHeader(request) {
  const value = request.get('authorization');
  if (!value || !/^Bearer\s+\S+$/i.test(value)) {
    throw new ServiceError(401, 'Acesso negado: token Bearer ausente.');
  }
  return value;
}

async function coreGet(path, authorization) {
  try {
    const response = await core.get(path, {
      headers: { Authorization: authorization },
    });
    return response.data;
  } catch (error) {
    if (axios.isAxiosError(error)) {
      if (error.response) {
        const message = error.response.data?.message || error.response.data || 'O Core não conseguiu responder.';
        throw new ServiceError(error.response.status, String(message));
      }
      throw new ServiceError(503, 'O serviço operacional está indisponível.');
    }
    throw error;
  }
}

function requiredEventId(value) {
  if (!/^\d+$/.test(String(value)) || Number(value) < 1) {
    throw new ServiceError(400, 'O id do evento deve ser um número inteiro positivo.');
  }
  return Number(value);
}

function optionalDateFilter(value, name) {
  if (value === undefined || value === '') return '';
  const text = String(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    throw new ServiceError(400, `A data ${name} deve estar no formato AAAA-MM-DD.`);
  }
  const date = new Date(`${text}T00:00:00.000Z`);
  if (Number.isNaN(date.getTime()) || date.toISOString().slice(0, 10) !== text) {
    throw new ServiceError(400, `A data ${name} é inválida.`);
  }
  return text;
}

function fullName(driver) {
  if (!driver) return 'Não informado';
  return [driver.nome, driver.sobrenome].filter(Boolean).join(' ').trim() || 'Não informado';
}

function tripStatusLabel(status) {
  const labels = { AGENDADA: 'Agendada', REALIZADA: 'Realizada', CANCELADA: 'Cancelada' };
  return labels[status] || labels.AGENDADA;
}

function tripStatusCounts(trips) {
  return trips.reduce((counts, trip) => {
    const status = ['AGENDADA', 'REALIZADA', 'CANCELADA'].includes(trip.status) ? trip.status : 'AGENDADA';
    counts[status] += 1;
    return counts;
  }, { AGENDADA: 0, REALIZADA: 0, CANCELADA: 0 });
}

function addressLabel(address) {
  if (!address) return 'Não informado';
  return [address.rua, address.numero, address.bairro]
    .filter((value) => value !== undefined && value !== null && value !== '')
    .join(', ')
    .replace(/, ([^,]+)$/, ' - $1');
}

function addressKey(address) {
  if (!address) return 'unknown';
  return [address.id, address.rua, address.numero, address.bairro, address.cidade, address.uf]
    .map((value) => String(value ?? '').trim().toLowerCase())
    .join('|');
}

function dateParts(value) {
  if (!value) return { date: null, time: null, sort: Number.POSITIVE_INFINITY };
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (!match) return { date: null, time: null, sort: Number.POSITIVE_INFINITY };

  const [, year, month, day, hour = '00', minute = '00', second = '00'] = match;
  const date = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
  const time = (Number(hour) * 3600 + Number(minute) * 60 + Number(second)) / 86400;
  return { date, time, sort: Date.UTC(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second)) };
}

function compareTrips(left, right) {
  const driverOrder = fullName(left.idMotorista).localeCompare(fullName(right.idMotorista), 'pt-BR', { sensitivity: 'base' });
  if (driverOrder !== 0) return driverOrder;
  return dateParts(left.dataHora).sort - dateParts(right.dataHora).sort;
}

function vehicleDetails(vehicle, plate) {
  return vehicle || { idPlaca: plate, marca: '', modelo: '', ano: '', cor: '' };
}

function styleTitle(sheet, title, endColumn) {
  sheet.getCell('A2').value = title;
  sheet.mergeCells(`A2:${endColumn}2`);
  sheet.getCell('A2').font = { name: 'Arial', size: 16, bold: true, color: COLORS.navy };
  sheet.getCell('A2').alignment = { vertical: 'middle' };
  sheet.getRow(2).height = 24;
}

function styleTable(sheet, headerRow, lastRow, widths) {
  const header = sheet.getRow(headerRow);
  header.height = 22;
  header.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
    cell.font = { name: 'Arial', size: 10, bold: true, color: COLORS.white };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  for (let rowNumber = headerRow + 1; rowNumber <= lastRow; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    row.height = 20;
    row.eachCell((cell) => {
      cell.font = { name: 'Arial', size: 10, color: COLORS.text };
      cell.alignment = { vertical: 'middle', wrapText: true };
      cell.border = { bottom: { style: 'thin', color: COLORS.border } };
      if ((rowNumber - headerRow) % 2 === 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EEF5FB' } };
      }
    });
  }

  Object.entries(widths).forEach(([column, width]) => {
    sheet.getColumn(column).width = width;
  });
  sheet.autoFilter = { from: `A${headerRow}`, to: `${Object.keys(widths).at(-1)}${lastRow}` };
  sheet.views = [{ state: 'frozen', ySplit: headerRow }];
}

function addTable(sheet, name, headerRow, lastRow, lastColumn) {
  const columns = sheet.getRow(headerRow).values
    .slice(1)
    .map((value, index) => ({ name: String(value || `Coluna ${index + 1}`) }));
  const rows = Array.from({ length: Math.max(0, lastRow - headerRow) }, (_value, index) =>
    sheet.getRow(headerRow + index + 1).values.slice(1, columns.length + 1));
  sheet.addTable({
    name,
    ref: `A${headerRow}:${lastColumn}${lastRow}`,
    headerRow: true,
    columns,
    rows,
    style: { theme: 'TableStyleMedium2', showRowStripes: true },
  });
}

function addDateTimeFormats(sheet, dateColumn, timeColumn, startRow, endRow) {
  for (let row = startRow; row <= endRow; row += 1) {
    sheet.getCell(`${dateColumn}${row}`).numFmt = 'dd/mm/yyyy';
    sheet.getCell(`${timeColumn}${row}`).numFmt = 'hh:mm';
  }
}

function formatDate(value) {
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return match ? `${match[3]}/${match[2]}/${match[1]}` : 'Não definido';
}

function formatTime(value) {
  const match = String(value || '').match(/T(\d{2}):(\d{2})/);
  return match ? `${match[1]}:${match[2]}` : 'Não definido';
}

function reportData(event, trips, passengersByTrip, vehicles) {
  const sortedTrips = [...trips].sort(compareTrips);
  const passengerCollections = new Map();
  sortedTrips.forEach((trip, index) => {
    passengerCollections.set(trip.id, {
      id: `P-${String(index + 1).padStart(3, '0')}`,
      count: Array.isArray(passengersByTrip.get(trip.id)) ? passengersByTrip.get(trip.id).length : 0,
    });
  });
  const vehicleMap = new Map((Array.isArray(vehicles) ? vehicles : []).map((vehicle) => [vehicle.idPlaca, vehicle]));
  return { event, sortedTrips, passengerCollections, vehicleMap };
}

function pdfText(value) {
  return String(value ?? '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}

function createPdfBuffer(event, trips, passengersByTrip, vehicles) {
  const { sortedTrips, passengerCollections, vehicleMap } = reportData(event, trips, passengersByTrip, vehicles);
  const totalPassengers = sortedTrips
    .filter((trip) => trip.status === 'REALIZADA')
    .reduce((sum, trip) => sum + (passengerCollections.get(trip.id)?.count || 0), 0);
  const statusCounts = tripStatusCounts(sortedTrips);

  return new Promise((resolve, reject) => {
    const document = new PDFDocument({ size: 'A4', margin: 42, bufferPages: true, compress: true });
    const chunks = [];
    document.on('data', (chunk) => chunks.push(chunk));
    document.on('error', reject);
    document.on('end', () => resolve(Buffer.concat(chunks)));

    const pageWidth = 595.28;
    const pageHeight = 841.89;
    const margin = 42;
    const contentWidth = pageWidth - margin * 2;
    const bottom = pageHeight - margin - 26;
    let cursor = margin;

    const drawPageHeader = (continuation = false) => {
      document.font('Helvetica-Bold').fontSize(9).fillColor(PDF_COLORS.navy)
        .text(continuation ? `Relatório do evento - ${pdfText(event.nome)}` : 'Loxecutivo', margin, margin - 2);
      document.moveTo(margin, margin + 15).lineTo(pageWidth - margin, margin + 15)
        .strokeColor(PDF_COLORS.border).lineWidth(0.7).stroke();
      cursor = margin + 28;
    };

    const newPage = () => {
      document.addPage();
      drawPageHeader(true);
    };

    const ensureSpace = (height) => {
      if (cursor + height > bottom) newPage();
    };

    const drawSectionTitle = (title) => {
      ensureSpace(42);
      document.font('Helvetica-Bold').fontSize(15).fillColor(PDF_COLORS.navy).text(pdfText(title), margin, cursor);
      cursor += 25;
    };

    const drawTable = (title, columns, rows) => {
      drawSectionTitle(title);
      const rowHeight = 23;
      const drawHeader = () => {
        document.rect(margin, cursor, contentWidth, rowHeight).fill(PDF_COLORS.navy);
        let x = margin;
        columns.forEach((column) => {
          document.font('Helvetica-Bold').fontSize(7.6).fillColor(PDF_COLORS.white)
            .text(pdfText(column.label), x + 4, cursor + 7, { width: column.width - 8, align: 'left', lineBreak: false });
          x += column.width;
        });
        cursor += rowHeight;
      };
      drawHeader();

      rows.forEach((row, rowIndex) => {
        const values = columns.map((column) => pdfText(typeof column.value === 'function' ? column.value(row) : row[column.key]));
        const heights = columns.map((column, index) => document.heightOfString(values[index], { width: column.width - 8, font: 'Helvetica', size: 7.5 }));
        const height = Math.max(rowHeight, Math.min(42, Math.max(...heights) + 10));
        if (cursor + height > bottom) {
          newPage();
          drawSectionTitle(`${title} (continuação)`);
          drawHeader();
        }
        if (rowIndex % 2 === 0) document.rect(margin, cursor, contentWidth, height).fill(PDF_COLORS.lightBlue);
        document.rect(margin, cursor, contentWidth, height).strokeColor(PDF_COLORS.border).lineWidth(0.35).stroke();
        let x = margin;
        columns.forEach((column, index) => {
          document.font('Helvetica').fontSize(7.5).fillColor(PDF_COLORS.text)
            .text(values[index], x + 4, cursor + 6, { width: column.width - 8, height: height - 8, ellipsis: true });
          x += column.width;
        });
        cursor += height;
      });
      cursor += 18;
    };

    drawPageHeader();
    document.font('Helvetica-Bold').fontSize(24).fillColor(PDF_COLORS.navy).text('Relatório do evento', margin, cursor);
    cursor += 31;
    document.font('Helvetica-Oblique').fontSize(10).fillColor(PDF_COLORS.gray).text('Exportação gerada pelo Loxecutivo', margin, cursor);
    cursor += 32;

    const metadata = [
      ['Evento', event.nome || 'Não informado'],
      ['Local', addressLabel(event.idEndereco)],
      ['Cidade/UF', [event.idEndereco?.cidade, event.idEndereco?.uf].filter(Boolean).join('/') || 'Não informado'],
      ['Total de viagens', sortedTrips.length],
      ['Viagens agendadas', statusCounts.AGENDADA],
      ['Viagens realizadas', statusCounts.REALIZADA],
      ['Viagens canceladas', statusCounts.CANCELADA],
      ['Passageiros em viagens realizadas', totalPassengers],
      ['Gerado em', new Date().toLocaleString('pt-BR')],
    ];
    metadata.forEach(([label, value]) => {
      document.rect(margin, cursor, 145, 22).fill(PDF_COLORS.blue);
      document.font('Helvetica-Bold').fontSize(9).fillColor(PDF_COLORS.navy).text(pdfText(label), margin + 7, cursor + 7);
      document.font('Helvetica').fontSize(9).fillColor(PDF_COLORS.text).text(pdfText(value), margin + 155, cursor + 7, { width: contentWidth - 155 });
      cursor += 22;
    });
    cursor += 20;

    const tripRows = sortedTrips.map((trip) => {
      const collection = passengerCollections.get(trip.id);
      const vehicle = vehicleDetails(vehicleMap.get(trip.idVeiculo), trip.idVeiculo);
      return { motorista: fullName(trip.idMotorista), data: formatDate(trip.dataHora), horario: formatTime(trip.dataHora), veiculo: `${vehicle.marca || ''} ${vehicle.modelo || ''}\n${trip.idVeiculo || ''}`.trim(), passageiros: collection.id, quantidade: collection.count, status: tripStatusLabel(trip.status) };
    });
    drawTable('Resumo das viagens', [
      { key: 'motorista', label: 'Motorista', width: 85 },
      { key: 'data', label: 'Data', width: 55 },
      { key: 'horario', label: 'Horário', width: 42 },
      { key: 'veiculo', label: 'Veículo', width: 85 },
      { key: 'passageiros', label: 'Coleção', width: 85 },
      { key: 'quantidade', label: 'Qtd.', width: 38 },
      { key: 'status', label: 'Status', width: 58 },
    ], tripRows);

    newPage();
    const detailedTripRows = sortedTrips.map((trip) => {
      const collection = passengerCollections.get(trip.id);
      const vehicle = vehicleDetails(vehicleMap.get(trip.idVeiculo), trip.idVeiculo);
      return { motorista: fullName(trip.idMotorista), data: formatDate(trip.dataHora), horario: formatTime(trip.dataHora), origem: addressLabel(trip.localDeOrigem), destino: addressLabel(trip.localDeDestino), veiculo: `${vehicle.marca || ''} ${vehicle.modelo || ''} - ${trip.idVeiculo || ''}`, passageiros: `${collection.id} (${collection.count})`, status: tripStatusLabel(trip.status) };
    });
    drawTable('Viagens', [
      { key: 'motorista', label: 'Motorista', width: 65 },
      { key: 'data', label: 'Data', width: 40 },
      { key: 'horario', label: 'Hora', width: 35 },
      { key: 'origem', label: 'Origem', width: 70 },
      { key: 'destino', label: 'Destino', width: 70 },
      { key: 'veiculo', label: 'Veículo', width: 80 },
      { key: 'passageiros', label: 'Passageiros', width: 80 },
      { key: 'status', label: 'Status', width: 55 },
    ], detailedTripRows);

    newPage();
    const passengerRows = sortedTrips.map((trip) => {
      const collection = passengerCollections.get(trip.id);
      return { id: collection.id, data: formatDate(trip.dataHora), horario: formatTime(trip.dataHora), motorista: fullName(trip.idMotorista), veiculo: trip.idVeiculo || '', quantidade: collection.count };
    });
    drawTable('Coleções de passageiros', [
      { key: 'id', label: 'Coleção', width: 105 },
      { key: 'data', label: 'Data', width: 65 },
      { key: 'horario', label: 'Horário', width: 55 },
      { key: 'motorista', label: 'Motorista', width: 125 },
      { key: 'veiculo', label: 'Veículo', width: 100 },
      { key: 'quantidade', label: 'Quantidade', width: 60 },
    ], passengerRows);

    newPage();
    const drivers = [...new Map(sortedTrips.map((trip) => [trip.idMotorista?.id || fullName(trip.idMotorista), trip.idMotorista])).values()]
      .filter(Boolean)
      .sort((left, right) => fullName(left).localeCompare(fullName(right), 'pt-BR', { sensitivity: 'base' }));
    drawTable('Motoristas utilizados', [
      { key: 'id', label: 'ID', width: 55 },
      { key: 'nome', label: 'Nome', width: 115 },
      { key: 'sobrenome', label: 'Sobrenome', width: 145 },
      { key: 'rg', label: 'RG', width: 90 },
      { key: 'cpf', label: 'CPF', width: 105 },
    ], drivers);

    const usedVehicles = [...new Set(sortedTrips.map((trip) => trip.idVeiculo).filter(Boolean))]
      .map((plate) => vehicleDetails(vehicleMap.get(plate), plate));
    drawTable('Veículos utilizados', [
      { key: 'idPlaca', label: 'Placa', width: 100 },
      { key: 'marca', label: 'Marca', width: 125 },
      { key: 'modelo', label: 'Modelo', width: 135 },
      { key: 'ano', label: 'Ano', width: 65 },
      { key: 'cor', label: 'Cor', width: 85 },
    ], usedVehicles);

    const addressRows = [];
    const addAddress = (type, address, reference) => {
      if (!address) return;
      const existing = addressRows.find((row) => row.type === type && addressKey(row.address) === addressKey(address));
      if (existing) {
        if (reference && !existing.references.includes(reference)) existing.references.push(reference);
        return;
      }
      addressRows.push({ type, address, references: reference ? [reference] : [] });
    };
    addAddress('Local do evento', event.idEndereco, event.nome);
    sortedTrips.forEach((trip) => {
      const reference = `${fullName(trip.idMotorista)} (${formatTime(trip.dataHora)})`;
      addAddress('Origem', trip.localDeOrigem, reference);
      addAddress('Destino', trip.localDeDestino, reference);
    });
    drawTable('Endereços utilizados', [
      { key: 'tipo', label: 'Tipo', width: 75 },
      { key: 'endereco', label: 'Endereço', width: 180 },
      { key: 'cidade', label: 'Cidade/UF', width: 95 },
      { key: 'referencia', label: 'Referência', width: 161 },
    ], addressRows.map((row) => ({ tipo: row.type, endereco: `${addressLabel(row.address)}`, cidade: [row.address.cidade, row.address.uf].filter(Boolean).join('/'), referencia: row.references.join(', ') })));

    const pages = document.bufferedPageRange();
    for (let page = pages.start; page < pages.start + pages.count; page += 1) {
      document.switchToPage(page);
      document.font('Helvetica').fontSize(8).fillColor(PDF_COLORS.gray)
        .text(`Loxecutivo - Página ${page - pages.start + 1} de ${pages.count}`, margin, pageHeight - margin - 10, { width: contentWidth, align: 'right' });
    }
    document.end();
  });
}

function createWorkbook(event, trips, passengersByTrip, vehicles) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Loxecutivo Export';
  workbook.lastModifiedBy = 'Loxecutivo Export';
  workbook.created = new Date();
  workbook.modified = new Date();

  const sortedTrips = [...trips].sort(compareTrips);
  const passengerCollections = new Map();
  sortedTrips.forEach((trip, index) => {
    passengerCollections.set(trip.id, {
      id: `P-${String(index + 1).padStart(3, '0')}`,
      count: Array.isArray(passengersByTrip.get(trip.id)) ? passengersByTrip.get(trip.id).length : 0,
    });
  });

  const vehicleMap = new Map((Array.isArray(vehicles) ? vehicles : []).map((vehicle) => [vehicle.idPlaca, vehicle]));
  const eventAddress = event.idEndereco;
  const totalPassengers = sortedTrips
    .filter((trip) => trip.status === 'REALIZADA')
    .reduce((sum, trip) => sum + (passengerCollections.get(trip.id)?.count || 0), 0);
  const statusCounts = tripStatusCounts(sortedTrips);

  const summary = workbook.addWorksheet('Resumo');
  summary.views = [{ showGridLines: false }];
  styleTitle(summary, 'Relatório do evento', 'F');
  summary.getCell('A3').value = 'Exportação gerada pelo Loxecutivo';
  summary.getCell('A3').font = { name: 'Arial', size: 10, italic: true, color: COLORS.gray };
  summary.getCell('A5').value = 'Evento';
  summary.getCell('B5').value = event.nome;
  summary.getCell('A6').value = 'Local';
  summary.getCell('B6').value = addressLabel(eventAddress);
  summary.getCell('A7').value = 'Cidade/UF';
  summary.getCell('B7').value = [eventAddress?.cidade, eventAddress?.uf].filter(Boolean).join('/') || 'Não informado';
  summary.getCell('A8').value = 'Total de viagens';
  summary.getCell('B8').value = sortedTrips.length;
  summary.getCell('A9').value = 'Passageiros realizados';
  summary.getCell('B9').value = totalPassengers;
  summary.getCell('A10').value = 'Gerado em';
  summary.getCell('B10').value = new Date();
  summary.getCell('B10').numFmt = 'dd/mm/yyyy hh:mm';
  summary.getCell('D5').value = 'Agendadas';
  summary.getCell('E5').value = statusCounts.AGENDADA;
  summary.getCell('D6').value = 'Realizadas';
  summary.getCell('E6').value = statusCounts.REALIZADA;
  summary.getCell('D7').value = 'Canceladas';
  summary.getCell('E7').value = statusCounts.CANCELADA;
  for (let row = 5; row <= 10; row += 1) {
    summary.getCell(`A${row}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
    summary.getCell(`A${row}`).font = { name: 'Arial', size: 10, bold: true, color: COLORS.navy };
    summary.getCell(`B${row}`).font = { name: 'Arial', size: 10, color: COLORS.text };
  }
  for (let row = 5; row <= 7; row += 1) {
    summary.getCell(`D${row}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
    summary.getCell(`D${row}`).font = { name: 'Arial', size: 10, bold: true, color: COLORS.navy };
    summary.getCell(`E${row}`).font = { name: 'Arial', size: 10, color: COLORS.text };
  }
  summary.getRow(13).values = ['Motorista', 'Data', 'Horário', 'Veículo', 'Coleção de passageiros', 'Qtd.', 'Status'];
  sortedTrips.forEach((trip, index) => {
    const row = summary.getRow(14 + index);
    const parts = dateParts(trip.dataHora);
    const collection = passengerCollections.get(trip.id);
    row.values = [fullName(trip.idMotorista), parts.date, parts.time, trip.idVeiculo, collection.id, collection.count, tripStatusLabel(trip.status)];
  });
  const summaryLastRow = 13 + sortedTrips.length;
  addDateTimeFormats(summary, 'B', 'C', 14, summaryLastRow);
  styleTable(summary, 13, summaryLastRow, { A: 20, B: 13, C: 11, D: 14, E: 24, F: 10, G: 14 });
  addTable(summary, 'ResumoTable', 13, summaryLastRow, 'G');

  const tripSheet = workbook.addWorksheet('Viagens');
  tripSheet.views = [{ showGridLines: false }];
  styleTitle(tripSheet, 'Viagens do evento', 'N');
  tripSheet.getRow(4).values = ['Motorista', 'Data', 'Horário', 'Origem', 'Destino', 'CPF', 'RG', 'Marca', 'Modelo', 'Ano', 'Cor', 'Placa', 'Coleção de passageiros', 'Status'];
  sortedTrips.forEach((trip, index) => {
    const row = tripSheet.getRow(5 + index);
    const parts = dateParts(trip.dataHora);
    const collection = passengerCollections.get(trip.id);
    const vehicle = vehicleDetails(vehicleMap.get(trip.idVeiculo), trip.idVeiculo);
    row.values = [
      fullName(trip.idMotorista),
      parts.date,
      parts.time,
      addressLabel(trip.localDeOrigem),
      addressLabel(trip.localDeDestino),
      trip.idMotorista?.cpf || '',
      trip.idMotorista?.rg || '',
      vehicle.marca || '',
      vehicle.modelo || '',
      vehicle.ano || '',
      vehicle.cor || '',
      trip.idVeiculo || '',
      collection.id,
      tripStatusLabel(trip.status),
    ];
  });
  const tripLastRow = 4 + sortedTrips.length;
  addDateTimeFormats(tripSheet, 'B', 'C', 5, tripLastRow);
  styleTable(tripSheet, 4, tripLastRow, { A: 20, B: 12, C: 10, D: 31, E: 35, F: 14, G: 13, H: 14, I: 16, J: 10, K: 12, L: 12, M: 24, N: 14 });
  addTable(tripSheet, 'ViagensTable', 4, tripLastRow, 'N');

  const passengerSheet = workbook.addWorksheet('Passageiros');
  passengerSheet.views = [{ showGridLines: false }];
  styleTitle(passengerSheet, 'Passageiros do evento', 'F');
  passengerSheet.getRow(4).values = ['Coleção de passageiros', 'Data', 'Horário', 'Motorista', 'Veículo', 'Quantidade'];
  sortedTrips.forEach((trip, index) => {
    const row = passengerSheet.getRow(5 + index);
    const parts = dateParts(trip.dataHora);
    const collection = passengerCollections.get(trip.id);
    row.values = [collection.id, parts.date, parts.time, fullName(trip.idMotorista), trip.idVeiculo || '', collection.count];
  });
  const passengerLastRow = 4 + sortedTrips.length;
  addDateTimeFormats(passengerSheet, 'B', 'C', 5, passengerLastRow);
  styleTable(passengerSheet, 4, passengerLastRow, { A: 24, B: 12, C: 10, D: 20, E: 14, F: 13 });
  addTable(passengerSheet, 'PassageirosTable', 4, passengerLastRow, 'F');

  const driverSheet = workbook.addWorksheet('Motoristas');
  driverSheet.views = [{ showGridLines: false }];
  styleTitle(driverSheet, 'Motoristas utilizados', 'E');
  const drivers = [...new Map(sortedTrips.map((trip) => [trip.idMotorista?.id || fullName(trip.idMotorista), trip.idMotorista])).values()]
    .filter(Boolean)
    .sort((left, right) => fullName(left).localeCompare(fullName(right), 'pt-BR', { sensitivity: 'base' }));
  driverSheet.getRow(4).values = ['ID', 'Nome', 'Sobrenome', 'RG', 'CPF'];
  drivers.forEach((driver, index) => {
    driverSheet.getRow(5 + index).values = [driver.id || '', driver.nome || '', driver.sobrenome || '', driver.rg || '', driver.cpf || ''];
  });
  const driverLastRow = 4 + drivers.length;
  styleTable(driverSheet, 4, driverLastRow, { A: 10, B: 16, C: 18, D: 14, E: 15 });
  addTable(driverSheet, 'MotoristasTable', 4, driverLastRow, 'E');

  const vehicleSheet = workbook.addWorksheet('Veículos');
  vehicleSheet.views = [{ showGridLines: false }];
  styleTitle(vehicleSheet, 'Veículos utilizados', 'E');
  const usedVehicles = [...new Set(sortedTrips.map((trip) => trip.idVeiculo).filter(Boolean))]
    .map((plate) => vehicleDetails(vehicleMap.get(plate), plate));
  vehicleSheet.getRow(4).values = ['Placa', 'Marca', 'Modelo', 'Ano', 'Cor'];
  usedVehicles.forEach((vehicle, index) => {
    vehicleSheet.getRow(5 + index).values = [vehicle.idPlaca || '', vehicle.marca || '', vehicle.modelo || '', vehicle.ano || '', vehicle.cor || ''];
  });
  const vehicleLastRow = 4 + usedVehicles.length;
  styleTable(vehicleSheet, 4, vehicleLastRow, { A: 14, B: 16, C: 18, D: 10, E: 14 });
  addTable(vehicleSheet, 'VeiculosTable', 4, vehicleLastRow, 'E');

  const addressSheet = workbook.addWorksheet('Endereços');
  addressSheet.views = [{ showGridLines: false }];
  styleTitle(addressSheet, 'Endereços utilizados', 'G');
  addressSheet.getRow(4).values = ['Tipo', 'Rua', 'Número', 'Bairro', 'Cidade', 'UF', 'Referência'];
  const addressRows = [];
  const addAddress = (type, address, reference) => {
    if (!address) return;
    const existing = addressRows.find((row) => row.type === type && addressKey(row.address) === addressKey(address));
    if (existing) {
      if (reference && !existing.references.includes(reference)) existing.references.push(reference);
      return;
    }
    addressRows.push({ type, address, references: reference ? [reference] : [] });
  };
  addAddress('Local do evento', eventAddress, event.nome);
  sortedTrips.forEach((trip) => {
    const parts = dateParts(trip.dataHora);
    const reference = `${fullName(trip.idMotorista)} (${parts.time === null ? 'sem horário' : new Date(parts.time * 86400000).toISOString().slice(11, 16)})`;
    addAddress('Origem', trip.localDeOrigem, reference);
    addAddress('Destino', trip.localDeDestino, reference);
  });
  addressRows.forEach((row, index) => {
    const address = row.address;
    addressSheet.getRow(5 + index).values = [row.type, address.rua || '', address.numero || '', address.bairro || '', address.cidade || '', address.uf || '', row.references.join(', ')];
  });
  const addressLastRow = 4 + addressRows.length;
  styleTable(addressSheet, 4, addressLastRow, { A: 18, B: 24, C: 10, D: 18, E: 18, F: 8, G: 34 });
  addTable(addressSheet, 'EnderecosTable', 4, addressLastRow, 'G');

  return workbook;
}

app.get('/health', (_request, response) => {
  response.json({ status: 'ok', service: 'loxecutivo-export' });
});

async function loadEventData(eventId, authorization) {
  const event = await coreGet(`/eventos/${eventId}`, authorization);
  const trips = await coreGet(`/viagens/evento/${eventId}`, authorization);
  const vehicles = await coreGet('/veiculos', authorization);
  const passengersByTrip = new Map();

  await Promise.all((Array.isArray(trips) ? trips : []).map(async (trip) => {
    passengersByTrip.set(trip.id, await coreGet(`/passageiros/viagem/${trip.id}`, authorization));
  }));

  return { event, trips: Array.isArray(trips) ? trips : [], passengersByTrip, vehicles };
}

async function loadCompanyReport(authorization, filters) {
  const [events, trips, drivers, vehicles, addresses, passengers] = await Promise.all([
    coreGet('/eventos', authorization),
    coreGet('/viagens', authorization),
    coreGet('/motoristas', authorization),
    coreGet('/veiculos', authorization),
    coreGet('/enderecos', authorization),
    coreGet('/passageiros', authorization),
  ]);
  return aggregateCompanyReport({ events, trips, drivers, vehicles, addresses, passengers }, filters);
}

function requestedCompanyReport(request) {
  const from = optionalDateFilter(request.query.de, 'inicial');
  const to = optionalDateFilter(request.query.ate, 'final');
  if (from && to && from > to) {
    throw new ServiceError(400, 'A data inicial não pode ser posterior à data final.');
  }
  return { from, to };
}

function companyReportBaseName(report) {
  const from = report.period.from || 'inicio';
  const to = report.period.to || 'hoje';
  return `relatorio-geral-${from}-${to}`;
}

app.get('/relatorio-geral', async (request, response, next) => {
  try {
    const authorization = authorizationHeader(request);
    const report = await loadCompanyReport(authorization, requestedCompanyReport(request));
    response.json(report);
  } catch (error) {
    next(error);
  }
});

app.get('/relatorio-geral/excel', async (request, response, next) => {
  try {
    const authorization = authorizationHeader(request);
    const report = await loadCompanyReport(authorization, requestedCompanyReport(request));
    const buffer = await createCompanyWorkbook(report).xlsx.writeBuffer();
    response.status(200).set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${companyReportBaseName(report)}.xlsx"`,
      'Content-Length': buffer.length,
    }).send(buffer);
  } catch (error) {
    next(error);
  }
});

app.get('/relatorio-geral/pdf', async (request, response, next) => {
  try {
    const authorization = authorizationHeader(request);
    const report = await loadCompanyReport(authorization, requestedCompanyReport(request));
    const buffer = await createCompanyPdf(report);
    response.status(200).set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="${companyReportBaseName(report)}.pdf"`,
      'Content-Length': buffer.length,
    }).send(buffer);
  } catch (error) {
    next(error);
  }
});

app.get('/eventos/:eventId/excel', async (request, response, next) => {
  try {
    const eventId = requiredEventId(request.params.eventId);
    const authorization = authorizationHeader(request);
    const { event, trips, passengersByTrip, vehicles } = await loadEventData(eventId, authorization);

    const workbook = createWorkbook(event, trips, passengersByTrip, vehicles);
    const buffer = await workbook.xlsx.writeBuffer();
    const baseName = String(event.nome || `evento-${eventId}`)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-zA-Z0-9]+/g, '-')
      .replace(/^-|-$/g, '')
      .toLowerCase() || `evento-${eventId}`;

    response.status(200);
    response.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${baseName}.xlsx"`,
      'Content-Length': buffer.length,
    });
    response.send(buffer);
  } catch (error) {
    next(error);
  }
});

app.get('/eventos/:eventId/pdf', async (request, response, next) => {
  try {
    const eventId = requiredEventId(request.params.eventId);
    const authorization = authorizationHeader(request);
    const { event, trips, passengersByTrip, vehicles } = await loadEventData(eventId, authorization);
    const buffer = await createPdfBuffer(event, trips, passengersByTrip, vehicles);
    const baseName = String(event.nome || `evento-${eventId}`)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-zA-Z0-9]+/g, '-')
      .replace(/^-|-$/g, '')
      .toLowerCase() || `evento-${eventId}`;

    response.status(200);
    response.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="${baseName}.pdf"`,
      'Content-Length': buffer.length,
    });
    response.send(buffer);
  } catch (error) {
    next(error);
  }
});

app.use((error, _request, response, _next) => {
  const status = Number(error.status) || 500;
  if (status >= 500) console.error('Export error:', error);
  response.status(status).json({
    status,
    message: status === 500 ? 'Não foi possível gerar o arquivo de exportação.' : error.message,
  });
});

if (require.main === module) {
  app.listen(port, () => {
    console.log(`Export service running on port ${port}`);
  });
}

module.exports = { app, createWorkbook, createPdfBuffer, loadCompanyReport, requestedCompanyReport };
