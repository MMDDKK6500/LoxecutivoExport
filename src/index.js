require('dotenv').config();

const axios = require('axios');
const cors = require('cors');
const ExcelJS = require('exceljs');
const express = require('express');

const app = express();
const port = Number(process.env.PORT || 8083);
const coreUrl = (process.env.CORE_URL || 'http://127.0.0.1:8081').replace(/\/$/, '');
const requestTimeout = Number(process.env.REQUEST_TIMEOUT_MS || 15000);

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

app.use(cors());
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

function fullName(driver) {
  if (!driver) return 'Não informado';
  return [driver.nome, driver.sobrenome].filter(Boolean).join(' ').trim() || 'Não informado';
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
  const totalPassengers = [...passengerCollections.values()].reduce((sum, item) => sum + item.count, 0);

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
  summary.getCell('A9').value = 'Total de passageiros';
  summary.getCell('B9').value = totalPassengers;
  summary.getCell('A10').value = 'Gerado em';
  summary.getCell('B10').value = new Date();
  summary.getCell('B10').numFmt = 'dd/mm/yyyy hh:mm';
  for (let row = 5; row <= 10; row += 1) {
    summary.getCell(`A${row}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
    summary.getCell(`A${row}`).font = { name: 'Arial', size: 10, bold: true, color: COLORS.navy };
    summary.getCell(`B${row}`).font = { name: 'Arial', size: 10, color: COLORS.text };
  }
  summary.getRow(13).values = ['Motorista', 'Data', 'Horário', 'Veículo', 'Coleção de passageiros', 'Qtd.'];
  sortedTrips.forEach((trip, index) => {
    const row = summary.getRow(14 + index);
    const parts = dateParts(trip.dataHora);
    const collection = passengerCollections.get(trip.id);
    row.values = [fullName(trip.idMotorista), parts.date, parts.time, trip.idVeiculo, collection.id, collection.count];
  });
  const summaryLastRow = 13 + sortedTrips.length;
  addDateTimeFormats(summary, 'B', 'C', 14, summaryLastRow);
  styleTable(summary, 13, summaryLastRow, { A: 20, B: 13, C: 11, D: 14, E: 24, F: 10 });
  addTable(summary, 'ResumoTable', 13, summaryLastRow, 'F');

  const tripSheet = workbook.addWorksheet('Viagens');
  tripSheet.views = [{ showGridLines: false }];
  styleTitle(tripSheet, 'Viagens do evento', 'M');
  tripSheet.getRow(4).values = ['Motorista', 'Data', 'Horário', 'Origem', 'Destino', 'CPF', 'RG', 'Marca', 'Modelo', 'Ano', 'Cor', 'Placa', 'Coleção de passageiros'];
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
    ];
  });
  const tripLastRow = 4 + sortedTrips.length;
  addDateTimeFormats(tripSheet, 'B', 'C', 5, tripLastRow);
  styleTable(tripSheet, 4, tripLastRow, { A: 20, B: 12, C: 10, D: 31, E: 35, F: 14, G: 13, H: 14, I: 16, J: 10, K: 12, L: 12, M: 24 });
  addTable(tripSheet, 'ViagensTable', 4, tripLastRow, 'M');

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

app.get('/eventos/:eventId/excel', async (request, response, next) => {
  try {
    const eventId = requiredEventId(request.params.eventId);
    const authorization = authorizationHeader(request);
    const event = await coreGet(`/eventos/${eventId}`, authorization);
    const trips = await coreGet(`/viagens/evento/${eventId}`, authorization);
    const vehicles = await coreGet('/veiculos', authorization);
    const passengersByTrip = new Map();

    await Promise.all((Array.isArray(trips) ? trips : []).map(async (trip) => {
      passengersByTrip.set(trip.id, await coreGet(`/passageiros/viagem/${trip.id}`, authorization));
    }));

    const workbook = createWorkbook(event, Array.isArray(trips) ? trips : [], passengersByTrip, vehicles);
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

app.use((error, _request, response, _next) => {
  const status = Number(error.status) || 500;
  if (status >= 500) console.error('Export error:', error);
  response.status(status).json({
    status,
    message: status === 500 ? 'Não foi possível gerar o arquivo Excel.' : error.message,
  });
});

if (require.main === module) {
  app.listen(port, () => {
    console.log(`Export service running on port ${port}`);
  });
}

module.exports = { app, createWorkbook };
