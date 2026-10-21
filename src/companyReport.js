const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');

const STATUS_LABELS = { AGENDADA: 'Agendada', REALIZADA: 'Realizada', CANCELADA: 'Cancelada' };
const STATUS_KEYS = Object.keys(STATUS_LABELS);
const WEEKDAYS = ['Domingo', 'Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado'];
const COLORS = { navy: '1F4E78', blue: 'D9EAF7', gray: '64748B', text: '1F2937', border: 'D7E1EA', white: 'FFFFFF', green: '16824A' };

function normalizeStatus(status) {
  return STATUS_KEYS.includes(status) ? status : 'AGENDADA';
}

function fullName(person) {
  return [person?.nome, person?.sobrenome].filter(Boolean).join(' ').trim() || 'Não informado';
}

function addressLabel(address) {
  if (!address) return 'Não informado';
  return [address.rua, address.numero, address.bairro, [address.cidade, address.uf].filter(Boolean).join('/')]
    .filter((value) => value !== undefined && value !== null && value !== '')
    .join(', ');
}

function tripDate(trip) {
  const match = String(trip?.dataHora || '').match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?/);
  if (!match) return null;
  const [, year, month, day, hour = '00', minute = '00'] = match;
  return {
    date: `${year}-${month}-${day}`,
    dateValue: new Date(Date.UTC(Number(year), Number(month) - 1, Number(day))),
    time: `${hour}:${minute}`,
    hour: Number(hour),
    weekday: new Date(Date.UTC(Number(year), Number(month) - 1, Number(day))).getUTCDay(),
  };
}

function idKey(value, fallback) {
  return value === undefined || value === null || value === '' ? fallback : String(value);
}

function aggregateCompanyReport(source, filters = {}) {
  const start = filters.from || '';
  const end = filters.to || '';
  const events = Array.isArray(source.events) ? source.events : [];
  const allTrips = Array.isArray(source.trips) ? source.trips : [];
  const registeredDrivers = Array.isArray(source.drivers) ? source.drivers : [];
  const registeredVehicles = Array.isArray(source.vehicles) ? source.vehicles : [];
  const passengerRows = Array.isArray(source.passengers) ? source.passengers : [];

  const trips = allTrips.filter((trip) => {
    const date = tripDate(trip)?.date;
    if (start && (!date || date < start)) return false;
    if (end && (!date || date > end)) return false;
    return true;
  });

  const passengersByTrip = new Map();
  passengerRows.forEach((passenger) => {
    const tripId = passenger?.idViagem?.id;
    if (tripId === undefined || tripId === null) return;
    const key = String(tripId);
    passengersByTrip.set(key, (passengersByTrip.get(key) || 0) + 1);
  });

  const statusCounts = { AGENDADA: 0, REALIZADA: 0, CANCELADA: 0 };
  const driverUsage = new Map();
  const vehicleUsage = new Map();
  const eventUsage = new Map();
  const routeUsage = new Map();
  const weekdayUsage = WEEKDAYS.map((nome) => ({ nome, viagens: 0, realizadas: 0, agendadas: 0, canceladas: 0 }));
  const hourlyUsage = Array.from({ length: 24 }, (_, hour) => ({ hora: `${String(hour).padStart(2, '0')}:00`, viagens: 0, realizadas: 0, agendadas: 0, canceladas: 0 }));
  const eventMap = new Map(events.map((event) => [String(event.id), event]));
  const driverMap = new Map(registeredDrivers.map((driver) => [idKey(driver.id, fullName(driver)), driver]));
  const vehicleMap = new Map(registeredVehicles.map((vehicle) => [idKey(vehicle.idPlaca), vehicle]));
  const actualPassengers = { total: 0 };

  trips.forEach((trip) => {
    const status = normalizeStatus(trip.status);
    const statusField = { AGENDADA: 'agendadas', REALIZADA: 'realizadas', CANCELADA: 'canceladas' }[status];
    statusCounts[status] += 1;
    const passengers = passengersByTrip.get(String(trip.id)) || 0;
    if (status === 'REALIZADA') actualPassengers.total += passengers;

    const driver = trip.idMotorista || {};
    const driverKey = idKey(driver.id, fullName(driver));
    const driverName = fullName(driver);
    if (!driverMap.has(driverKey)) driverMap.set(driverKey, driver);
    const driverEntry = driverUsage.get(driverKey) || { id: driver.id ?? '', nome: driverName, realizadas: 0, agendadas: 0, canceladas: 0, passageiros: 0 };
    driverEntry[statusField] += 1;
    if (status === 'REALIZADA') driverEntry.passageiros += passengers;
    driverUsage.set(driverKey, driverEntry);

    const plate = trip.idVeiculo || 'Não informado';
    if (plate !== 'Não informado' && !vehicleMap.has(String(plate))) vehicleMap.set(String(plate), { idPlaca: plate });
    const vehicleEntry = vehicleUsage.get(String(plate)) || { placa: plate, marca: '', modelo: '', realizadas: 0, agendadas: 0, canceladas: 0, passageiros: 0 };
    const vehicle = vehicleMap.get(String(plate)) || {};
    vehicleEntry.marca = vehicle.marca || vehicleEntry.marca;
    vehicleEntry.modelo = vehicle.modelo || vehicleEntry.modelo;
    vehicleEntry[statusField] += 1;
    if (status === 'REALIZADA') vehicleEntry.passageiros += passengers;
    vehicleUsage.set(String(plate), vehicleEntry);

    const event = trip.idEvento || {};
    const eventKey = idKey(event.id, event.nome || 'Sem evento');
    if (event.id !== undefined && !eventMap.has(String(event.id))) eventMap.set(String(event.id), event);
    const eventEntry = eventUsage.get(eventKey) || { id: event.id ?? '', nome: event.nome || 'Sem evento', endereco: addressLabel(event.idEndereco), realizadas: 0, agendadas: 0, canceladas: 0, passageiros: 0 };
    eventEntry[statusField] += 1;
    if (status === 'REALIZADA') eventEntry.passageiros += passengers;
    eventUsage.set(eventKey, eventEntry);

    const date = tripDate(trip);
    if (date) {
      const dayEntry = weekdayUsage[date.weekday];
      dayEntry.viagens += 1;
      dayEntry[statusField] += 1;
      const hourEntry = hourlyUsage[date.hour];
      hourEntry.viagens += 1;
      hourEntry[statusField] += 1;
    }

    const origin = addressLabel(trip.localDeOrigem);
    const destination = addressLabel(trip.localDeDestino);
    const routeKey = `${origin}|||${destination}`;
    const routeEntry = routeUsage.get(routeKey) || { origem: origin, destino: destination, realizadas: 0, agendadas: 0, canceladas: 0 };
    routeEntry[statusField] += 1;
    routeUsage.set(routeKey, routeEntry);
  });

  const drivers = [...driverMap.entries()].map(([key, driver]) => {
    const usage = driverUsage.get(key) || { id: driver.id ?? '', nome: fullName(driver), realizadas: 0, agendadas: 0, canceladas: 0, passageiros: 0 };
    return usage;
  }).sort((a, b) => b.realizadas - a.realizadas || a.nome.localeCompare(b.nome, 'pt-BR', { sensitivity: 'base' }));

  const vehicles = [...vehicleMap.entries()].map(([plate, vehicle]) => {
    const usage = vehicleUsage.get(plate) || { placa: vehicle.idPlaca || plate, marca: vehicle.marca || '', modelo: vehicle.modelo || '', realizadas: 0, agendadas: 0, canceladas: 0, passageiros: 0 };
    return usage;
  }).sort((a, b) => b.realizadas - a.realizadas || a.placa.localeCompare(b.placa, 'pt-BR'));

  const eventRows = [...eventMap.entries()].map(([key, event]) => {
    const usage = eventUsage.get(key) || { realizadas: 0, agendadas: 0, canceladas: 0, passageiros: 0 };
    return { id: event.id ?? '', nome: event.nome || 'Sem nome', endereco: addressLabel(event.idEndereco), ...usage };
  }).sort((a, b) => (b.realizadas + b.agendadas + b.canceladas) - (a.realizadas + a.agendadas + a.canceladas) || a.nome.localeCompare(b.nome, 'pt-BR'));
  const attendedEvents = [...eventUsage.values()].filter((event) => event.realizadas > 0).length;
  const activeDrivers = drivers.filter((driver) => driver.realizadas > 0).length;
  const activeVehicles = vehicles.filter((vehicle) => vehicle.realizadas > 0).length;
  const completedTrips = statusCounts.REALIZADA;
  const peakHour = [...hourlyUsage].sort((a, b) => b.realizadas - a.realizadas || b.agendadas - a.agendadas)[0];
  const peakWeekday = [...weekdayUsage].sort((a, b) => b.realizadas - a.realizadas || b.agendadas - a.agendadas)[0];
  const routes = [...routeUsage.values()].sort((a, b) => b.realizadas - a.realizadas || b.agendadas - a.agendadas).slice(0, 15);

  return {
    period: { from: start || null, to: end || null },
    summary: {
      totalViagens: trips.length,
      agendadas: statusCounts.AGENDADA,
      realizadas: statusCounts.REALIZADA,
      canceladas: statusCounts.CANCELADA,
      passageirosTransportados: actualPassengers.total,
      mediaPassageirosPorViagemRealizada: completedTrips ? Number((actualPassengers.total / completedTrips).toFixed(1)) : 0,
      eventosAtendidos: attendedEvents,
      motoristasUtilizados: activeDrivers,
      veiculosUtilizados: activeVehicles,
      motoristasCadastrados: registeredDrivers.length,
      veiculosCadastrados: registeredVehicles.length,
      enderecosCadastrados: Array.isArray(source.addresses) ? source.addresses.length : 0,
      picoDia: peakWeekday?.realizadas ? peakWeekday.nome : peakWeekday?.agendadas ? `${peakWeekday.nome} (agendadas)` : 'Sem viagens',
      picoHorario: peakHour?.realizadas ? peakHour.hora : peakHour?.agendadas ? `${peakHour.hora} (agendadas)` : 'Sem viagens',
    },
    statusCounts,
    trips: [...trips].sort((a, b) => String(a.dataHora || '').localeCompare(String(b.dataHora || ''))),
    drivers,
    vehicles,
    events: eventRows,
    weekdays: weekdayUsage,
    hours: hourlyUsage,
    routes,
    passengerCounts: Object.fromEntries(passengersByTrip),
  };
}

function displayDate(value) {
  if (!value) return '';
  const [year, month, day] = value.split('-');
  return `${day}/${month}/${year}`;
}

function tableSheet(workbook, name, title, headers, rows, widths) {
  const sheet = workbook.addWorksheet(name, { views: [{ showGridLines: false, state: 'frozen', ySplit: 4 }] });
  const lastColumn = String.fromCharCode(64 + headers.length);
  sheet.mergeCells(`A1:${lastColumn}1`);
  sheet.getCell('A1').value = title;
  sheet.getCell('A1').font = { name: 'Arial', size: 16, bold: true, color: COLORS.navy };
  sheet.getCell('A1').alignment = { vertical: 'middle' };
  sheet.getRow(1).height = 26;
  sheet.getRow(3).values = headers;
  const header = sheet.getRow(3);
  header.height = 28;
  header.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.navy } };
    cell.font = { name: 'Arial', size: 10, bold: true, color: COLORS.white };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });
  rows.forEach((values, index) => {
    const row = sheet.getRow(index + 4);
    row.values = values;
    row.height = 24;
    row.eachCell((cell) => {
      cell.font = { name: 'Arial', size: 10, color: COLORS.text };
      cell.alignment = { vertical: 'middle', wrapText: true };
      cell.border = { bottom: { style: 'thin', color: COLORS.border } };
      if (index % 2 === 0) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EEF5FB' } };
    });
  });
  widths.forEach((width, index) => { sheet.getColumn(index + 1).width = width; });
  if (rows.length) sheet.autoFilter = { from: 'A3', to: `${lastColumn}${rows.length + 3}` };
  return sheet;
}

function createCompanyWorkbook(report) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Loxecutivo';
  workbook.created = new Date();
  const { summary } = report;
  const overview = workbook.addWorksheet('Resumo', { views: [{ showGridLines: false }] });
  overview.mergeCells('A1:E1');
  overview.getCell('A1').value = 'Relatório geral da empresa';
  overview.getCell('A1').font = { name: 'Arial', size: 18, bold: true, color: COLORS.navy };
  overview.getRow(1).height = 30;
  overview.getCell('A3').value = 'Período';
  overview.getCell('B3').value = report.period.from || report.period.to
    ? `${report.period.from ? displayDate(report.period.from) : 'Sem data inicial'} a ${report.period.to ? displayDate(report.period.to) : 'Sem data final'}`
    : 'Todo o período';
  const indicators = [
    ['Total de viagens', summary.totalViagens],
    ['Passageiros transportados', summary.passageirosTransportados],
    ['Média de passageiros por viagem realizada', summary.mediaPassageirosPorViagemRealizada],
    ['Eventos com viagens realizadas', summary.eventosAtendidos],
    ['Motoristas utilizados', summary.motoristasUtilizados],
    ['Veículos utilizados', summary.veiculosUtilizados],
  ];
  indicators.forEach(([label, value], index) => {
    const row = index + 5;
    overview.getCell(`A${row}`).value = label;
    overview.getCell(`B${row}`).value = value;
  });
  const states = [['Agendadas', summary.agendadas], ['Realizadas', summary.realizadas], ['Canceladas', summary.canceladas]];
  states.forEach(([label, value], index) => {
    const row = index + 5;
    overview.getCell(`D${row}`).value = label;
    overview.getCell(`E${row}`).value = value;
  });
  overview.getCell('D10').value = 'Dia de maior movimento';
  overview.getCell('E10').value = summary.picoDia;
  overview.getCell('D11').value = 'Horário de maior movimento';
  overview.getCell('E11').value = summary.picoHorario;
  for (const row of [...Array(6)].map((_, index) => index + 5)) {
    overview.getCell(`A${row}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
    overview.getCell(`A${row}`).font = { bold: true, color: COLORS.navy };
  }
  for (const row of [5, 6, 7, 10, 11]) {
    overview.getCell(`D${row}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.blue } };
    overview.getCell(`D${row}`).font = { bold: true, color: COLORS.navy };
  }
  overview.getColumn('A').width = 38;
  overview.getColumn('B').width = 22;
  overview.getColumn('C').width = 4;
  overview.getColumn('D').width = 31;
  overview.getColumn('E').width = 24;

  tableSheet(workbook, 'Viagens', 'Viagens do período',
    ['Data', 'Horário', 'Status', 'Evento', 'Motorista', 'Veículo', 'Origem', 'Destino', 'Passageiros'],
    report.trips.map((trip) => {
      const date = tripDate(trip);
      return [date ? displayDate(date.date) : 'Sem data', date?.time || '', STATUS_LABELS[normalizeStatus(trip.status)], trip.idEvento?.nome || 'Sem evento', fullName(trip.idMotorista), trip.idVeiculo || '', addressLabel(trip.localDeOrigem), addressLabel(trip.localDeDestino), report.passengerCounts[String(trip.id)] || 0];
    }), [13, 11, 14, 24, 22, 15, 32, 32, 13]);

  tableSheet(workbook, 'Motoristas', 'Utilização por motorista',
    ['Motorista', 'Viagens realizadas', 'Agendadas', 'Canceladas', 'Passageiros transportados'],
    report.drivers.map((item) => [item.nome, item.realizadas, item.agendadas, item.canceladas, item.passageiros]), [30, 19, 15, 15, 27]);
  tableSheet(workbook, 'Veículos', 'Utilização por veículo',
    ['Placa', 'Marca', 'Modelo', 'Viagens realizadas', 'Agendadas', 'Canceladas', 'Passageiros transportados'],
    report.vehicles.map((item) => [item.placa, item.marca, item.modelo, item.realizadas, item.agendadas, item.canceladas, item.passageiros]), [15, 20, 22, 19, 14, 14, 27]);
  tableSheet(workbook, 'Eventos', 'Viagens por evento',
    ['Evento', 'Local', 'Viagens realizadas', 'Agendadas', 'Canceladas', 'Passageiros transportados'],
    report.events.filter((item) => item.realizadas + item.agendadas + item.canceladas > 0)
      .map((item) => [item.nome, item.endereco, item.realizadas, item.agendadas, item.canceladas, item.passageiros]), [30, 38, 19, 14, 14, 27]);
  tableSheet(workbook, 'Dias da semana', 'Distribuição por dia da semana',
    ['Dia', 'Viagens', 'Realizadas', 'Agendadas', 'Canceladas'],
    report.weekdays.map((item) => [item.nome, item.viagens, item.realizadas, item.agendadas, item.canceladas]), [24, 14, 15, 15, 15]);
  tableSheet(workbook, 'Horários', 'Distribuição por horário de saída',
    ['Horário', 'Viagens', 'Realizadas', 'Agendadas', 'Canceladas'],
    report.hours.filter((item) => item.viagens > 0).map((item) => [item.hora, item.viagens, item.realizadas, item.agendadas, item.canceladas]), [18, 14, 15, 15, 15]);
  tableSheet(workbook, 'Rotas', 'Rotas mais utilizadas',
    ['Origem', 'Destino', 'Viagens realizadas', 'Agendadas', 'Canceladas'],
    report.routes.map((item) => [item.origem, item.destino, item.realizadas, item.agendadas, item.canceladas]), [38, 38, 19, 14, 14]);

  return workbook;
}

function createCompanyPdf(report) {
  const document = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 35, bufferPages: true, compress: true });
  const chunks = [];
  const result = new Promise((resolve, reject) => {
    document.on('data', (chunk) => chunks.push(chunk));
    document.on('error', reject);
    document.on('end', () => resolve(Buffer.concat(chunks)));
  });
  const pageWidth = 841.89;
  const pageHeight = 595.28;
  const margin = 35;
  const contentWidth = pageWidth - margin * 2;
  let y = margin;

  const header = (continuation = false) => {
    document.font('Helvetica-Bold').fontSize(9).fillColor('#1F4E78').text('Loxecutivo · Relatório geral da empresa', margin, margin - 5);
    document.moveTo(margin, margin + 12).lineTo(pageWidth - margin, margin + 12).strokeColor('#D7E1EA').stroke();
    y = margin + (continuation ? 27 : 34);
  };
  const ensure = (height) => {
    if (y + height > pageHeight - margin) {
      document.addPage({ size: 'A4', layout: 'landscape', margin });
      header(true);
    }
  };
  const sectionTitle = (title) => {
    ensure(28);
    document.font('Helvetica-Bold').fontSize(13).fillColor('#1F4E78').text(title, margin, y);
    y += 21;
  };
  const table = (title, columns, rows) => {
    sectionTitle(title);
    const rowHeight = 20;
    const drawHeader = () => {
      document.rect(margin, y, contentWidth, rowHeight).fill('#1F4E78');
      let x = margin;
      columns.forEach((column) => {
        document.font('Helvetica-Bold').fontSize(7.5).fillColor('#FFFFFF').text(column.title, x + 4, y + 6, { width: column.width - 8, lineBreak: false, ellipsis: true });
        x += column.width;
      });
      y += rowHeight;
    };
    drawHeader();
    rows.forEach((row, index) => {
      if (y + rowHeight > pageHeight - margin) {
        document.addPage({ size: 'A4', layout: 'landscape', margin });
        header(true);
        sectionTitle(`${title} (continuação)`);
        drawHeader();
      }
      if (index % 2 === 0) document.rect(margin, y, contentWidth, rowHeight).fill('#EEF5FB');
      let x = margin;
      columns.forEach((column) => {
        document.font('Helvetica').fontSize(7.3).fillColor('#1F2937').text(String(row[column.key] ?? ''), x + 4, y + 5, { width: column.width - 8, height: rowHeight - 4, ellipsis: true, lineBreak: false });
        x += column.width;
      });
      y += rowHeight;
    });
    y += 14;
  };

  header();
  document.font('Helvetica-Bold').fontSize(22).fillColor('#1F4E78').text('Relatório geral da empresa', margin, y);
  y += 30;
  const period = report.period.from || report.period.to
    ? `${report.period.from ? displayDate(report.period.from) : 'Sem data inicial'} a ${report.period.to ? displayDate(report.period.to) : 'Sem data final'}`
    : 'Todo o período';
  document.font('Helvetica').fontSize(10).fillColor('#64748B').text(`Período: ${period} · Gerado em ${new Date().toLocaleString('pt-BR')}`, margin, y);
  y += 26;

  const summary = report.summary;
  const kpis = [
    ['Viagens', summary.totalViagens], ['Realizadas', summary.realizadas], ['Agendadas', summary.agendadas],
    ['Canceladas', summary.canceladas], ['Passageiros transportados', summary.passageirosTransportados],
    ['Eventos atendidos', summary.eventosAtendidos], ['Motoristas utilizados', summary.motoristasUtilizados], ['Veículos utilizados', summary.veiculosUtilizados],
  ];
  const boxGap = 8;
  const boxWidth = (contentWidth - boxGap * 3) / 4;
  kpis.forEach(([label, value], index) => {
    const row = Math.floor(index / 4);
    const column = index % 4;
    const x = margin + column * (boxWidth + boxGap);
    const boxY = y + row * 46;
    document.roundedRect(x, boxY, boxWidth, 39, 5).fill('#EEF5FB');
    document.font('Helvetica-Bold').fontSize(14).fillColor('#1F4E78').text(String(value), x + 10, boxY + 5, { width: boxWidth - 20, lineBreak: false });
    document.font('Helvetica').fontSize(8).fillColor('#64748B').text(label, x + 10, boxY + 23, { width: boxWidth - 20, lineBreak: false });
  });
  y += 102;
  document.font('Helvetica').fontSize(9).fillColor('#64748B')
    .text(`Pico de movimento: ${summary.picoDia} · ${summary.picoHorario} | Média: ${summary.mediaPassageirosPorViagemRealizada} passageiros por viagem realizada`, margin, y);
  y += 24;

  table('Utilização por motorista', [
    { key: 'nome', title: 'Motorista', width: 300 }, { key: 'realizadas', title: 'Realizadas', width: 110 },
    { key: 'agendadas', title: 'Agendadas', width: 100 }, { key: 'canceladas', title: 'Canceladas', width: 100 }, { key: 'passageiros', title: 'Passageiros', width: 161.89 },
  ], report.drivers.slice(0, 20));
  table('Utilização por veículo', [
    { key: 'placa', title: 'Placa', width: 110 }, { key: 'marca', title: 'Marca', width: 115 }, { key: 'modelo', title: 'Modelo', width: 135 },
    { key: 'realizadas', title: 'Realizadas', width: 105 }, { key: 'agendadas', title: 'Agendadas', width: 100 }, { key: 'canceladas', title: 'Canceladas', width: 100 }, { key: 'passageiros', title: 'Passageiros', width: 106.89 },
  ], report.vehicles.slice(0, 20));
  table('Viagens por evento', [
    { key: 'nome', title: 'Evento', width: 180 }, { key: 'endereco', title: 'Local', width: 250 },
    { key: 'realizadas', title: 'Realizadas', width: 115 }, { key: 'agendadas', title: 'Agendadas', width: 110 }, { key: 'canceladas', title: 'Canceladas', width: 116.89 },
  ], report.events.filter((item) => item.realizadas + item.agendadas + item.canceladas > 0).slice(0, 30));
  table('Distribuição por dia da semana', [
    { key: 'nome', title: 'Dia', width: 220 }, { key: 'viagens', title: 'Total', width: 130 }, { key: 'realizadas', title: 'Realizadas', width: 140 }, { key: 'agendadas', title: 'Agendadas', width: 140 }, { key: 'canceladas', title: 'Canceladas', width: 141.89 },
  ], report.weekdays);
  table('Horários com viagens', [
    { key: 'hora', title: 'Horário', width: 200 }, { key: 'viagens', title: 'Total', width: 140 }, { key: 'realizadas', title: 'Realizadas', width: 144 }, { key: 'agendadas', title: 'Agendadas', width: 144 }, { key: 'canceladas', title: 'Canceladas', width: 143.89 },
  ], report.hours.filter((item) => item.viagens > 0));
  table('Rotas mais utilizadas', [
    { key: 'origem', title: 'Origem', width: 250 }, { key: 'destino', title: 'Destino', width: 250 }, { key: 'realizadas', title: 'Realizadas', width: 90 }, { key: 'agendadas', title: 'Agendadas', width: 90 }, { key: 'canceladas', title: 'Canceladas', width: 91.89 },
  ], report.routes);

  const pages = document.bufferedPageRange();
  for (let page = pages.start; page < pages.start + pages.count; page += 1) {
    document.switchToPage(page);
    document.font('Helvetica').fontSize(8).fillColor('#64748B').text(`Página ${page - pages.start + 1} de ${pages.count}`, margin, pageHeight - margin + 5, { width: contentWidth, align: 'right' });
  }
  document.end();
  return result;
}

module.exports = { aggregateCompanyReport, createCompanyWorkbook, createCompanyPdf, normalizeStatus, tripDate };
