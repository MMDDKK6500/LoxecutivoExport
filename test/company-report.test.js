const test = require('node:test');
const assert = require('node:assert/strict');
const { aggregateCompanyReport, createCompanyWorkbook, createCompanyPdf } = require('../src/companyReport');

const origin = { rua: 'Rua Central', numero: 10, bairro: 'Centro', cidade: 'São Paulo', uf: 'SP' };
const destination = { rua: 'Avenida Norte', numero: 20, bairro: 'Centro', cidade: 'São Paulo', uf: 'SP' };

function sampleReport(filters = {}) {
  return aggregateCompanyReport({
    events: [{ id: 7, nome: 'Evento corporativo', idEndereco: origin }],
    drivers: [
      { id: 2, nome: 'Ana', sobrenome: 'Silva' },
      { id: 3, nome: 'Bruno', sobrenome: 'Souza' },
    ],
    vehicles: [
      { idPlaca: 'ABC1234', marca: 'Marca', modelo: 'Sedã' },
      { idPlaca: 'DEF5678', marca: 'Marca', modelo: 'Van' },
    ],
    addresses: [origin, destination],
    trips: [
      { id: 100, dataHora: '2026-10-20T08:30:00', status: 'REALIZADA', idMotorista: { id: 2, nome: 'Ana', sobrenome: 'Silva' }, idVeiculo: 'ABC1234', idEvento: { id: 7, nome: 'Evento corporativo', idEndereco: origin }, localDeOrigem: origin, localDeDestino: destination },
      { id: 101, dataHora: '2026-10-20T09:00:00', status: 'AGENDADA', idMotorista: { id: 2, nome: 'Ana', sobrenome: 'Silva' }, idVeiculo: 'ABC1234', idEvento: { id: 7, nome: 'Evento corporativo', idEndereco: origin }, localDeOrigem: origin, localDeDestino: destination },
      { id: 102, dataHora: '2026-10-21T18:00:00', status: 'CANCELADA', idMotorista: { id: 3, nome: 'Bruno', sobrenome: 'Souza' }, idVeiculo: 'DEF5678', idEvento: { id: 7, nome: 'Evento corporativo', idEndereco: origin }, localDeOrigem: destination, localDeDestino: origin },
    ],
    passengers: [
      { id: 1, idViagem: { id: 100 } },
      { id: 2, idViagem: { id: 100 } },
      { id: 3, idViagem: { id: 101 } },
      { id: 4, idViagem: { id: 102 } },
    ],
  }, filters);
}

test('company summary uses completed trips for transported-passenger totals and usage counts', () => {
  const report = sampleReport();
  assert.deepEqual(report.statusCounts, { AGENDADA: 1, REALIZADA: 1, CANCELADA: 1 });
  assert.equal(report.summary.totalViagens, 3);
  assert.equal(report.summary.passageirosTransportados, 2);
  assert.equal(report.summary.mediaPassageirosPorViagemRealizada, 2);
  assert.equal(report.summary.eventosAtendidos, 1);
  assert.equal(report.summary.motoristasUtilizados, 1);
  assert.equal(report.summary.veiculosUtilizados, 1);
  assert.equal(report.drivers.find((driver) => driver.id === 2).realizadas, 1);
  assert.equal(report.vehicles.find((vehicle) => vehicle.placa === 'ABC1234').passageiros, 2);
  assert.equal(report.weekdays.find((day) => day.nome === 'Terça-feira').realizadas, 1);
  assert.equal(report.hours.find((hour) => hour.hora === '08:00').realizadas, 1);
});

test('company summary applies inclusive date filters', () => {
  const report = sampleReport({ from: '2026-10-21' });
  assert.equal(report.summary.totalViagens, 1);
  assert.equal(report.summary.canceladas, 1);
  assert.equal(report.summary.passageirosTransportados, 0);
});

test('company report generates Excel and PDF documents', async () => {
  const report = sampleReport();
  const workbook = createCompanyWorkbook(report);
  const excel = await workbook.xlsx.writeBuffer();
  const pdf = await createCompanyPdf(report);
  assert.ok(excel.length > 1500);
  assert.equal(workbook.getWorksheet('Resumo').getCell('B5').value, 3);
  assert.equal(workbook.getWorksheet('Motoristas').getCell('B4').value, 1);
  assert.equal(pdf.subarray(0, 5).toString(), '%PDF-');
  assert.ok(pdf.length > 1000);
});
