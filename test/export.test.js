const test = require('node:test');
const assert = require('node:assert/strict');
const request = require('supertest');
const { app, createWorkbook, createPdfBuffer } = require('../src');

const event = {
  id: 1,
  nome: 'Evento de teste',
  idEndereco: { id: 10, rua: 'Rua Central', numero: '10', bairro: 'Centro', cidade: 'São Paulo', uf: 'SP' },
};
const trip = {
  id: 20,
  dataHora: '2026-10-20T08:30:00',
  idMotorista: { id: 2, nome: 'Ana', sobrenome: 'Silva', cpf: '123', rg: '456' },
  idVeiculo: 'ABC-1234',
  status: 'REALIZADA',
  localDeOrigem: event.idEndereco,
  localDeDestino: { id: 11, rua: 'Rua Norte', numero: '20', bairro: 'Bairro', cidade: 'Campinas', uf: 'SP' },
};

test('health endpoint is public and identifies the export service', async () => {
  const response = await request(app).get('/health');
  assert.equal(response.status, 200);
  assert.equal(response.body.service, 'loxecutivo-export');
});

test('report endpoints require a Bearer token before contacting Core', async () => {
  const excel = await request(app).get('/eventos/1/excel');
  const pdf = await request(app).get('/eventos/1/pdf');
  const company = await request(app).get('/relatorio-geral');
  const weekly = await request(app).get('/relatorio-geral/semanais');
  assert.equal(excel.status, 401);
  assert.equal(pdf.status, 401);
  assert.equal(company.status, 401);
  assert.equal(weekly.status, 401);
});

test('weekly report download rejects malformed IDs', async () => {
  const response = await request(app)
    .get('/relatorio-geral/semanais/not-a-date/pdf')
    .set('Authorization', 'Bearer token');
  assert.equal(response.status, 400);
});

test('company report validates date filters before contacting Core', async () => {
  const response = await request(app)
    .get('/relatorio-geral?de=2026-02-30')
    .set('Authorization', 'Bearer token');
  assert.equal(response.status, 400);
});

test('rejects invalid event identifiers', async () => {
  const response = await request(app).get('/eventos/not-a-number/pdf').set('Authorization', 'Bearer token');
  assert.equal(response.status, 400);
});

test('generates both report formats with the event data', async () => {
  const trips = [
    trip,
    { ...trip, id: 21, status: 'AGENDADA' },
    { ...trip, id: 22, status: 'CANCELADA' },
    { ...trip, id: 23, status: undefined },
  ];
  const passengers = new Map([
    [trip.id, [{ id: 30, nome: 'Pessoa' }, { id: 31, nome: 'Outra pessoa' }]],
    [21, [{ id: 32, nome: 'Passageiro agendado' }]],
    [22, [{ id: 33, nome: 'Passageiro cancelado' }]],
  ]);
  const vehicles = [{ idPlaca: trip.idVeiculo, marca: 'Marca', modelo: 'Modelo', ano: 2024, cor: 'Preto' }];

  const workbook = createWorkbook(event, trips, passengers, vehicles);
  const excelBuffer = await workbook.xlsx.writeBuffer();
  const pdfBuffer = await createPdfBuffer(event, trips, passengers, vehicles);

  assert.ok(excelBuffer.length > 1000);
  assert.equal(workbook.getWorksheet('Resumo').getCell('E5').value, 2);
  assert.equal(workbook.getWorksheet('Resumo').getCell('E6').value, 1);
  assert.equal(workbook.getWorksheet('Resumo').getCell('E7').value, 1);
  assert.equal(workbook.getWorksheet('Viagens').getCell('N5').value, 'Realizada');
  assert.equal(workbook.getWorksheet('Viagens').getCell('N6').value, 'Agendada');
  assert.equal(workbook.getWorksheet('Viagens').getCell('N7').value, 'Cancelada');
  assert.equal(workbook.getWorksheet('Viagens').getCell('N8').value, 'Agendada');
  assert.equal(workbook.getWorksheet('Resumo').getCell('B9').value, 2);
  assert.equal(pdfBuffer.subarray(0, 5).toString(), '%PDF-');
  assert.ok(pdfBuffer.length > 1000);
});
