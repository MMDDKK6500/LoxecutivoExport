const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { mostRecentSunday, listWeeklyReports, saveWeeklyReport, reportPaths } = require('../src/weeklyReports');

test('weekly schedule chooses the local Sunday and waits until the configured hour', () => {
  assert.equal(mostRecentSunday(new Date('2026-10-18T04:00:00Z'), 'America/Sao_Paulo', 1), '2026-10-18');
  assert.equal(mostRecentSunday(new Date('2026-10-18T03:30:00Z'), 'America/Sao_Paulo', 1), null);
  assert.equal(mostRecentSunday(new Date('2026-10-19T12:00:00Z'), 'America/Sao_Paulo', 1), '2026-10-18');
});

test('weekly reports save once, list newest first, and point to stored PDFs', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'loxecutivo-weekly-'));
  try {
    const report = { period: { from: null, to: null }, summary: { totalViagens: 3 } };
    const generatedAt = new Date('2026-10-18T04:00:00Z');
    let pdfCalls = 0;
    const createPdf = async () => { pdfCalls += 1; return Buffer.from('%PDF-test'); };
    const first = await saveWeeklyReport(directory, '2026-10-18', report, generatedAt, createPdf);
    const duplicate = await saveWeeklyReport(directory, '2026-10-18', report, generatedAt, createPdf);
    await saveWeeklyReport(directory, '2026-10-11', report, generatedAt, createPdf);

    assert.equal(first.created, true);
    assert.equal(duplicate.created, false);
    assert.equal(pdfCalls, 2);
    assert.deepEqual((await listWeeklyReports(directory)).map(({ id }) => id), ['2026-10-18', '2026-10-11']);
    assert.equal((await fs.readFile(reportPaths(directory, '2026-10-18').pdf)).toString(), '%PDF-test');
    assert.throws(() => reportPaths(directory, '../escape'));
  } finally { await fs.rm(directory, { recursive: true, force: true }); }
});
