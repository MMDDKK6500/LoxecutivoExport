const fs = require('node:fs/promises');
const path = require('node:path');

function localDateParts(date, timeZone) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone, year: 'numeric', month: '2-digit', day: '2-digit', weekday: 'short', hour: '2-digit', hourCycle: 'h23',
  }).formatToParts(date);
  return Object.fromEntries(parts.map(({ type, value }) => [type, value]));
}

function mostRecentSunday(date = new Date(), timeZone = 'America/Sao_Paulo', scheduleHour = 0) {
  const local = localDateParts(date, timeZone);
  const weekdays = { Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6 };
  const offset = weekdays[local.weekday];
  const sunday = new Date(Date.UTC(Number(local.year), Number(local.month) - 1, Number(local.day) - offset));
  if (offset === 0 && Number(local.hour || 0) < scheduleHour) return null;
  return sunday.toISOString().slice(0, 10);
}

function reportPaths(directory, id) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(id)) {
    const error = new Error('O identificador do relatório deve estar no formato AAAA-MM-DD.');
    error.status = 400;
    throw error;
  }
  return {
    pdf: path.join(directory, `relatorio-semanal-${id}.pdf`),
    metadata: path.join(directory, `relatorio-semanal-${id}.json`),
  };
}

async function listWeeklyReports(directory) {
  let names;
  try { names = await fs.readdir(directory); } catch (error) {
    if (error.code === 'ENOENT') return [];
    throw error;
  }
  const metadataFiles = names.filter((name) => /^relatorio-semanal-\d{4}-\d{2}-\d{2}\.json$/.test(name));
  const reports = await Promise.all(metadataFiles.map(async (name) => {
    try {
      const report = JSON.parse(await fs.readFile(path.join(directory, name), 'utf8'));
      const { pdf } = reportPaths(directory, report.id);
      await fs.access(pdf);
      return report;
    } catch (_error) { return null; }
  }));
  return reports.filter(Boolean).sort((a, b) => b.id.localeCompare(a.id));
}

async function saveWeeklyReport(directory, id, report, generatedAt, createPdf) {
  await fs.mkdir(directory, { recursive: true });
  const paths = reportPaths(directory, id);
  try {
    await fs.access(paths.pdf);
    await fs.access(paths.metadata);
    return { created: false, metadata: JSON.parse(await fs.readFile(paths.metadata, 'utf8')) };
  } catch (_error) { /* Missing report: create it below. */ }

  const metadata = {
    id,
    generatedAt: generatedAt.toISOString(),
    period: report.period,
    summary: report.summary,
  };
  const suffix = `${process.pid}-${Date.now()}`;
  const temporaryPdf = `${paths.pdf}.${suffix}.tmp`;
  const temporaryMetadata = `${paths.metadata}.${suffix}.tmp`;
  try {
    await fs.writeFile(temporaryPdf, await createPdf(report), { flag: 'wx' });
    await fs.writeFile(temporaryMetadata, JSON.stringify(metadata, null, 2), { flag: 'wx' });
    await fs.rename(temporaryPdf, paths.pdf);
    await fs.rename(temporaryMetadata, paths.metadata);
  } finally {
    await Promise.all([fs.rm(temporaryPdf, { force: true }), fs.rm(temporaryMetadata, { force: true })]);
  }
  return { created: true, metadata };
}

function startWeeklyReportScheduler({ directory, timezone, scheduleHour, generate, intervalMs = 60_000, now = () => new Date() }) {
  let running = false;
  const run = async () => {
    if (running) return;
    const id = mostRecentSunday(now(), timezone, scheduleHour);
    if (!id) return;
    running = true;
    try {
      const created = await generate(id, directory);
      if (created) console.info(`Weekly company report created for ${id}.`);
    } catch (error) {
      console.error(`Could not create weekly report for ${id}:`, error);
    } finally { running = false; }
  };
  void run();
  const timer = setInterval(run, intervalMs);
  timer.unref?.();
  return () => clearInterval(timer);
}

module.exports = { mostRecentSunday, reportPaths, listWeeklyReports, saveWeeklyReport, startWeeklyReportScheduler };
