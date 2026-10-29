import 'dotenv/config';
import Fastify from 'fastify';
import fastifyStatic from '@fastify/static';
import { google } from 'googleapis';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const directory = path.dirname(fileURLToPath(import.meta.url));
const app = Fastify({ logger: true });
const spreadsheetId = process.env.GOOGLE_SHEET_ID;
const serviceAccount = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;

function getSheets() {
  if (!spreadsheetId || !serviceAccount) throw new Error('尚未設定 GOOGLE_SHEET_ID 或 GOOGLE_SERVICE_ACCOUNT_JSON');
  const auth = new google.auth.GoogleAuth({ credentials: JSON.parse(serviceAccount), scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
  return google.sheets({ version: 'v4', auth });
}
async function readRange(range) {
  const result = await getSheets().spreadsheets.values.get({ spreadsheetId, range });
  return result.data.values || [];
}
function rowsToObjects(values) {
  const [headers = [], ...rows] = values;
  return rows.map(row => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ''])));
}
function now() {
  return new Intl.DateTimeFormat('sv-SE', { timeZone: process.env.TIMEZONE || 'Asia/Hong_Kong', dateStyle: 'short', timeStyle: 'medium' }).format(new Date());
}

app.post('/api', async (request, reply) => {
  try {
    const body = request.body || {};
    if (body.action === 'searchStudent') {
      const students = rowsToObjects(await readRange(process.env.STUDENTS_RANGE || 'Students!A1:D'));
      const query = String(body.studentId || '').trim().toLowerCase();
      const row = students.find(item => Object.values(item).some(value => String(value).trim().toLowerCase() === query));
      if (!row) return { status: 'not_found' };
      return { status: 'found', student: {
        studentNum: row.studentNum || row.編號 || '', name: row.name || row.姓名 || '',
        studentClass: row.studentClass || row.班別 || '', studentId: row.studentId || row.學號 || query
      }};
    }
    if (body.action === 'saveSitUpBatch') {
      const students = Array.isArray(body.students) ? body.students : [];
      const duration = Number(body.duration);
      if (students.length < 1 || students.length > 10 || !Number.isInteger(duration) || duration < 1) {
        return reply.code(400).send({ status: 'error', message: '學生數量或測驗時間格式不正確' });
      }
      const rows = students.map(student => [now(), student.studentClass || '', student.studentId || '', student.name || '', Number(student.count), duration]);
      if (rows.some(row => !row[3] || !Number.isInteger(row[4]) || row[4] < 0)) {
        return reply.code(400).send({ status: 'error', message: '學生資料或次數格式不正確' });
      }
      await getSheets().spreadsheets.values.append({
        spreadsheetId, range: process.env.SIT_UP_LOGS_RANGE || 'SitUpLogs!A:F',
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: rows }
      });
      return { status: 'success' };
    }
    if (body.action === 'getSitUpLogs') {
      const logs = rowsToObjects(await readRange(process.env.SIT_UP_LOGS_RANGE || 'SitUpLogs!A1:F'));
      return { status: 'success', logs: logs.slice(-100).reverse().map(row => ({
        time: row.time || row.時間 || '', studentClass: row.studentClass || row.班別 || '',
        name: row.name || row.姓名 || '', count: row.count || row.次數 || '', duration: row.duration || row.秒數 || ''
      })) };
    }
    return reply.code(400).send({ status: 'error', message: 'Unknown action' });
  } catch (error) {
    request.log.error(error);
    return reply.code(500).send({ status: 'error', message: error.message });
  }
});
app.get('/health', async () => ({ status: 'ok' }));
app.register(fastifyStatic, { root: directory, index: ['sit-up.html'] });
app.listen({ port: Number(process.env.PORT || 3010), host: process.env.HOST || '127.0.0.1' })
  .then(address => app.log.info(`Server listening at ${address}`))
  .catch(error => { app.log.error(error); process.exit(1); });
