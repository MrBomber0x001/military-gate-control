const express = require('express');
const Database = require('better-sqlite3');
const path = require('path');
const crypto = require('crypto');
const ExcelJS = require('exceljs');
const fs = require('fs');
const fsp = fs.promises;

const PUB = path.join(__dirname, 'public');
const DB_PATH = path.resolve(process.env.GATE_DB_PATH || path.join(__dirname, 'gate.db'));
const BACKUP_DIR = path.resolve(process.env.GATE_BACKUP_DIR || path.join(path.dirname(DB_PATH), 'backups'));
const BACKUP_NAME = /^gate-\d{4}-\d{2}-\d{2}-\d{2}-\d{2}-\d{2}-(?:daily|manual|pre-restore)\.db$/;
const BACKUP_RETENTION = 30;
const db = new Database(DB_PATH);
db.pragma('journal_mode = WAL');
db.pragma('foreign_keys = ON');
db.exec(`
CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL, pass TEXT NOT NULL, full_name TEXT, role TEXT NOT NULL DEFAULT 'operator', active INTEGER NOT NULL DEFAULT 1);
CREATE TABLE IF NOT EXISTS sessions(token TEXT PRIMARY KEY, user_id INTEGER NOT NULL, expires INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS lookups(id INTEGER PRIMARY KEY AUTOINCREMENT, kind TEXT NOT NULL, name TEXT NOT NULL, UNIQUE(kind,name));
CREATE TABLE IF NOT EXISTS personnel(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, mid TEXT UNIQUE NOT NULL, rank TEXT, unit TEXT, phone TEXT, inside INTEGER NOT NULL DEFAULT 0, retired INTEGER NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS vehicles(id INTEGER PRIMARY KEY AUTOINCREMENT, plate TEXT UNIQUE NOT NULL, vehicle TEXT NOT NULL, driver_id INTEGER REFERENCES personnel(id) ON DELETE SET NULL, purpose TEXT, inside INTEGER NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS items(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, category TEXT, unit TEXT, stock INTEGER NOT NULL DEFAULT 0, reorder INTEGER NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS movements(id INTEGER PRIMARY KEY AUTOINCREMENT, kind TEXT NOT NULL, ref_id INTEGER, ref TEXT NOT NULL, type TEXT NOT NULL, ts TEXT NOT NULL, gate TEXT, destination TEXT, qty INTEGER, by TEXT);
CREATE INDEX IF NOT EXISTS idx_mov_ts ON movements(ts);
`);

// Default dropdown lists (admin can add/remove from the admin console)
const SEED = {
  unit: ['الأولى', 'الثانية', 'الثالثة'], vehicle: ['سيدان', 'بيك أب', 'فان', 'دفع رباعي'],
  purpose: ['رسمي', 'إمداد', 'صيانة', 'زائر'], destination: [], visit_reason: [], gate: ['البوابة الرئيسية', 'البوابة الجانبية'],
};
if (!db.prepare('SELECT COUNT(*) c FROM lookups').get().c) {
  const ins = db.prepare('INSERT INTO lookups(kind,name) VALUES (?,?)');
  db.transaction(() => Object.entries(SEED).forEach(([k, a]) => a.forEach(n => ins.run(k, n))))();
}

// Ranks are fixed (they decide the group): جندي | صف ضابط | ضابط | مدني.
const RANKS = [['جندي', 'جندي'], ['رقيب', 'صف ضابط'], ['رقيب أ', 'صف ضابط'], ['مساعد', 'صف ضابط'], ['مساعد أ', 'صف ضابط'],
  ['ملازم', 'ضابط'], ['ملازم أول', 'ضابط'], ['نقيب', 'ضابط'], ['رائد', 'ضابط'], ['مقدم', 'ضابط'], ['عقيد', 'ضابط'], ['عميد', 'ضابط'], ['لواء', 'ضابط'],
  ['مدني', 'مدني']]
  .map(([name, cat], ord) => ({ name, cat, ord, grp: cat === 'ضابط' ? 'officers' : cat === 'مدني' ? 'civilians' : cat === 'صف ضابط' ? 'ncos' : 'soldiers' }));
const RK = Object.fromEntries(RANKS.map(r => [r.name, r]));

db.exec(`
CREATE TABLE IF NOT EXISTS attendance(date TEXT NOT NULL, person_id INTEGER NOT NULL REFERENCES personnel(id) ON DELETE CASCADE, by TEXT, PRIMARY KEY(date, person_id));
CREATE TABLE IF NOT EXISTS duty(id INTEGER PRIMARY KEY AUTOINCREMENT, person_id INTEGER REFERENCES personnel(id) ON DELETE SET NULL, person TEXT NOT NULL, start_date TEXT NOT NULL, end_date TEXT NOT NULL, notes TEXT, by TEXT);
CREATE TABLE IF NOT EXISTS missions(id INTEGER PRIMARY KEY AUTOINCREMENT, person_id INTEGER REFERENCES personnel(id) ON DELETE SET NULL, person TEXT NOT NULL, dest TEXT NOT NULL, purpose TEXT, notes TEXT, start_date TEXT NOT NULL, start_ts TEXT, end_ts TEXT, status TEXT NOT NULL DEFAULT 'active', by TEXT);
CREATE TABLE IF NOT EXISTS purchases(id INTEGER PRIMARY KEY AUTOINCREMENT, person_id INTEGER REFERENCES personnel(id) ON DELETE SET NULL, person TEXT NOT NULL, date TEXT NOT NULL, notes TEXT, by TEXT);
CREATE TABLE IF NOT EXISTS purchase_lines(id INTEGER PRIMARY KEY AUTOINCREMENT, purchase_id INTEGER NOT NULL REFERENCES purchases(id) ON DELETE CASCADE, item TEXT NOT NULL, price_cents INTEGER NOT NULL, qty INTEGER NOT NULL);
CREATE INDEX IF NOT EXISTS idx_pur_date ON purchases(date);
CREATE TABLE IF NOT EXISTS clinics(id INTEGER PRIMARY KEY AUTOINCREMENT, person_id INTEGER REFERENCES personnel(id) ON DELETE SET NULL, person TEXT NOT NULL, case_details TEXT NOT NULL, hospital TEXT, start_ts TEXT NOT NULL, end_ts TEXT, status TEXT NOT NULL DEFAULT 'active', by TEXT);
CREATE TABLE IF NOT EXISTS leaves(id INTEGER PRIMARY KEY AUTOINCREMENT, person_id INTEGER REFERENCES personnel(id) ON DELETE SET NULL, person TEXT NOT NULL, start_date TEXT NOT NULL, end_date TEXT NOT NULL, reason TEXT, by TEXT, expected_return_at TEXT, departed_at TEXT, returned_at TEXT);
CREATE TABLE IF NOT EXISTS visitors(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, rank TEXT NOT NULL, reason TEXT NOT NULL DEFAULT '', unit TEXT NOT NULL DEFAULT '', enter_ts TEXT NOT NULL, exit_ts TEXT, by TEXT);
CREATE TABLE IF NOT EXISTS squads(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, start_date TEXT NOT NULL, end_date TEXT NOT NULL, by TEXT);
CREATE TABLE IF NOT EXISTS squad_members(squad_id INTEGER NOT NULL REFERENCES squads(id) ON DELETE CASCADE, person_id INTEGER NOT NULL REFERENCES personnel(id) ON DELETE CASCADE, PRIMARY KEY(squad_id, person_id));
CREATE TABLE IF NOT EXISTS squad_attendance(id INTEGER PRIMARY KEY AUTOINCREMENT, squad_id INTEGER NOT NULL, person_id INTEGER NOT NULL, date TEXT NOT NULL, enter_ts TEXT NOT NULL, exit_ts TEXT, by TEXT, UNIQUE(squad_id, person_id, date), FOREIGN KEY(squad_id, person_id) REFERENCES squad_members(squad_id, person_id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS squad_absences(squad_id INTEGER NOT NULL, person_id INTEGER NOT NULL, date TEXT NOT NULL, by TEXT NOT NULL, created_at TEXT NOT NULL, PRIMARY KEY(squad_id, person_id, date), FOREIGN KEY(squad_id, person_id) REFERENCES squad_members(squad_id, person_id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS notifications(id INTEGER PRIMARY KEY AUTOINCREMENT, recipient_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE, actor_id INTEGER NOT NULL, actor_name TEXT NOT NULL, type TEXT NOT NULL, summary TEXT NOT NULL, created_at TEXT NOT NULL, read_at TEXT);
CREATE TABLE IF NOT EXISTS vehicle_assignments(date TEXT NOT NULL, vehicle_id INTEGER NOT NULL REFERENCES vehicles(id) ON DELETE CASCADE, driver_id INTEGER REFERENCES personnel(id) ON DELETE SET NULL, companion_id INTEGER REFERENCES personnel(id) ON DELETE SET NULL, by TEXT NOT NULL, updated_at TEXT NOT NULL, PRIMARY KEY(date, vehicle_id));
CREATE TABLE IF NOT EXISTS vehicle_assignment_passengers(date TEXT NOT NULL, vehicle_id INTEGER NOT NULL, person_id INTEGER NOT NULL REFERENCES personnel(id) ON DELETE CASCADE, PRIMARY KEY(date, vehicle_id, person_id), FOREIGN KEY(date, vehicle_id) REFERENCES vehicle_assignments(date, vehicle_id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS tasks(id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL, details TEXT NOT NULL DEFAULT '', created_by INTEGER REFERENCES users(id) ON DELETE SET NULL, created_by_name TEXT NOT NULL, created_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS task_assignments(id INTEGER PRIMARY KEY AUTOINCREMENT, task_id INTEGER NOT NULL REFERENCES tasks(id) ON DELETE CASCADE, recipient_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE, confirmed_at TEXT, UNIQUE(task_id, recipient_id));
CREATE TABLE IF NOT EXISTS daily_handovers(date TEXT PRIMARY KEY, summary_json TEXT NOT NULL, notes TEXT NOT NULL DEFAULT '', prepared_by INTEGER REFERENCES users(id) ON DELETE SET NULL, prepared_by_name TEXT NOT NULL, prepared_at TEXT NOT NULL, acknowledged_by INTEGER REFERENCES users(id) ON DELETE SET NULL, acknowledged_by_name TEXT, acknowledged_at TEXT);
CREATE TABLE IF NOT EXISTS correspondence(id INTEGER PRIMARY KEY AUTOINCREMENT, reference_number TEXT NOT NULL DEFAULT '', kind TEXT NOT NULL, correspondence_date TEXT NOT NULL, sender TEXT NOT NULL DEFAULT '', recipient TEXT NOT NULL DEFAULT '', responsible_section TEXT NOT NULL DEFAULT '', subject TEXT NOT NULL, summary TEXT NOT NULL DEFAULT '', assigned_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL, status TEXT NOT NULL DEFAULT 'new', due_date TEXT, created_by INTEGER REFERENCES users(id) ON DELETE SET NULL, created_by_name TEXT NOT NULL, created_at TEXT NOT NULL, updated_at TEXT NOT NULL, updated_by TEXT NOT NULL);
CREATE UNIQUE INDEX IF NOT EXISTS idx_correspondence_reference ON correspondence(reference_number) WHERE reference_number<>'';
CREATE INDEX IF NOT EXISTS idx_correspondence_date ON correspondence(correspondence_date);
CREATE INDEX IF NOT EXISTS idx_correspondence_assignee ON correspondence(assigned_user_id);
CREATE TABLE IF NOT EXISTS correspondence_attachments(id INTEGER PRIMARY KEY AUTOINCREMENT, correspondence_id INTEGER NOT NULL REFERENCES correspondence(id) ON DELETE CASCADE, filename TEXT NOT NULL, mime_type TEXT NOT NULL, size INTEGER NOT NULL, content BLOB NOT NULL);
CREATE TABLE IF NOT EXISTS correspondence_messages(id INTEGER PRIMARY KEY AUTOINCREMENT, correspondence_id INTEGER NOT NULL REFERENCES correspondence(id) ON DELETE CASCADE, user_id INTEGER REFERENCES users(id) ON DELETE SET NULL, author_name TEXT NOT NULL, body TEXT NOT NULL, created_at TEXT NOT NULL);
CREATE INDEX IF NOT EXISTS idx_correspondence_messages_parent ON correspondence_messages(correspondence_id,id);
CREATE TABLE IF NOT EXISTS correspondence_message_attachments(id INTEGER PRIMARY KEY AUTOINCREMENT, message_id INTEGER NOT NULL REFERENCES correspondence_messages(id) ON DELETE CASCADE, filename TEXT NOT NULL, mime_type TEXT NOT NULL, size INTEGER NOT NULL, content BLOB NOT NULL);
CREATE INDEX IF NOT EXISTS idx_correspondence_message_attachments_parent ON correspondence_message_attachments(message_id,id);
CREATE INDEX IF NOT EXISTS idx_clinics_start ON clinics(start_ts);
CREATE INDEX IF NOT EXISTS idx_leaves_dates ON leaves(start_date, end_date);
CREATE INDEX IF NOT EXISTS idx_visitors_enter ON visitors(enter_ts);
CREATE INDEX IF NOT EXISTS idx_visitors_name ON visitors(name COLLATE NOCASE, id DESC);
CREATE INDEX IF NOT EXISTS idx_squad_attendance_date ON squad_attendance(date);
CREATE INDEX IF NOT EXISTS idx_squad_absences_date ON squad_absences(date);
CREATE INDEX IF NOT EXISTS idx_notifications_recipient ON notifications(recipient_id, id DESC);
CREATE INDEX IF NOT EXISTS idx_vehicle_assignments_date ON vehicle_assignments(date);
CREATE INDEX IF NOT EXISTS idx_vehicle_assignment_passengers ON vehicle_assignment_passengers(date, vehicle_id);
CREATE INDEX IF NOT EXISTS idx_task_assignments_recipient ON task_assignments(recipient_id, id DESC);
`);
const notificationColumns = new Set(db.prepare('PRAGMA table_info(notifications)').all().map(column => column.name));
if (!notificationColumns.has('title')) db.exec("ALTER TABLE notifications ADD COLUMN title TEXT NOT NULL DEFAULT ''");
if (!notificationColumns.has('body')) db.exec("ALTER TABLE notifications ADD COLUMN body TEXT NOT NULL DEFAULT ''");
if (!notificationColumns.has('related_id')) db.exec('ALTER TABLE notifications ADD COLUMN related_id INTEGER');
const columns = table => db.prepare(`PRAGMA table_info(${table})`).all().map(c => c.name);
if (!columns('vehicles').includes('companion_id')) db.exec('ALTER TABLE vehicles ADD COLUMN companion_id INTEGER REFERENCES personnel(id) ON DELETE SET NULL');
if (!columns('personnel').includes('retired')) db.exec('ALTER TABLE personnel ADD COLUMN retired INTEGER NOT NULL DEFAULT 0');
if (!columns('visitors').includes('reason')) db.exec("ALTER TABLE visitors ADD COLUMN reason TEXT NOT NULL DEFAULT ''");
if (!columns('visitors').includes('unit')) db.exec("ALTER TABLE visitors ADD COLUMN unit TEXT NOT NULL DEFAULT ''");
if (!columns('movements').includes('destination')) db.exec('ALTER TABLE movements ADD COLUMN destination TEXT');
if (!columns('movements').includes('details')) db.exec('ALTER TABLE movements ADD COLUMN details TEXT');
if (!columns('duty').includes('role')) db.exec("ALTER TABLE duty ADD COLUMN role TEXT NOT NULL DEFAULT 'general'");
if (!columns('leaves').includes('expected_return_at')) db.exec('ALTER TABLE leaves ADD COLUMN expected_return_at TEXT');
if (!columns('leaves').includes('departed_at')) db.exec('ALTER TABLE leaves ADD COLUMN departed_at TEXT');
if (!columns('leaves').includes('returned_at')) db.exec('ALTER TABLE leaves ADD COLUMN returned_at TEXT');
if (!columns('missions').includes('start_ts')) db.exec('ALTER TABLE missions ADD COLUMN start_ts TEXT');
db.exec("UPDATE missions SET start_ts=start_date || ' 00:00:00' WHERE start_ts IS NULL");
// cleanup of the old canteen / rank-list data (ranks are now fixed in code)
db.exec("DELETE FROM movements WHERE kind='canteen'; DELETE FROM lookups WHERE kind IN ('rank','item_category');");

// ---------- helpers ----------
const pad = n => String(n).padStart(2, '0');
const now = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`; };
const day = d => (/^\d{4}-\d{2}-\d{2}$/.test(d) ? d : now().slice(0, 10));
const DUTY_ROLES = {
  general: 'نوبتجي',
  assistant_officer: 'مساعد ظابط نباطشي',
  duty_officer: 'ظابط نباطشي',
  security_commander: 'حكمدار الأمن',
  services_commander: 'حكمدار الخدمات',
};
const dutyRoleLabel = role => DUTY_ROLES[role] || DUTY_ROLES.general;
const sha = s => crypto.createHash('sha256').update(s).digest('hex');
const hash = p => { const s = crypto.randomBytes(16); return s.toString('hex') + ':' + crypto.scryptSync(p, s, 64).toString('hex'); };
const verify = (p, st) => { const [s, h] = st.split(':'), a = crypto.scryptSync(p, Buffer.from(s, 'hex'), 64), b = Buffer.from(h, 'hex'); return a.length === b.length && crypto.timingSafeEqual(a, b); };
const okPw = p => typeof p === 'string' && p.length >= 8;
const okUser = u => /^[a-z0-9_.-]{3,32}$/.test(u);
const cookie = (req, n) => ((req.headers.cookie || '').split(';').map(s => s.trim().split('=')).find(a => a[0] === n) || [])[1];
const publicUser = u => ({ id: u.id, username: u.username, full_name: u.full_name, role: u.role });
const timestampForFile = date => `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}-${pad(date.getHours())}-${pad(date.getMinutes())}-${pad(date.getSeconds())}`;
const backupName = (kind, date = new Date()) => `gate-${timestampForFile(date)}-${kind}.db`;
const backupPath = name => path.join(BACKUP_DIR, name);
let backupBusy = false;

async function pruneBackups(protectedNames = []) {
  const files = (await fsp.readdir(BACKUP_DIR))
    .filter(name => BACKUP_NAME.test(name))
    .map(async name => ({ name, mtime: (await fsp.stat(backupPath(name))).mtimeMs }));
  const ordered = (await Promise.all(files)).sort((a, b) => b.mtime - a.mtime);
  const protectedFiles = ordered.filter(file => protectedNames.includes(file.name));
  const keep = new Set([
    ...protectedFiles.map(file => file.name),
    ...ordered.filter(file => !protectedNames.includes(file.name))
      .slice(0, Math.max(0, BACKUP_RETENTION - protectedFiles.length))
      .map(file => file.name),
  ]);
  await Promise.all(ordered.filter(file => !keep.has(file.name)).map(file => fsp.unlink(backupPath(file.name))));
}

async function writeBackup(kind, protectedNames = []) {
  await fsp.mkdir(BACKUP_DIR, { recursive: true });
  let stamp = new Date(), destination = backupPath(backupName(kind, stamp));
  while (fs.existsSync(destination)) {
    stamp = new Date(stamp.getTime() + 1000);
    destination = backupPath(backupName(kind, stamp));
  }
  const temporary = `${destination}.${process.pid}.tmp`;
  try {
    await db.backup(temporary);
    await fsp.rename(temporary, destination);
    await pruneBackups(protectedNames);
    return destination;
  } catch (error) {
    await fsp.rm(temporary, { force: true }).catch(cleanupError => {
      console.error('Unable to remove incomplete database backup:', cleanupError);
    });
    throw error;
  }
}

async function createBackup(kind) {
  if (backupBusy) throw new Error('هناك عملية نسخ احتياطي أو استعادة قيد التنفيذ');
  backupBusy = true;
  try {
    return await writeBackup(kind);
  } finally {
    backupBusy = false;
  }
}

async function ensureDailyBackup() {
  if (backupBusy) return;
  const dayPrefix = `gate-${timestampForFile(new Date()).slice(0, 10)}-`;
  await fsp.mkdir(BACKUP_DIR, { recursive: true });
  const todaysBackup = (await fsp.readdir(BACKUP_DIR)).some(name => BACKUP_NAME.test(name) && name.startsWith(dayPrefix));
  if (!todaysBackup) await createBackup('daily');
}

async function backupFiles() {
  await fsp.mkdir(BACKUP_DIR, { recursive: true });
  const files = (await fsp.readdir(BACKUP_DIR)).filter(name => BACKUP_NAME.test(name));
  const rows = await Promise.all(files.map(async name => {
    const stat = await fsp.stat(backupPath(name));
    return {
      name,
      size: stat.size,
      created_at: stat.mtime.toISOString(),
      type: name.endsWith('-daily.db') ? 'تلقائية يومية' : name.endsWith('-pre-restore.db') ? 'حفظ قبل الاستعادة' : 'يدوية',
    };
  }));
  return rows.sort((a, b) => b.created_at.localeCompare(a.created_at));
}

async function validateBackup(file) {
  const source = new Database(file, { readonly: true, fileMustExist: true });
  try {
    const integrity = source.pragma('integrity_check', { simple: true });
    const requiredTables = ['users', 'personnel', 'movements'];
    const tables = source.prepare("SELECT name FROM sqlite_master WHERE type='table'").all().map(row => row.name);
    if (integrity !== 'ok' || requiredTables.some(table => !tables.includes(table)))
      throw new Error('النسخة الاحتياطية تالفة أو غير صالحة للاستعادة');
  } finally {
    source.close();
  }
}

function startSession(res, userId) {
  const t = crypto.randomBytes(32).toString('hex');
  db.prepare('DELETE FROM sessions WHERE expires < ?').run(Date.now());
  db.prepare('INSERT INTO sessions(token,user_id,expires) VALUES (?,?,?)').run(sha(t), userId, Date.now() + 12 * 3600e3);
  res.setHeader('Set-Cookie', `sid=${t}; HttpOnly; SameSite=Strict; Path=/; Max-Age=43200${process.env.COOKIE_SECURE ? '; Secure' : ''}`);
}
const session = req => { const t = cookie(req, 'sid'); return t && db.prepare('SELECT u.* FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token=? AND s.expires>? AND u.active=1').get(sha(t), Date.now()); };
const requireAuth = (req, res, next) => { const u = session(req); if (!u) return res.status(401).json({ error: 'يجب تسجيل الدخول' }); req.user = u; next(); };
const admin = (req, res, next) => req.user.role === 'admin' ? next() : res.status(403).json({ error: 'هذا الإجراء للمدير فقط' });
const page = adminOnly => (req, res, next) => { const u = session(req); if (!u) return res.redirect('/login.html'); if (adminOnly && u.role !== 'admin') return res.redirect('/'); next(); };

const app = express();
app.use(express.json({ limit: '22mb' }));
app.get(['/', '/index.html'], page(false), (q, r) => r.sendFile(path.join(PUB, 'index.html')));
app.get('/admin.html', page(true), (q, r) => r.sendFile(path.join(PUB, 'admin.html')));
app.use(express.static(PUB, { index: false }));

// ---------- auth ----------
const fails = new Map();
app.get('/api/auth/status', (q, r) => r.json({ needsSetup: !db.prepare('SELECT COUNT(*) c FROM users').get().c }));

// The very first account created becomes the admin; this route locks itself afterwards.
app.post('/api/auth/setup', (req, res) => {
  if (db.prepare('SELECT COUNT(*) c FROM users').get().c) return res.status(403).json({ error: 'تم إعداد النظام مسبقاً' });
  const u = String(req.body.username || '').trim().toLowerCase();
  if (!okUser(u)) return res.status(400).json({ error: 'اسم المستخدم: ٣–٣٢ حرفاً إنجليزياً أو أرقاماً' });
  if (!okPw(req.body.password)) return res.status(400).json({ error: 'كلمة المرور ٨ أحرف على الأقل' });
  const r = db.prepare('INSERT INTO users(username,pass,full_name,role) VALUES (?,?,?,?)').run(u, hash(req.body.password), String(req.body.full_name || '').trim(), 'admin');
  startSession(res, r.lastInsertRowid);
  res.json({ ok: true });
});

app.post('/api/auth/login', (req, res) => {
  const username = String(req.body.username || '').trim().toLowerCase(), password = String(req.body.password || '');
  const key = req.ip + '|' + username, f = fails.get(key), fresh = f && Date.now() - f.t < 9e5;
  if (fresh && f.n >= 5) return res.status(429).json({ error: 'محاولات كثيرة، حاول بعد ١٥ دقيقة' });
  const u = db.prepare('SELECT * FROM users WHERE username=? AND active=1').get(username);
  if (!u || !verify(password, u.pass)) { fails.set(key, { n: (fresh ? f.n : 0) + 1, t: Date.now() }); return res.status(401).json({ error: 'اسم المستخدم أو كلمة المرور غير صحيحة' }); }
  fails.delete(key); startSession(res, u.id); res.json(publicUser(u));
});

app.post('/api/auth/logout', (req, res) => {
  const t = cookie(req, 'sid'); if (t) db.prepare('DELETE FROM sessions WHERE token=?').run(sha(t));
  res.setHeader('Set-Cookie', 'sid=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0'); res.json({ ok: true });
});

app.use('/api', requireAuth);
app.get('/api/me', (req, res) => res.json(publicUser(req.user)));

function mutationDescription(req) {
  const method = req.method, pathname = new URL(req.originalUrl, 'http://localhost').pathname;
  const idName = (table, id, column) => id ? db.prepare(`SELECT ${column} AS name FROM ${table} WHERE id=?`).get(id)?.name : null;
  const label = value => String(value || '').trim();
  let type = 'system', action = 'إجراء', target = '';
  let match;
  if ((match = pathname.match(/^\/api\/admin\/(personnel|vehicles)(?:\/(\d+))?$/))) {
    type = match[1] === 'personnel' ? 'personnel' : 'vehicle';
    action = method === 'POST' ? 'إضافة' : method === 'PUT' ? 'تعديل' : 'حذف';
    target = match[2] ? idName(match[1], match[2], match[1] === 'personnel' ? 'name' : 'plate') : label(req.body.name || req.body.plate);
  } else if (/^\/api\/admin\/lookups(?:\/\d+)?$/.test(pathname)) {
    type = 'lookup'; action = method === 'POST' ? 'إضافة عنصر قائمة' : 'حذف عنصر قائمة'; target = label(req.body.name);
  } else if (/^\/api\/admin\/users(?:\/\d+)?$/.test(pathname)) {
    type = 'user'; action = method === 'POST' ? 'إنشاء مستخدم' : method === 'PUT' ? 'تعديل مستخدم' : 'حذف مستخدم';
    target = label(req.body.full_name || req.body.username);
  } else if (pathname === '/api/admin/tasks') {
    type = 'task'; action = 'إرسال تكليف'; target = label(req.body.title);
  } else if (pathname === '/api/admin/member-notifications') {
    type = 'notification'; action = 'إرسال إشعار للأعضاء'; target = label(req.body.title);
  } else if ((match = pathname.match(/^\/api\/(?:admin\/)?correspondence(?:\/(\d+)(?:\/status)?)?$/))) {
    type = 'correspondence';
    action = method === 'POST' ? 'تسجيل مكاتبة' : method === 'DELETE' ? 'حذف مكاتبة' : 'تحديث حالة مكاتبة';
    target = label(req.body.subject) || (match[1] ? db.prepare('SELECT subject FROM correspondence WHERE id=?').get(match[1])?.subject : '');
  } else if ((match = pathname.match(/^\/api\/tasks\/(\d+)\/confirm$/))) {
    type = 'task'; action = 'تأكيد استلام تكليف';
    target = db.prepare('SELECT t.title FROM task_assignments a JOIN tasks t ON t.id=a.task_id WHERE a.id=?').get(match[1])?.title || '';
  } else if (pathname === '/api/admin/backups') {
    type = 'system'; action = 'إنشاء نسخة احتياطية';
  } else if (/^\/api\/admin\/backups\/[^/]+\/restore$/.test(pathname)) {
    type = 'system'; action = 'استعادة نسخة احتياطية';
  } else if (pathname === '/api/auth/change-password') {
    type = 'user'; action = 'تغيير كلمة المرور'; target = req.user.username;
  } else if ((match = pathname.match(/^\/api\/(personnel|vehicles)\/(\d+)\/toggle$/))) {
    type = match[1] === 'personnel' ? 'personnel' : 'vehicle';
    const row = db.prepare(`SELECT ${match[1] === 'personnel' ? 'name' : 'plate'} AS name, inside FROM ${match[1]} WHERE id=?`).get(match[2]);
    action = row?.inside ? 'تسجيل خروج' : 'تسجيل دخول'; target = row?.name || '';
  } else if (pathname === '/api/vehicles/manual-movement') {
    type = 'vehicle'; action = req.body.type === 'OUT' ? 'خروج مركبة زائرة' : 'دخول مركبة زائرة'; target = label(req.body.plate);
  } else if (pathname === '/api/attendance') {
    type = 'attendance'; action = req.body.present ? 'تسجيل تمام' : 'إلغاء تمام';
    const ids = (Array.isArray(req.body.ids) ? req.body.ids : []).map(Number).filter(Number.isInteger).slice(0, 2000);
    target = ids.length ? `${ids.length} فرد` : '';
  } else if (pathname === '/api/duty' || /^\/api\/admin\/duty\/\d+$/.test(pathname)) {
    type = 'duty'; action = method === 'DELETE' ? 'حذف نوبتجية' : 'تسجيل نوبتجية';
    target = method === 'DELETE'
      ? label(req.body.notes)
      : [idName('personnel', req.body.person_id, 'name'), DUTY_ROLES[req.body.role] || DUTY_ROLES.general].filter(Boolean).join(' — ');
  } else if (pathname === '/api/missions' || /^\/api\/missions\/\d+\/return$/.test(pathname) || /^\/api\/admin\/missions\/\d+$/.test(pathname)) {
    type = 'mission'; action = method === 'DELETE' ? 'حذف مأمورية' : pathname.endsWith('/return') ? 'تسجيل عودة من مأمورية' : 'تسجيل مأمورية';
    target = idName('personnel', req.body.person_id, 'name') || label(req.body.dest);
  } else if (pathname === '/api/clinics' || /^\/api\/clinics\/\d+\/return$/.test(pathname) || /^\/api\/admin\/clinics\/\d+$/.test(pathname)) {
    type = 'clinic'; action = method === 'DELETE' ? 'حذف حالة عيادة' : pathname.endsWith('/return') ? 'تسجيل عودة من العيادة' : 'تسجيل خروج للعيادة';
    target = idName('personnel', req.body.person_id, 'name') || label(req.body.hospital);
  } else if (pathname === '/api/leaves' || /^\/api\/admin\/leaves\/\d+$/.test(pathname)) {
    type = 'leave'; action = method === 'DELETE' ? 'حذف إجازة' : 'تسجيل إجازة'; target = idName('personnel', req.body.person_id, 'name');
  } else if (pathname === '/api/visitors' || /^\/api\/visitors\/\d+\/exit$/.test(pathname) || /^\/api\/admin\/visitors\/\d+$/.test(pathname)) {
    type = 'visitor'; action = method === 'DELETE' ? 'حذف زيارة' : pathname.endsWith('/exit') ? 'تسجيل خروج زائر' : 'تسجيل زيارة';
    target = label(req.body.name);
  } else if (pathname === '/api/squads' || /^\/api\/squads\/\d+\/members$/.test(pathname) || /^\/api\/admin\/squads\/\d+(?:\/members\/\d+)?$/.test(pathname) || /^\/api\/squads\/\d+\/attendance\/(?:in|out|absent|leave)$/.test(pathname)) {
    type = 'squad'; action = method === 'DELETE' ? 'حذف أو تعديل أفراد فرقة' : pathname.includes('/attendance/') ? (pathname.endsWith('/in') ? 'تسجيل حضور فرقة' : pathname.endsWith('/out') ? 'تسجيل انصراف فرقة' : pathname.endsWith('/absent') ? 'تسجيل غياب فرقة' : 'إلغاء غياب فرقة') : pathname.endsWith('/members') ? 'إضافة فرد لفرقة' : 'تسجيل فرقة';
    target = label(req.body.name) || idName('personnel', req.body.person_id, 'name');
  } else if (pathname === '/api/purchases' || /^\/api\/admin\/purchases\/\d+$/.test(pathname)) {
    type = 'purchase'; action = method === 'DELETE' ? 'حذف مشتريات' : 'تسجيل مشتريات'; target = idName('personnel', req.body.person_id, 'name');
  } else {
    return null;
  }
  return { type, summary: [action, target].filter(Boolean).join(' — ') };
}

app.use('/api', (req, res, next) => {
  if (['GET', 'HEAD', 'OPTIONS'].includes(req.method) || req.originalUrl.startsWith('/api/admin/notifications')) return next();
  const description = mutationDescription(req);
  if (description) res.on('finish', () => {
    if (res.statusCode < 200 || res.statusCode >= 300) return;
    try {
      const admins = db.prepare("SELECT id FROM users WHERE role='admin' AND active=1 AND id<>?").all(req.user.id);
      const insert = db.prepare('INSERT INTO notifications(recipient_id,actor_id,actor_name,type,summary,created_at) VALUES (?,?,?,?,?,?)');
      const write = db.transaction(() => admins.forEach(adminUser => insert.run(
        adminUser.id, req.user.id, req.user.full_name || req.user.username,
        description.type, description.summary, now()
      )));
      write();
    } catch (error) {
      console.error('Failed to create admin action notification:', error);
    }
  });
  next();
});

app.post('/api/auth/change-password', (req, res) => {
  const currentPassword = String(req.body.current_password || '');
  const newPassword = String(req.body.new_password || '');
  if (!verify(currentPassword, req.user.pass)) return res.status(400).json({ error: 'كلمة المرور الحالية غير صحيحة' });
  if (!okPw(newPassword)) return res.status(400).json({ error: 'كلمة المرور الجديدة يجب أن تكون ٨ أحرف على الأقل' });
  const token = cookie(req, 'sid');
  db.prepare('UPDATE users SET pass=? WHERE id=?').run(hash(newPassword), req.user.id);
  db.prepare('DELETE FROM sessions WHERE user_id=? AND token<>?').run(req.user.id, sha(token));
  res.json({ ok: true });
});

app.get('/api/admin/notifications', admin, (req, res) => {
  const limit = Math.min(Math.max(parseInt(req.query.limit, 10) || 50, 1), 200);
  res.json({
    unread: db.prepare('SELECT COUNT(*) AS count FROM notifications WHERE recipient_id=? AND read_at IS NULL').get(req.user.id).count,
    items: db.prepare('SELECT id,actor_name,type,summary,created_at,read_at FROM notifications WHERE recipient_id=? ORDER BY id DESC LIMIT ?').all(req.user.id, limit),
  });
});
app.post('/api/admin/notifications/:id/read', admin, (req, res) => {
  db.prepare('UPDATE notifications SET read_at=? WHERE id=? AND recipient_id=? AND read_at IS NULL').run(now(), req.params.id, req.user.id);
  res.json({ ok: true });
});
app.post('/api/admin/notifications/read-all', admin, (req, res) => {
  db.prepare('UPDATE notifications SET read_at=? WHERE recipient_id=? AND read_at IS NULL').run(now(), req.user.id);
  res.json({ ok: true });
});
app.get('/api/notifications', (req, res) => {
  const limit = Math.min(Math.max(parseInt(req.query.limit, 10) || 50, 1), 200);
  res.json({
    unread: db.prepare("SELECT COUNT(*) AS count FROM notifications WHERE recipient_id=? AND type IN ('member_message','correspondence_message') AND read_at IS NULL").get(req.user.id).count,
    items: db.prepare(`SELECT id,actor_name,type,summary,title,body,related_id,created_at,read_at
      FROM notifications WHERE recipient_id=? AND type IN ('member_message','correspondence_message') ORDER BY id DESC LIMIT ?`).all(req.user.id, limit),
  });
});
app.post('/api/notifications/:id/read', (req, res) => {
  db.prepare("UPDATE notifications SET read_at=? WHERE id=? AND recipient_id=? AND type IN ('member_message','correspondence_message') AND read_at IS NULL")
    .run(now(), req.params.id, req.user.id);
  res.json({ ok: true });
});
app.post('/api/notifications/read-all', (req, res) => {
  db.prepare("UPDATE notifications SET read_at=? WHERE recipient_id=? AND type IN ('member_message','correspondence_message') AND read_at IS NULL")
    .run(now(), req.user.id);
  res.json({ ok: true });
});
app.post('/api/admin/member-notifications', admin, (req, res) => {
  const title = String(req.body.title || '').trim(), body = String(req.body.body || '').trim();
  const recipient = String(req.body.recipient || '');
  if (!title || title.length > 150 || !body || body.length > 3000)
    return res.status(400).json({ error: 'عنوان الإشعار مطلوب (حتى ١٥٠ حرفاً) ونصه مطلوب (حتى ٣٠٠٠ حرف)' });
  if (recipient !== 'all' && (!/^\d+$/.test(recipient) || Number(recipient) < 1))
    return res.status(400).json({ error: 'اختر عضواً صحيحاً أو اختر كل الأعضاء' });
  const recipients = recipient === 'all'
    ? db.prepare("SELECT id FROM users WHERE role='operator' AND active=1 ORDER BY id").all()
    : db.prepare("SELECT id FROM users WHERE id=? AND role='operator' AND active=1").all(Number(recipient));
  if (!recipients.length) return res.status(400).json({ error: 'لا يوجد أعضاء نشطون لإرسال الإشعار إليهم' });
  const actorName = req.user.full_name || req.user.username, createdAt = now();
  const insert = db.prepare(`INSERT INTO notifications(recipient_id,actor_id,actor_name,type,summary,title,body,created_at)
    VALUES (?,?,?,'member_message',?,?,?,?)`);
  db.transaction(() => recipients.forEach(user =>
    insert.run(user.id, req.user.id, actorName, title, title, body, createdAt)
  ))();
  res.json({ ok: true, recipients: recipients.length });
});

app.get('/api/tasks', (req, res) => {
  res.json(db.prepare(`SELECT a.id AS assignment_id, t.title, t.details, t.created_at, t.created_by_name, a.confirmed_at
    FROM task_assignments a JOIN tasks t ON t.id=a.task_id
    WHERE a.recipient_id=? ORDER BY (a.confirmed_at IS NULL) DESC, t.id DESC`).all(req.user.id));
});
app.post('/api/tasks/:id/confirm', (req, res) => {
  const assignment = db.prepare('SELECT id, confirmed_at FROM task_assignments WHERE id=? AND recipient_id=?').get(req.params.id, req.user.id);
  if (!assignment) return res.status(404).json({ error: 'التكليف غير موجود' });
  if (assignment.confirmed_at) return res.json({ ok: true, confirmed_at: assignment.confirmed_at });
  const confirmedAt = now();
  db.prepare('UPDATE task_assignments SET confirmed_at=? WHERE id=? AND recipient_id=? AND confirmed_at IS NULL')
    .run(confirmedAt, assignment.id, req.user.id);
  res.json({ ok: true, confirmed_at: confirmedAt });
});
app.get('/api/admin/tasks', admin, (req, res) => {
  res.json(db.prepare(`SELECT t.id, t.title, t.details, t.created_at, t.created_by_name,
      a.id AS assignment_id, u.full_name, u.username, a.confirmed_at
    FROM tasks t JOIN task_assignments a ON a.task_id=t.id
    JOIN users u ON u.id=a.recipient_id
    ORDER BY t.id DESC, u.full_name COLLATE NOCASE`).all());
});
app.post('/api/admin/tasks', admin, (req, res) => {
  const title = String(req.body.title || '').trim();
  const details = String(req.body.details || '').trim();
  if (!title || title.length > 150 || details.length > 2000) {
    return res.status(400).json({ error: 'عنوان التكليف مطلوب (حتى ١٥٠ حرفاً)، والتفاصيل حتى ٢٠٠٠ حرف' });
  }
  const recipient = String(req.body.recipient || '');
  const recipientId = Number(recipient);
  if (recipient !== 'all' && (!Number.isInteger(recipientId) || recipientId < 1)) {
    return res.status(400).json({ error: 'اختر عضواً صحيحاً لإرسال التكليف' });
  }
  const users = recipient === 'all'
    ? db.prepare("SELECT id FROM users WHERE role='operator' AND active=1 ORDER BY id").all()
    : db.prepare("SELECT id FROM users WHERE id=? AND role='operator' AND active=1").all(recipientId);
  if (!users.length) return res.status(400).json({ error: 'اختر عضواً نشطاً أو أضف مستخدم تشغيل أولاً' });
  const createdAt = now(), createdByName = req.user.full_name || req.user.username;
  const createTask = db.transaction(() => {
    const taskId = db.prepare('INSERT INTO tasks(title,details,created_by,created_by_name,created_at) VALUES (?,?,?,?,?)')
      .run(title, details, req.user.id, createdByName, createdAt).lastInsertRowid;
    const assign = db.prepare('INSERT INTO task_assignments(task_id,recipient_id) VALUES (?,?)');
    users.forEach(user => assign.run(taskId, user.id));
    return taskId;
  });
  res.json({ ok: true, id: createTask(), recipients: users.length });
});

// ---------- correspondence archive ----------
const CORRESPONDENCE_TYPES = { incoming: 'واردة', outgoing: 'صادرة', internal: 'داخلية' };
const CORRESPONDENCE_STATUSES = { new: 'جديدة', in_progress: 'قيد الإجراء', replied: 'تم الرد', closed: 'مغلقة' };
const CORRESPONDENCE_MIMES = new Set([
  'application/pdf', 'image/jpeg', 'image/png', 'image/webp', 'text/plain',
  'application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
]);
const CORRESPONDENCE_FILE_TYPES = {
  '.pdf': 'application/pdf', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
  '.png': 'image/png', '.webp': 'image/webp', '.txt': 'text/plain',
  '.doc': 'application/msword',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xls': 'application/vnd.ms-excel',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};
const validCorrespondenceDate = value => {
  if (!dateOk(value)) return false;
  const parsed = new Date(`${value}T00:00:00Z`);
  return Number.isFinite(parsed.getTime()) && parsed.toISOString().slice(0, 10) === value;
};
function correspondenceFilter(req, res) {
  const from = String(req.query.from || ''), to = String(req.query.to || '');
  const kind = String(req.query.kind || ''), status = String(req.query.status || ''), section = String(req.query.section || '').trim();
  if ((from && !validCorrespondenceDate(from)) || (to && !validCorrespondenceDate(to)) || (from && to && from > to)) {
    res.status(400).json({ error: 'فترة التاريخ غير صحيحة' });
    return null;
  }
  if ((kind && !CORRESPONDENCE_TYPES[kind]) || (status && !CORRESPONDENCE_STATUSES[status])) {
    res.status(400).json({ error: 'نوع المكاتبة أو حالتها غير صحيحة' });
    return null;
  }
  if (section.length > 120) {
    res.status(400).json({ error: 'اسم القسم يجب ألا يتجاوز ١٢٠ حرفاً' });
    return null;
  }
  const clauses = [], values = [];
  if (from) { clauses.push('c.correspondence_date>=?'); values.push(from); }
  if (to) { clauses.push('c.correspondence_date<=?'); values.push(to); }
  if (kind) { clauses.push('c.kind=?'); values.push(kind); }
  if (status) { clauses.push('c.status=?'); values.push(status); }
  if (section) { clauses.push('c.responsible_section=?'); values.push(section); }
  if (req.user.role !== 'admin') {
    clauses.push('(c.created_by=? OR c.assigned_user_id=?)');
    values.push(req.user.id, req.user.id);
  }
  return { where: clauses.length ? `WHERE ${clauses.join(' AND ')}` : '', values, from, to, kind, status, section };
}
const correspondenceListSql = `SELECT c.*, u.full_name AS assigned_name, u.username AS assigned_username,
  (SELECT COUNT(*) FROM correspondence_attachments a WHERE a.correspondence_id=c.id) AS attachment_count,
  (SELECT COUNT(*) FROM correspondence_messages m WHERE m.correspondence_id=c.id) AS message_count
  FROM correspondence c LEFT JOIN users u ON u.id=c.assigned_user_id`;
app.get('/api/correspondence', (req, res) => {
  const filter = correspondenceFilter(req, res);
  if (!filter) return;
  res.json(db.prepare(`${correspondenceListSql} ${filter.where} ORDER BY c.correspondence_date DESC,c.id DESC`).all(...filter.values));
});
app.post('/api/correspondence', (req, res) => {
  const b = req.body;
  const reference = String(b.reference_number || '').trim();
  const kind = String(b.kind || ''), correspondenceDate = String(b.correspondence_date || '');
  const sender = String(b.sender || '').trim(), recipient = String(b.recipient || '').trim();
  const responsibleSection = String(b.responsible_section || '').trim();
  const subject = String(b.subject || '').trim(), summary = String(b.summary || '').trim();
  const dueDate = String(b.due_date || '');
  if (!CORRESPONDENCE_TYPES[kind] || !validCorrespondenceDate(correspondenceDate) ||
      (dueDate && !validCorrespondenceDate(dueDate)) || !subject ||
      reference.length > 80 || sender.length > 160 || recipient.length > 160 || responsibleSection.length > 120 || subject.length > 200 || summary.length > 3000) {
    return res.status(400).json({ error: 'أدخل نوع المكاتبة وتاريخها وموضوعاً صحيحاً، وتأكد من أطوال الحقول' });
  }
  if (dueDate && dueDate < correspondenceDate) return res.status(400).json({ error: 'موعد المتابعة لا يمكن أن يسبق تاريخ المكاتبة' });
  const requestedAssignee = b.assigned_user_id == null || b.assigned_user_id === '' ? null : Number(b.assigned_user_id);
  let assignedUserId = requestedAssignee;
  if (req.user.role !== 'admin') {
    if (requestedAssignee != null && requestedAssignee !== req.user.id)
      return res.status(403).json({ error: 'يمكن إسناد المكاتبة إلى حسابك فقط' });
    assignedUserId = req.user.id;
  } else if (requestedAssignee != null &&
      (!Number.isInteger(requestedAssignee) || !db.prepare('SELECT 1 FROM users WHERE id=? AND active=1').get(requestedAssignee))) {
    return res.status(400).json({ error: 'اختر مستخدماً نشطاً للمسؤولية عن المكاتبة' });
  }
  const attachments = b.attachments == null ? [] : b.attachments;
  if (!Array.isArray(attachments) || attachments.length > 3) return res.status(400).json({ error: 'يمكن إرفاق ٣ ملفات كحد أقصى' });
  const files = [];
  let totalSize = 0;
  for (const item of attachments) {
    const filename = path.basename(String(item?.filename || '').replace(/\\/g, '/')).replace(/[\u0000-\u001f\u007f]/g, '').trim();
    const mimeType = String(item?.mime_type || '');
    const match = typeof item?.data === 'string' && item.data.match(/^data:([^;,]+);base64,([A-Za-z0-9+/]*={0,2})$/);
    if (!filename || filename.length > 180 || CORRESPONDENCE_FILE_TYPES[path.extname(filename).toLowerCase()] !== mimeType ||
        !CORRESPONDENCE_MIMES.has(mimeType) || !match || match[1] !== mimeType) {
      return res.status(400).json({ error: 'أحد المرفقات غير صالح؛ الصيغ المدعومة PDF وWord وExcel والصور وTXT' });
    }
    const encoded = match[2], content = Buffer.from(encoded, 'base64');
    if (!content.length || content.toString('base64') !== encoded || content.length > 5 * 1024 * 1024) {
      return res.status(400).json({ error: 'حجم المرفق يجب ألا يتجاوز ٥ ميجابايت' });
    }
    totalSize += content.length;
    if (totalSize > 15 * 1024 * 1024) return res.status(400).json({ error: 'إجمالي أحجام المرفقات يجب ألا يتجاوز ١٥ ميجابايت' });
    files.push({ filename, mimeType, content });
  }
  const createdAt = now(), createdByName = req.user.full_name || req.user.username;
  try {
    const id = db.transaction(() => {
      const result = db.prepare(`INSERT INTO correspondence
        (reference_number,kind,correspondence_date,sender,recipient,responsible_section,subject,summary,assigned_user_id,status,due_date,created_by,created_by_name,created_at,updated_at,updated_by)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`).run(
        reference, kind, correspondenceDate, sender, recipient, responsibleSection, subject, summary, assignedUserId,
        'new', dueDate || null, req.user.id, createdByName, createdAt, createdAt, req.user.username
      );
      const insertAttachment = db.prepare(`INSERT INTO correspondence_attachments(correspondence_id,filename,mime_type,size,content)
        VALUES (?,?,?,?,?)`);
      files.forEach(file => insertAttachment.run(result.lastInsertRowid, file.filename, file.mimeType, file.content.length, file.content));
      return result.lastInsertRowid;
    })();
    res.json({ ok: true, id });
  } catch (error) {
    if (error.code === 'SQLITE_CONSTRAINT_UNIQUE') return res.status(409).json({ error: 'رقم المكاتبة مسجل من قبل' });
    throw error;
  }
});
app.put('/api/correspondence/:id/status', (req, res) => {
  const status = String(req.body.status || '');
  if (!CORRESPONDENCE_STATUSES[status]) return res.status(400).json({ error: 'اختر حالة مكاتبة صحيحة' });
  const record = db.prepare('SELECT id,created_by,assigned_user_id FROM correspondence WHERE id=?').get(req.params.id);
  if (!record || (req.user.role !== 'admin' && record.created_by !== req.user.id && record.assigned_user_id !== req.user.id))
    return res.status(404).json({ error: 'المكاتبة غير موجودة أو ليس لديك صلاحية الاطلاع عليها' });
  db.prepare('UPDATE correspondence SET status=?,updated_at=?,updated_by=? WHERE id=?')
    .run(status, now(), req.user.username, record.id);
  res.json({ ok: true });
});
app.get('/api/correspondence/:id', (req, res) => {
  const record = db.prepare(`${correspondenceListSql} WHERE c.id=?`).get(req.params.id);
  if (!record || (req.user.role !== 'admin' && record.created_by !== req.user.id && record.assigned_user_id !== req.user.id))
    return res.status(404).json({ error: 'المكاتبة غير موجودة أو ليس لديك صلاحية الاطلاع عليها' });
  const attachments = db.prepare('SELECT id,filename,mime_type,size FROM correspondence_attachments WHERE correspondence_id=? ORDER BY id')
    .all(record.id);
  res.json({ ...record, attachments });
});
app.get('/api/correspondence/:id/messages', (req, res) => {
  const record = db.prepare('SELECT id,created_by,assigned_user_id FROM correspondence WHERE id=?').get(req.params.id);
  if (!record || (req.user.role !== 'admin' && record.created_by !== req.user.id && record.assigned_user_id !== req.user.id))
    return res.status(404).json({ error: 'المكاتبة غير موجودة أو ليس لديك صلاحية الاطلاع عليها' });
  const messages = db.prepare(`SELECT id,author_name,body,created_at FROM correspondence_messages
    WHERE correspondence_id=? ORDER BY id ASC LIMIT 500`).all(record.id);
  const attachments = db.prepare(`SELECT id,message_id,filename,mime_type,size FROM correspondence_message_attachments
    WHERE message_id IN (SELECT id FROM correspondence_messages WHERE correspondence_id=?)
    ORDER BY message_id,id`).all(record.id);
  const byMessage = new Map(messages.map(message => [message.id, []]));
  attachments.forEach(attachment => byMessage.get(attachment.message_id)?.push(attachment));
  res.json(messages.map(message => ({ ...message, attachments: byMessage.get(message.id) || [] })));
});
app.post('/api/correspondence/:id/messages', (req, res) => {
  const record = db.prepare('SELECT id,created_by,assigned_user_id FROM correspondence WHERE id=?').get(req.params.id);
  if (!record || (req.user.role !== 'admin' && record.created_by !== req.user.id && record.assigned_user_id !== req.user.id))
    return res.status(404).json({ error: 'المكاتبة غير موجودة أو ليس لديك صلاحية الاطلاع عليها' });
  const body = String(req.body.body || '').trim();
  if (body.length > 3000) return res.status(400).json({ error: 'رسالة المتابعة يجب ألا تتجاوز ٣٠٠٠ حرف' });
  const attachments = req.body.attachments == null ? [] : req.body.attachments;
  if (!Array.isArray(attachments) || attachments.length > 5)
    return res.status(400).json({ error: 'يمكن إرفاق ٥ ملفات كحد أقصى في كل رسالة' });
  const files = [];
  let totalSize = 0;
  for (const item of attachments) {
    const filename = path.basename(String(item?.filename || '').replace(/\\/g, '/')).replace(/[\u0000-\u001f\u007f]/g, '').trim();
    const mimeType = String(item?.mime_type || '');
    const match = typeof item?.data === 'string' && item.data.match(/^data:([^;,]+);base64,([A-Za-z0-9+/]*={0,2})$/);
    if (!filename || filename.length > 180 || CORRESPONDENCE_FILE_TYPES[path.extname(filename).toLowerCase()] !== mimeType ||
        !CORRESPONDENCE_MIMES.has(mimeType) || !match || match[1] !== mimeType)
      return res.status(400).json({ error: 'أحد المرفقات غير صالح؛ الصيغ المدعومة PDF وWord وExcel والصور وTXT' });
    const encoded = match[2], content = Buffer.from(encoded, 'base64');
    if (!content.length || content.toString('base64') !== encoded || content.length > 5 * 1024 * 1024)
      return res.status(400).json({ error: 'حجم كل ملف يجب ألا يتجاوز ٥ ميجابايت' });
    totalSize += content.length;
    if (totalSize > 15 * 1024 * 1024) return res.status(400).json({ error: 'إجمالي أحجام الملفات يجب ألا يتجاوز ١٥ ميجابايت' });
    files.push({ filename, mimeType, content });
  }
  if (!body && !files.length) return res.status(400).json({ error: 'اكتب رسالة متابعة أو أرفق نسخة من الملف' });
  const createdAt = now(), authorName = req.user.full_name || req.user.username;
  const subject = db.prepare('SELECT subject FROM correspondence WHERE id=?').get(record.id).subject;
  const messageId = db.transaction(() => {
    const message = db.prepare(`INSERT INTO correspondence_messages
      (correspondence_id,user_id,author_name,body,created_at) VALUES (?,?,?,?,?)`)
      .run(record.id, req.user.id, authorName, body, createdAt);
    const insertAttachment = db.prepare(`INSERT INTO correspondence_message_attachments
      (message_id,filename,mime_type,size,content) VALUES (?,?,?,?,?)`);
    files.forEach(file => insertAttachment.run(message.lastInsertRowid, file.filename, file.mimeType, file.content.length, file.content));
    db.prepare('UPDATE correspondence SET updated_at=?,updated_by=? WHERE id=?')
      .run(createdAt, req.user.username, record.id);
    const recipients = new Set([record.created_by, record.assigned_user_id].filter(id => id && id !== req.user.id));
    const notify = db.prepare(`INSERT INTO notifications(recipient_id,actor_id,actor_name,type,summary,title,body,related_id,created_at)
      VALUES (?,?,?,'correspondence_message',?,?,?,?,?)`);
    recipients.forEach(recipientId => notify.run(
      recipientId, req.user.id, authorName, `رد جديد على: ${subject}`,
      `متابعة المكاتبة: ${subject}`, body || `أُرفقت ${files.length} نسخة جديدة للمراجعة.`, record.id, createdAt
    ));
    return message.lastInsertRowid;
  })();
  res.json({ ok: true, id: messageId });
});
app.get('/api/correspondence/:id/messages/:messageId/attachments/:attachmentId', (req, res) => {
  const record = db.prepare('SELECT id,created_by,assigned_user_id FROM correspondence WHERE id=?').get(req.params.id);
  if (!record || (req.user.role !== 'admin' && record.created_by !== req.user.id && record.assigned_user_id !== req.user.id))
    return res.status(404).json({ error: 'المرفق غير موجود' });
  const attachment = db.prepare(`SELECT a.filename,a.mime_type,a.content FROM correspondence_message_attachments a
    JOIN correspondence_messages m ON m.id=a.message_id
    WHERE a.id=? AND m.id=? AND m.correspondence_id=?`).get(req.params.attachmentId, req.params.messageId, record.id);
  if (!attachment) return res.status(404).json({ error: 'المرفق غير موجود' });
  res.set({
    'Content-Type': attachment.mime_type,
    'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(attachment.filename)}`,
    'Cache-Control': 'private, no-store',
    'X-Content-Type-Options': 'nosniff',
    'Content-Security-Policy': "default-src 'none'; sandbox",
  });
  res.send(attachment.content);
});
app.get('/api/correspondence/:id/attachments/:attachmentId', (req, res) => {
  const record = db.prepare('SELECT id,created_by,assigned_user_id FROM correspondence WHERE id=?').get(req.params.id);
  if (!record || (req.user.role !== 'admin' && record.created_by !== req.user.id && record.assigned_user_id !== req.user.id))
    return res.status(404).json({ error: 'المرفق غير موجود' });
  const attachment = db.prepare('SELECT filename,mime_type,content FROM correspondence_attachments WHERE id=? AND correspondence_id=?')
    .get(req.params.attachmentId, record.id);
  if (!attachment) return res.status(404).json({ error: 'المرفق غير موجود' });
  res.set({
    'Content-Type': attachment.mime_type,
    'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(attachment.filename)}`,
    'Cache-Control': 'private, no-store',
    'X-Content-Type-Options': 'nosniff',
    'Content-Security-Policy': "default-src 'none'; sandbox",
  });
  res.send(attachment.content);
});
app.delete('/api/admin/correspondence/:id', admin, (req, res) => {
  db.prepare('DELETE FROM correspondence WHERE id=?').run(req.params.id);
  res.json({ ok: true });
});
app.get('/api/correspondence.xlsx', async (req, res, next) => {
  const filter = correspondenceFilter(req, res);
  if (!filter) return;
  try {
    const rows = db.prepare(`${correspondenceListSql} ${filter.where} ORDER BY c.correspondence_date DESC,c.id DESC`).all(...filter.values);
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'التحكم في البوابات';
    const sheet = workbook.addWorksheet('المكاتبات', { views: [{ rightToLeft: true, state: 'frozen', ySplit: 1 }] });
    sheet.columns = [
      { header: 'رقم المكاتبة', key: 'reference_number', width: 20 },
      { header: 'النوع', key: 'kind', width: 14 },
      { header: 'التاريخ', key: 'correspondence_date', width: 16 },
      { header: 'الجهة المرسلة', key: 'sender', width: 24 },
      { header: 'الجهة الموجهة إليها', key: 'recipient', width: 24 },
      { header: 'القسم المسؤول', key: 'responsible_section', width: 20 },
      { header: 'الموضوع', key: 'subject', width: 36 },
      { header: 'الملخص', key: 'summary', width: 48 },
      { header: 'المسؤول', key: 'assigned_name', width: 24 },
      { header: 'الحالة', key: 'status', width: 18 },
      { header: 'موعد المتابعة', key: 'due_date', width: 16 },
    ];
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5BC4' } };
    rows.forEach(row => sheet.addRow({
      ...row, kind: CORRESPONDENCE_TYPES[row.kind], status: CORRESPONDENCE_STATUSES[row.status],
      assigned_name: row.assigned_name || row.assigned_username || '',
    }));
    sheet.autoFilter = { from: 'A1', to: `K${rows.length + 1}` };
    const suffix = filter.from && filter.to ? `${filter.from}-${filter.to}` : 'all-dates';
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="gate-correspondence-${suffix}.xlsx"`,
      'Cache-Control': 'private, no-store',
    });
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    next(error);
  }
});

app.get('/api/admin/backups', admin, async (req, res) => {
  try {
    res.json({ backups: await backupFiles(), retention: BACKUP_RETENTION });
  } catch (error) {
    console.error('Unable to list database backups:', error);
    res.status(500).json({ error: 'تعذّر تحميل النسخ الاحتياطية' });
  }
});
app.post('/api/admin/backups', admin, async (req, res) => {
  try {
    const file = await createBackup('manual');
    res.json({ ok: true, name: path.basename(file) });
  } catch (error) {
    console.error('Unable to create database backup:', error);
    res.status(500).json({ error: 'تعذّر إنشاء النسخة الاحتياطية' });
  }
});
app.post('/api/admin/backups/:name/restore', admin, async (req, res) => {
  const name = String(req.params.name || '');
  if (!BACKUP_NAME.test(name)) return res.status(400).json({ error: 'اسم النسخة الاحتياطية غير صحيح' });
  const file = backupPath(name);
  if (path.resolve(file) === DB_PATH) return res.status(400).json({ error: 'لا يمكن استعادة قاعدة البيانات من نفسها' });
  if (backupBusy) return res.status(409).json({ error: 'هناك عملية نسخ احتياطي أو استعادة قيد التنفيذ' });
  backupBusy = true;
  let source;
  try {
    await fsp.access(file, fs.constants.R_OK);
    await validateBackup(file);
    await writeBackup('pre-restore', [name]);
    source = new Database(file, { readonly: true, fileMustExist: true });
    await source.backup(DB_PATH);
    source.close();
    source = null;
    await pruneBackups().catch(error => {
      console.error('Unable to prune old database backups after restore:', error);
    });
    res.json({ ok: true, name });
  } catch (error) {
    if (source?.open) source.close();
    console.error('Unable to restore database backup:', error);
    res.status(500).json({ error: 'تعذّرت استعادة النسخة؛ تم الاحتفاظ بحفظ ما قبل الاستعادة إن أمكن' });
  } finally {
    backupBusy = false;
  }
});

// ---------- lookups (dropdown lists) ----------
app.get('/api/lookups', (q, res) => {
  const out = {}; Object.keys(SEED).forEach(k => out[k] = []);
  db.prepare('SELECT * FROM lookups ORDER BY id').all().forEach(r => out[r.kind]?.push({ id: r.id, name: r.name }));
  out.rank = RANKS.map(r => ({ id: r.ord, name: r.name, cat: r.cat }));
  res.json(out);
});
app.post('/api/admin/lookups', admin, (req, res) => {
  const name = String(req.body.name || '').trim();
  if (!SEED[req.body.kind] || !name) return res.status(400).json({ error: 'بيانات غير صالحة' });
  try { res.json({ id: db.prepare('INSERT INTO lookups(kind,name) VALUES (?,?)').run(req.body.kind, name).lastInsertRowid }); }
  catch { res.status(409).json({ error: 'العنصر موجود مسبقاً' }); }
});
app.delete('/api/admin/lookups/:id', admin, (req, res) => { db.prepare('DELETE FROM lookups WHERE id=?').run(req.params.id); res.json({ ok: true }); });

// ---------- admin CRUD: personnel / vehicles / items ----------
const ENT = {
  personnel: { cols: ['name', 'mid', 'rank', 'retired', 'unit', 'phone'], req: ['name', 'mid', 'rank'], lk: { unit: 'unit' } },
  vehicles: { cols: ['plate', 'vehicle', 'purpose'], req: ['plate', 'vehicle', 'purpose'], lk: { vehicle: 'vehicle', purpose: 'purpose' } },
};
function clean(e, b, isNew) {
  const E = ENT[e], o = {};
  for (const c of E.cols) {
    if (!isNew && c === 'stock') continue; // stock changes only through stock movements
    let v = b[c];
    if (c === 'stock' || c === 'reorder') v = Math.max(0, parseInt(v, 10) || 0);
    else if (c === 'driver_id' || c === 'companion_id') v = v ? +v : null;
    else if (c === 'retired') {
      if ([true, 1, '1'].includes(v)) v = 1;
      else if ([false, 0, '0', undefined, null, ''].includes(v)) v = 0;
      else return { error: 'حدد ما إذا كان الفرد من الضباط المتقاعدين' };
    }
    else v = String(v ?? '').trim();
    if (E.req.includes(c) && !v) return { error: 'يرجى ملء الحقول المطلوبة' };
    if (E.lk[c] && v && !db.prepare('SELECT 1 FROM lookups WHERE kind=? AND name=?').get(E.lk[c], v)) return { error: 'القيمة غير موجودة في القوائم' };
    if (c === 'rank' && !RK[v]) return { error: 'اختر رتبة صحيحة' };
    if (c === 'driver_id' && v && !db.prepare('SELECT 1 FROM personnel WHERE id=?').get(v)) return { error: 'السائق غير موجود' };
    if (c === 'companion_id' && v && !db.prepare('SELECT 1 FROM personnel WHERE id=?').get(v)) return { error: 'المرافق غير موجود' };
    o[c] = v;
  }
  if (e === 'personnel' && o.retired && RK[o.rank]?.cat !== 'ضابط') {
    return { error: 'الضباط المتقاعدون يجب أن تكون رتبتهم من فئة الضباط' };
  }
  if (o.driver_id && o.companion_id && o.driver_id === o.companion_id) return { error: 'اختر فرداً مختلفاً كمرافق للسائق' };
  return { o };
}
app.post('/api/admin/:e(personnel|vehicles)', admin, (req, res) => {
  const { o, error } = clean(req.params.e, req.body, true); if (error) return res.status(400).json({ error });
  const ks = Object.keys(o);
  try { res.json({ id: db.prepare(`INSERT INTO ${req.params.e}(${ks}) VALUES (${ks.map(() => '?')})`).run(...Object.values(o)).lastInsertRowid }); }
  catch { res.status(409).json({ error: 'السجل موجود مسبقاً (رقم عسكري أو لوحة مكررة)' }); }
});
app.put('/api/admin/:e(personnel|vehicles)/:id', admin, (req, res) => {
  const { o, error } = clean(req.params.e, req.body, false); if (error) return res.status(400).json({ error });
  try { db.prepare(`UPDATE ${req.params.e} SET ${Object.keys(o).map(k => k + '=?')} WHERE id=?`).run(...Object.values(o), req.params.id); res.json({ ok: true }); }
  catch { res.status(409).json({ error: 'السجل موجود مسبقاً (رقم عسكري أو لوحة مكررة)' }); }
});
app.delete('/api/admin/:e(personnel|vehicles)/:id', admin, (req, res) => {
  db.prepare(`DELETE FROM ${req.params.e} WHERE id=?`).run(req.params.id); res.json({ ok: true });
});

// ---------- admin: users ----------
app.get('/api/admin/users', admin, (q, res) => res.json(db.prepare('SELECT id,username,full_name,role,active FROM users ORDER BY id').all()));
app.post('/api/admin/users', admin, (req, res) => {
  const u = String(req.body.username || '').trim().toLowerCase(), role = req.body.role === 'admin' ? 'admin' : 'operator';
  if (!okUser(u)) return res.status(400).json({ error: 'اسم المستخدم: ٣–٣٢ حرفاً إنجليزياً أو أرقاماً' });
  if (!okPw(req.body.password)) return res.status(400).json({ error: 'كلمة المرور ٨ أحرف على الأقل' });
  try { db.prepare('INSERT INTO users(username,pass,full_name,role) VALUES (?,?,?,?)').run(u, hash(req.body.password), String(req.body.full_name || '').trim(), role); res.json({ ok: true }); }
  catch { res.status(409).json({ error: 'اسم المستخدم مستخدم مسبقاً' }); }
});
app.put('/api/admin/users/:id', admin, (req, res) => {
  const id = +req.params.id, b = req.body, t = db.prepare('SELECT * FROM users WHERE id=?').get(id);
  if (!t) return res.status(404).json({ error: 'غير موجود' });
  if (b.role != null && !['admin', 'operator'].includes(b.role)) return res.status(400).json({ error: 'الصلاحية غير صحيحة' });
  if (b.active != null && ![0, 1, true, false].includes(b.active)) return res.status(400).json({ error: 'حالة الحساب غير صحيحة' });
  const role = b.role == null ? t.role : b.role, active = b.active == null ? t.active : (b.active ? 1 : 0);
  if (id === req.user.id && (role !== 'admin' || !active)) return res.status(400).json({ error: 'لا يمكنك تعطيل نفسك أو إزالة صلاحية المدير عنك' });
  if (b.password && !okPw(b.password)) return res.status(400).json({ error: 'كلمة المرور ٨ أحرف على الأقل' });
  if (t.role === 'admin' && t.active && (role !== 'admin' || !active) &&
      !db.prepare("SELECT 1 FROM users WHERE role='admin' AND active=1 AND id<>?").get(id))
    return res.status(409).json({ error: 'لا يمكن تعطيل أو تخفيض صلاحية آخر مدير نشط' });
  db.prepare('UPDATE users SET role=?, active=?, pass=? WHERE id=?').run(role, active, b.password ? hash(b.password) : t.pass, id);
  if (b.password || !active || role !== t.role) db.prepare('DELETE FROM sessions WHERE user_id=?').run(id);
  res.json({ ok: true });
});

// ---------- operations ----------
const kinds = { personnel: { table: 'personnel', label: 'name', kind: 'personnel' }, vehicles: { table: 'vehicles', label: 'plate', kind: 'vehicle' } };
const gateOf = g => (db.prepare("SELECT name FROM lookups WHERE kind='gate' AND name=?").get(g) || db.prepare("SELECT name FROM lookups WHERE kind='gate' ORDER BY id").get() || { name: '—' }).name;

app.get('/api/vehicles', (q, res) => {
  const assignmentDate = dateOk(q.query.date) ? q.query.date : now().slice(0, 10);
  const registered = db.prepare(`SELECT v.*, v.vehicle AS type, a.driver_id, a.companion_id,
      p.name AS driver, p.rank AS rank, c.name AS companion, c.rank AS companion_rank,
      CASE WHEN a.driver_id IS NOT NULL AND a.companion_id IS NOT NULL THEN 1 ELSE 0 END AS crew_assigned,
      0 AS is_manual
    FROM vehicles v
    LEFT JOIN vehicle_assignments a ON a.vehicle_id=v.id AND a.date=?
    LEFT JOIN personnel p ON p.id=a.driver_id
    LEFT JOIN personnel c ON c.id=a.companion_id
    ORDER BY v.id`).all(assignmentDate).map(vehicle => ({
      ...vehicle,
      passengers: db.prepare(`SELECT p.id, p.name, p.rank FROM vehicle_assignment_passengers ap
        JOIN personnel p ON p.id=ap.person_id
        WHERE ap.date=? AND ap.vehicle_id=? ORDER BY p.name`).all(assignmentDate, vehicle.id),
    }));
  const manual = db.prepare(`SELECT m.ref AS plate,
      (SELECT details FROM movements WHERE kind='vehicle' AND ref=m.ref ORDER BY id DESC LIMIT 1) AS details,
      (SELECT destination FROM movements WHERE kind='vehicle' AND ref=m.ref ORDER BY id DESC LIMIT 1) AS last_destination,
      (SELECT type FROM movements WHERE kind='vehicle' AND ref=m.ref ORDER BY id DESC LIMIT 1) AS last_type
    FROM movements m
    WHERE m.kind='vehicle' AND NOT EXISTS (SELECT 1 FROM vehicles v WHERE v.plate=m.ref)
    GROUP BY m.ref ORDER BY MAX(m.id) DESC`).all().map(m => ({
      ...m,
      id: null,
      type: m.details || 'مركبة غير مسجلة',
      driver: null,
      rank: null,
      companion: null,
      companion_rank: null,
      purpose: m.last_destination || '',
      inside: m.last_type === 'IN' ? 1 : 0,
      is_completed: m.last_type === 'OUT',
      is_manual: 1,
    }));
  res.json([...registered, ...manual]);
});

app.post('/api/vehicles/:id/assignment', (req, res) => {
  const date = now().slice(0, 10);
  const vehicle = db.prepare('SELECT id FROM vehicles WHERE id=?').get(req.params.id);
  if (!vehicle) return res.status(404).json({ error: 'المركبة غير موجودة' });
  if (req.body.date && req.body.date !== date) return res.status(400).json({ error: 'يمكن تسجيل طاقم اليوم فقط' });
  const driverId = Number(req.body.driver_id), companionId = Number(req.body.companion_id);
  const driver = Number.isInteger(driverId) && db.prepare('SELECT id, rank FROM personnel WHERE id=?').get(driverId);
  const companion = Number.isInteger(companionId) && db.prepare('SELECT id FROM personnel WHERE id=?').get(companionId);
  if (!driver || !companion) return res.status(400).json({ error: 'اختر السائق والمرافق من قائمة الأفراد' });
  if (driver.id === companion.id) return res.status(400).json({ error: 'اختر فرداً مختلفاً كمرافق للسائق' });
  if (!['ncos', 'soldiers'].includes(RK[driver.rank]?.grp)) return res.status(400).json({ error: 'يجب اختيار السائق من صف الضباط أو الجنود' });
  const passengerIds = req.body.passenger_ids;
  if (passengerIds != null && (!Array.isArray(passengerIds) || passengerIds.length > 100))
    return res.status(400).json({ error: 'قائمة الأفراد الإضافيين غير صحيحة' });
  const uniquePassengerIds = [...new Set((passengerIds || []).map(Number))];
  if (uniquePassengerIds.some(personId => !Number.isInteger(personId) || personId < 1 || personId === driver.id || personId === companion.id))
    return res.status(400).json({ error: 'الأفراد الإضافيون يجب أن يكونوا أفراداً مختلفين عن السائق والمرافق' });
  if (uniquePassengerIds.length) {
    const placeholders = uniquePassengerIds.map(() => '?').join(',');
    const validCount = db.prepare(`SELECT COUNT(*) AS count FROM personnel WHERE id IN (${placeholders})`).get(...uniquePassengerIds).count;
    if (validCount !== uniquePassengerIds.length) return res.status(400).json({ error: 'أحد الأفراد الإضافيين غير موجود' });
  }
  db.transaction(() => {
    db.prepare(`INSERT INTO vehicle_assignments(date,vehicle_id,driver_id,companion_id,by,updated_at)
      VALUES (?,?,?,?,?,?) ON CONFLICT(date,vehicle_id) DO UPDATE SET
      driver_id=excluded.driver_id, companion_id=excluded.companion_id, by=excluded.by, updated_at=excluded.updated_at`)
      .run(date, vehicle.id, driver.id, companion.id, req.user.username, now());
    db.prepare('DELETE FROM vehicle_assignment_passengers WHERE date=? AND vehicle_id=?').run(date, vehicle.id);
    const insertPassenger = db.prepare('INSERT INTO vehicle_assignment_passengers(date,vehicle_id,person_id) VALUES (?,?,?)');
    uniquePassengerIds.forEach(personId => insertPassenger.run(date, vehicle.id, personId));
  })();
  res.json({ ok: true });
});

const vehicleHistory = plate => db.prepare(`SELECT * FROM movements
  WHERE kind='vehicle' AND (ref=? OR ref_id=(SELECT id FROM vehicles WHERE plate=?))
  ORDER BY id DESC LIMIT 500`).all(plate, plate);
app.get('/api/vehicles/history', (req, res) => {
  const plate = String(req.query.plate || '').trim();
  if (!plate) return res.status(400).json({ error: 'أدخل رقم اللوحة' });
  res.json(vehicleHistory(plate));
});
app.get('/api/vehicles/:id/history', (req, res) => {
  const v = db.prepare('SELECT plate FROM vehicles WHERE id=?').get(req.params.id);
  if (!v) return res.status(404).json({ error: 'المركبة غير موجودة' });
  res.json(vehicleHistory(v.plate));
});
app.post('/api/vehicles/manual-movement', (req, res) => {
  const b = req.body, plate = String(b.plate || '').trim(), type = String(b.type || '');
  const destination = String(b.destination || '').trim();
  if (!plate || !['IN', 'OUT'].includes(type)) return res.status(400).json({ error: 'أدخل رقم اللوحة ونوع الحركة' });
  if (db.prepare('SELECT 1 FROM vehicles WHERE plate=?').get(plate)) return res.status(409).json({ error: 'المركبة مسجلة بالفعل؛ استخدم سجل المركبات لتحديث حالتها' });
  const previous = db.prepare("SELECT type FROM movements WHERE kind='vehicle' AND ref=? ORDER BY id DESC LIMIT 1").get(plate);
  if (!previous && type !== 'IN') return res.status(409).json({ error: 'يجب تسجيل دخول المركبة أولاً' });
  if (previous?.type === 'OUT') return res.status(409).json({ error: 'تم تسجيل خروج المركبة؛ الزيارة منتهية ولا يمكن تسجيل حركة أخرى لها' });
  if (previous?.type === type) return res.status(409).json({ error: 'المركبة مسجلة بالداخل بالفعل' });
  const vehicleType = String(b.vehicle_type || '').trim(), driver = String(b.driver || '').trim(), rank = String(b.rank || '').trim();
  if (rank && !RK[rank]) return res.status(400).json({ error: 'اختر رتبة صحيحة' });
  const details = [vehicleType && `نوع المركبة: ${vehicleType}`, driver && `السائق: ${driver}`, rank && `الرتبة: ${rank}`].filter(Boolean).join(' • ');
  db.prepare('INSERT INTO movements(kind,ref_id,ref,type,ts,gate,destination,details,by) VALUES (?,?,?,?,?,?,?,?,?)')
    .run('vehicle', null, plate, type, now(), gateOf(b.gate), type === 'OUT' ? destination || null : null, details || null, req.user.username);
  res.json({ ok: true });
});

app.post('/api/:k(personnel|vehicles)/:id/toggle', (req, res) => {
  const k = kinds[req.params.k];
  const out = db.transaction(() => {
    const r = db.prepare(`SELECT * FROM ${k.table} WHERE id=?`).get(req.params.id);
    if (!r) return null;
    const ts = now(), inside = r.inside ? 0 : 1, type = inside ? 'IN' : 'OUT';
    const destination = req.params.k === 'vehicles' && !inside ? String(req.body.destination || '').trim() || null : null;
    if (req.params.k === 'personnel' && inside) return { error: 'سجّل دخول الفرد من زر «إضافة فرد»' };
    if (req.params.k === 'personnel' && db.prepare('SELECT 1 FROM movements WHERE kind=? AND ref_id=? AND type=? AND ts LIKE ? LIMIT 1')
      .get(k.kind, r.id, type, ts.slice(0, 10) + '%')) {
      return { error: type === 'IN'
        ? 'تم تسجيل دخول الفرد اليوم بالفعل؛ سجّل أي خروج إضافي كمأمورية'
        : 'تم تسجيل خروج الفرد اليوم بالفعل؛ سجّل أي خروج إضافي كمأمورية' };
    }
    let details = null;
    if (req.params.k === 'vehicles') {
      const assignment = db.prepare(`SELECT p.name AS driver, p.rank AS driver_rank, c.name AS companion, c.rank AS companion_rank
        FROM vehicle_assignments a JOIN personnel p ON p.id=a.driver_id
        JOIN personnel c ON c.id=a.companion_id
        WHERE a.vehicle_id=? AND a.date=?`).get(r.id, ts.slice(0, 10));
      if (!assignment) return { error: 'حدد سائق ومرافق المركبة لهذا اليوم أولاً' };
      const passengers = db.prepare(`SELECT p.name, p.rank FROM vehicle_assignment_passengers ap
        JOIN personnel p ON p.id=ap.person_id WHERE ap.vehicle_id=? AND ap.date=? ORDER BY p.name`).all(r.id, ts.slice(0, 10));
      details = [
        `السائق: ${assignment.driver} — ${assignment.driver_rank}`,
        `المرافق: ${assignment.companion} — ${assignment.companion_rank}`,
        passengers.length ? `أفراد إضافيون: ${passengers.map(person => `${person.name} — ${person.rank}`).join('، ')}` : '',
      ].filter(Boolean).join(' • ');
    }
    db.prepare(`UPDATE ${k.table} SET inside=? WHERE id=?`).run(inside, r.id);
    db.prepare('INSERT INTO movements(kind,ref_id,ref,type,ts,gate,destination,details,by) VALUES (?,?,?,?,?,?,?,?,?)')
      .run(k.kind, r.id, r[k.label], type, ts, gateOf(req.body.gate), destination, details, req.user.username);
    return { ...r, inside, destination, details };
  })();
  if (!out) return res.status(404).json({ error: 'غير موجود' });
  if (out.error) return res.status(409).json({ error: out.error });
  res.json(out);
});

app.get('/api/personnel/:id/history', (req, res) =>
  res.json(db.prepare("SELECT * FROM movements WHERE kind='personnel' AND ref_id=? ORDER BY id DESC").all(req.params.id)));

app.post('/api/personnel/:id/entry', (req, res) => {
  const person = db.prepare('SELECT * FROM personnel WHERE id=?').get(req.params.id);
  if (!person) return res.status(404).json({ error: 'الفرد غير موجود' });
  const rank = RK[person.rank];
  if (!rank || !['officers', 'ncos', 'civilians'].includes(rank.grp))
    return res.status(400).json({ error: 'اختر ضابطاً أو صف ضابط أو مدنياً؛ عودة الجندي تسجل من الإجازات' });
  if (person.inside) return res.status(409).json({ error: 'الفرد مسجل بالداخل بالفعل' });
  const ts = now();
  if (db.prepare("SELECT 1 FROM movements WHERE kind='personnel' AND ref_id=? AND type='IN' AND ts LIKE ? LIMIT 1").get(person.id, ts.slice(0, 10) + '%'))
    return res.status(409).json({ error: 'تم تسجيل دخول الفرد اليوم بالفعل' });
  db.transaction(() => {
    db.prepare('UPDATE personnel SET inside=1 WHERE id=?').run(person.id);
    db.prepare('INSERT INTO movements(kind,ref_id,ref,type,ts,gate,details,by) VALUES (?,?,?,?,?,?,?,?)')
      .run('personnel', person.id, person.name, 'IN', ts, gateOf(req.body.gate), String(req.body.reason || '').trim() || null, req.user.username);
  })();
  res.json({ ok: true, name: person.name, inside: 1 });
});

app.get('/api/movements', (req, res) => res.json(db.prepare('SELECT * FROM movements WHERE ts LIKE ? ORDER BY id DESC').all(day(req.query.date) + '%')));



// ---------- personnel list (with daily roll / mission / duty flags) ----------
app.get('/api/personnel', (req, res) => {
  const d = day(req.query.date);
  const rows = db.prepare(`SELECT p.*,
    EXISTS(SELECT 1 FROM attendance a WHERE a.person_id=p.id AND a.date=?) AS present,
    EXISTS(SELECT 1 FROM movements mi WHERE mi.kind='personnel' AND mi.ref_id=p.id AND mi.type='IN' AND mi.ts LIKE ? || '%') AS entered_today,
    EXISTS(SELECT 1 FROM movements mo WHERE mo.kind='personnel' AND mo.ref_id=p.id AND mo.type='OUT' AND mo.ts LIKE ? || '%') AS exited_today,
    (SELECT mi.ts FROM movements mi WHERE mi.kind='personnel' AND mi.ref_id=p.id AND mi.type='IN' AND mi.ts LIKE ? || '%' ORDER BY mi.id LIMIT 1) AS entry_time_today,
    (SELECT mo.ts FROM movements mo WHERE mo.kind='personnel' AND mo.ref_id=p.id AND mo.type='OUT' AND mo.ts LIKE ? || '%' ORDER BY mo.id LIMIT 1) AS exit_time_today,
    (SELECT m.details FROM movements m WHERE m.kind='personnel' AND m.ref_id=p.id ORDER BY m.id DESC LIMIT 1) AS latest_movement_details,
    (SELECT COUNT(*) FROM missions m WHERE m.person_id=p.id AND m.status='active') AS on_mission,
    EXISTS(SELECT 1 FROM duty u WHERE u.person_id=p.id AND ? BETWEEN u.start_date AND u.end_date) AS on_duty
    FROM personnel p`).all(d, d, d, d, d, d);
  res.json(rows.map(p => { const r = RK[p.rank] || { cat: '—', grp: 'soldiers', ord: -1 }; return { ...p, cat: p.retired ? 'ضابط متقاعد' : r.cat, grp: p.retired ? 'retired_officers' : r.grp, ord: r.ord }; })
    .sort((a, b) => b.ord - a.ord || a.name.localeCompare(b.name, 'ar')));
});

// ---------- daily roll (التمام) ----------
app.post('/api/attendance', (req, res) => {
  const d = String(req.body.date || ''), ids = (Array.isArray(req.body.ids) ? req.body.ids : []).map(Number).filter(Number.isInteger).slice(0, 2000);
  if (!validCorrespondenceDate(d)) return res.status(400).json({ error: 'تاريخ التمام غير صحيح' });
  if (!ids.length) return res.status(400).json({ error: 'اختر فرداً واحداً على الأقل' });
  const retiredCount = db.prepare(`SELECT COUNT(*) AS count FROM personnel
    WHERE retired=1 AND id IN (${ids.map(() => '?').join(',')})`).get(...ids).count;
  if (retiredCount !== ids.length) return res.status(400).json({ error: 'التمام اليومي متاح للضباط المتقاعدين فقط' });
  const ins = db.prepare('INSERT OR IGNORE INTO attendance(date,person_id,by) VALUES (?,?,?)'), del = db.prepare('DELETE FROM attendance WHERE date=? AND person_id=?');
  try { db.transaction(() => ids.forEach(id => req.body.present ? ins.run(d, id, req.user.username) : del.run(d, id)))(); res.json({ ok: true }); }
  catch { res.status(400).json({ error: 'فرد غير موجود' }); }
});

const dateOk = s => /^\d{4}-\d{2}-\d{2}$/.test(s || '');
const personOf = id => db.prepare('SELECT id,name FROM personnel WHERE id=?').get(id);

// ---------- duty (نوبتجي) ----------
app.get('/api/duty', (req, res) => res.json(db.prepare('SELECT d.*, p.rank FROM duty d LEFT JOIN personnel p ON p.id=d.person_id WHERE ? BETWEEN d.start_date AND d.end_date ORDER BY d.id DESC').all(day(req.query.date))
  .map(record => ({ ...record, role_label: dutyRoleLabel(record.role) }))));
app.post('/api/duty', (req, res) => {
  const b = req.body, p = personOf(b.person_id);
  if (!p) return res.status(400).json({ error: 'اختر الفرد' });
  if (!dateOk(b.start_date) || !dateOk(b.end_date) || b.end_date < b.start_date) return res.status(400).json({ error: 'التواريخ غير صحيحة' });
  if ((new Date(b.end_date) - new Date(b.start_date)) / 864e5 > 30) return res.status(400).json({ error: 'أقصى مدة ٣١ يوماً' });
  const role = String(b.role || 'general');
  if (!Object.hasOwn(DUTY_ROLES, role)) return res.status(400).json({ error: 'اختر صفة نوبتجية صحيحة' });
  if (role !== 'general' && db.prepare('SELECT 1 FROM duty WHERE role=? AND role<>? AND start_date<=? AND end_date>=?').get(role, 'general', b.end_date, b.start_date)) {
    return res.status(409).json({ error: 'هذه الصفة مسجلة لفرد آخر خلال جزء من الفترة المحددة' });
  }
  db.prepare('INSERT INTO duty(person_id,person,start_date,end_date,notes,by,role) VALUES (?,?,?,?,?,?,?)')
    .run(p.id, p.name, b.start_date, b.end_date, String(b.notes || '').trim(), req.user.username, role);
  res.json({ ok: true });
});
app.delete('/api/admin/duty/:id', admin, (req, res) => { db.prepare('DELETE FROM duty WHERE id=?').run(req.params.id); res.json({ ok: true }); });

// ---------- missions (المأموريات) ----------
app.get('/api/missions', (q, res) => res.json(db.prepare("SELECT m.*, p.rank FROM missions m LEFT JOIN personnel p ON p.id=m.person_id ORDER BY (m.status='active') DESC, m.id DESC LIMIT 300").all()));
app.post('/api/missions', (req, res) => {
  const b = req.body, p = personOf(b.person_id), dest = String(b.dest || '').trim();
  const startAt = editableTimestamp(b.start_at);
  if (!p || !dest || !startAt) return res.status(400).json({ error: 'اختر الفرد وجهة المأمورية وتاريخ ووقت الخروج' });
  db.prepare('INSERT INTO missions(person_id,person,dest,purpose,notes,start_date,start_ts,by) VALUES (?,?,?,?,?,?,?,?)')
    .run(p.id, p.name, dest, String(b.purpose || '').trim(), String(b.notes || '').trim(), startAt.slice(0, 10), startAt, req.user.username);
  res.json({ ok: true });
});
app.post('/api/missions/:id/return', (req, res) => {
  const r = db.prepare("UPDATE missions SET status='returned', end_ts=? WHERE id=? AND status='active'").run(now(), req.params.id);
  r.changes ? res.json({ ok: true }) : res.status(404).json({ error: 'المأمورية غير جارية' });
});
app.delete('/api/admin/missions/:id', admin, (req, res) => { db.prepare('DELETE FROM missions WHERE id=?').run(req.params.id); res.json({ ok: true }); });

// ---------- clinics, leaves, and visitors ----------
app.get('/api/clinics', (req, res) => {
  const d = day(req.query.date);
  res.json(db.prepare("SELECT c.*, p.rank FROM clinics c LEFT JOIN personnel p ON p.id=c.person_id WHERE c.start_ts LIKE ? OR c.status='active' ORDER BY (c.status='active') DESC, c.id DESC LIMIT 300").all(d + '%'));
});
app.post('/api/clinics', (req, res) => {
  const b = req.body, p = personOf(b.person_id), condition = String(b.case || '').trim();
  if (!p || !condition) return res.status(400).json({ error: 'اختر الفرد وسجّل الحالة' });
  db.prepare('INSERT INTO clinics(person_id,person,case_details,hospital,start_ts,by) VALUES (?,?,?,?,?,?)')
    .run(p.id, p.name, condition, String(b.hospital || '').trim(), now(), req.user.username);
  res.json({ ok: true });
});
app.post('/api/clinics/:id/return', (req, res) => {
  const r = db.prepare("UPDATE clinics SET status='returned', end_ts=? WHERE id=? AND status='active'").run(now(), req.params.id);
  r.changes ? res.json({ ok: true }) : res.status(404).json({ error: 'حالة العيادة غير جارية' });
});
app.delete('/api/admin/clinics/:id', admin, (req, res) => { db.prepare('DELETE FROM clinics WHERE id=?').run(req.params.id); res.json({ ok: true }); });

app.get('/api/leaves', (req, res) => {
  const d = day(req.query.date);
  res.json(db.prepare(`SELECT l.*, p.rank FROM leaves l LEFT JOIN personnel p ON p.id=l.person_id
    WHERE (l.departed_at IS NOT NULL AND l.returned_at IS NULL)
      OR ? BETWEEN l.start_date AND l.end_date
      OR substr(l.departed_at,1,10)=? OR substr(l.returned_at,1,10)=?
    ORDER BY (l.departed_at IS NOT NULL AND l.returned_at IS NULL) DESC,l.start_date DESC,l.id DESC`).all(d, d, d));
});
app.post('/api/leaves', (req, res) => {
  const b = req.body, p = personOf(b.person_id);
  if (!p) return res.status(400).json({ error: 'اختر الفرد' });
  const expectedReturn = String(b.expected_return_at || '');
  if (!validCorrespondenceDate(b.start_date) || !validCorrespondenceDate(expectedReturn.slice(0, 10)) ||
      expectedReturn.length !== 16 || expectedReturn[10] !== 'T' || expectedReturn.slice(0, 10) < b.start_date ||
      !/^([01]\d|2[0-3]):[0-5]\d$/.test(expectedReturn.slice(11)))
    return res.status(400).json({ error: 'أدخل تاريخ ووقت عودة صحيحين' });
  const endDate = expectedReturn.slice(0, 10);
  const result = db.prepare('INSERT INTO leaves(person_id,person,start_date,end_date,reason,by,expected_return_at) VALUES (?,?,?,?,?,?,?)')
    .run(p.id, p.name, b.start_date, endDate, String(b.reason || '').trim(), req.user.username, expectedReturn.replace('T', ' '));
  res.json({ ok: true, id: result.lastInsertRowid });
});
app.post('/api/leaves/:id/depart', (req, res) => {
  const leave = db.prepare('SELECT * FROM leaves WHERE id=?').get(req.params.id);
  if (!leave) return res.status(404).json({ error: 'الإجازة غير موجودة' });
  if (leave.departed_at || leave.returned_at) return res.status(409).json({ error: 'تم تسجيل حركة هذه الإجازة بالفعل' });
  const person = leave.person_id && db.prepare('SELECT * FROM personnel WHERE id=?').get(leave.person_id);
  const ts = now();
  db.transaction(() => {
    db.prepare('UPDATE leaves SET departed_at=? WHERE id=?').run(ts, leave.id);
    if (person) {
      db.prepare('UPDATE personnel SET inside=0 WHERE id=?').run(person.id);
      db.prepare('INSERT INTO movements(kind,ref_id,ref,type,ts,gate,details,by) VALUES (?,?,?,?,?,?,?,?)')
        .run('personnel', person.id, person.name, 'OUT', ts, gateOf(req.body.gate), 'خروج إجازة', req.user.username);
    }
  })();
  res.json({ ok: true, departed_at: ts });
});
app.post('/api/leaves/:id/return', (req, res) => {
  const leave = db.prepare('SELECT * FROM leaves WHERE id=?').get(req.params.id);
  if (!leave) return res.status(404).json({ error: 'الإجازة غير موجودة' });
  if (!leave.departed_at || leave.returned_at) return res.status(409).json({ error: 'سجّل الخروج للإجازة أولاً، أو تم تأكيد العودة بالفعل' });
  const person = leave.person_id && db.prepare('SELECT * FROM personnel WHERE id=?').get(leave.person_id);
  const ts = now();
  db.transaction(() => {
    db.prepare('UPDATE leaves SET returned_at=? WHERE id=?').run(ts, leave.id);
    if (person) {
      db.prepare('UPDATE personnel SET inside=1 WHERE id=?').run(person.id);
      db.prepare('INSERT INTO movements(kind,ref_id,ref,type,ts,gate,details,by) VALUES (?,?,?,?,?,?,?,?)')
        .run('personnel', person.id, person.name, 'IN', ts, gateOf(req.body.gate), 'عودة إجازة', req.user.username);
    }
  })();
  res.json({ ok: true, returned_at: ts });
});
app.delete('/api/admin/leaves/:id', admin, (req, res) => { db.prepare('DELETE FROM leaves WHERE id=?').run(req.params.id); res.json({ ok: true }); });

app.get('/api/visitors', (req, res) => {
  const d = day(req.query.date);
  res.json(db.prepare("SELECT * FROM visitors WHERE enter_ts LIKE ? OR exit_ts LIKE ? OR exit_ts IS NULL ORDER BY (exit_ts IS NULL) DESC, id DESC LIMIT 300").all(d + '%', d + '%'));
});
app.get('/api/visitors/suggestions', (req, res) => {
  const query = String(req.query.query || '').trim().slice(0, 100);
  res.json(db.prepare(`SELECT v.name,v.rank,v.unit FROM visitors v
    JOIN (SELECT name,MAX(id) AS latest_id FROM visitors GROUP BY name COLLATE NOCASE) latest
      ON latest.latest_id=v.id
    WHERE ?='' OR instr(lower(v.name),lower(?))>0
    ORDER BY v.name COLLATE NOCASE LIMIT 100`).all(query, query));
});
app.post('/api/visitors', (req, res) => {
  const name = String(req.body.name || '').trim(), rank = String(req.body.rank || '').trim();
  const reasonOther = req.body.reason_other === true;
  const reason = String(req.body.reason || '').trim(), unit = String(req.body.unit || '').trim();
  if (!name || !RK[rank] || !reason) return res.status(400).json({ error: 'أدخل اسم الزائر ورتبته وسبب الزيارة' });
  if (reason.length > 200) return res.status(400).json({ error: 'سبب الزيارة يجب ألا يتجاوز ٢٠٠ حرف' });
  if (!reasonOther && !db.prepare("SELECT 1 FROM lookups WHERE kind='visit_reason' AND name=?").get(reason))
    return res.status(400).json({ error: 'اختر سبب زيارة من القائمة أو اختر «أخرى» لكتابة سبب جديد' });
  if (unit && !db.prepare("SELECT 1 FROM lookups WHERE kind='unit' AND name=?").get(unit))
    return res.status(400).json({ error: 'اختر وحدة صحيحة من القائمة' });
  db.prepare('INSERT INTO visitors(name,rank,reason,unit,enter_ts,by) VALUES (?,?,?,?,?,?)')
    .run(name, rank, reason, unit, now(), req.user.username);
  res.json({ ok: true });
});
app.post('/api/visitors/:id/exit', (req, res) => {
  const r = db.prepare('UPDATE visitors SET exit_ts=? WHERE id=? AND exit_ts IS NULL').run(now(), req.params.id);
  r.changes ? res.json({ ok: true }) : res.status(404).json({ error: 'الزائر غير مسجل بالداخل' });
});
app.delete('/api/admin/visitors/:id', admin, (req, res) => { db.prepare('DELETE FROM visitors WHERE id=?').run(req.params.id); res.json({ ok: true }); });

function editableTimestamp(value) {
  const text = String(value || '');
  if (!/^\d{4}-\d{2}-\d{2}T([01]\d|2[0-3]):[0-5]\d$/.test(text) ||
      !validCorrespondenceDate(text.slice(0, 10))) return null;
  return text.replace('T', ' ') + ':00';
}
app.put('/api/records/:kind/:id/time', (req, res) => {
  const { kind } = req.params;
  const id = Number(req.params.id);
  const field = String(req.body.field || '');
  const timestamp = editableTimestamp(req.body.value);
  if (!Number.isInteger(id) || id < 1 || !timestamp) return res.status(400).json({ error: 'التاريخ أو الوقت غير صحيح' });
  const mappings = {
    movement: {
      fields: { ts: { table: 'movements', column: 'ts' } },
      query: 'SELECT id,by,ts,kind,ref_id,ref,type,details FROM movements WHERE id=?',
    },
    visitor: {
      fields: { enter_ts: { table: 'visitors', column: 'enter_ts' }, exit_ts: { table: 'visitors', column: 'exit_ts' } },
      query: 'SELECT id,by,enter_ts,exit_ts FROM visitors WHERE id=?',
    },
    clinic: {
      fields: { start_ts: { table: 'clinics', column: 'start_ts' }, end_ts: { table: 'clinics', column: 'end_ts' } },
      query: 'SELECT id,by,start_ts,end_ts FROM clinics WHERE id=?',
    },
    leave: {
      fields: {
        departed_at: { table: 'leaves', column: 'departed_at' },
        returned_at: { table: 'leaves', column: 'returned_at' },
        expected_return_at: { table: 'leaves', column: 'expected_return_at' },
      },
      query: 'SELECT id,by,person_id,departed_at,returned_at,expected_return_at,start_date FROM leaves WHERE id=?',
    },
    mission: {
      fields: { start_ts: { table: 'missions', column: 'start_ts' }, end_ts: { table: 'missions', column: 'end_ts' } },
      query: 'SELECT id,by,start_ts,end_ts FROM missions WHERE id=?',
    },
    squad: {
      fields: { enter_ts: { table: 'squad_attendance', column: 'enter_ts' }, exit_ts: { table: 'squad_attendance', column: 'exit_ts' } },
      query: 'SELECT id,by,enter_ts,exit_ts,squad_id,person_id FROM squad_attendance WHERE id=?',
    },
  };
  const mapping = mappings[kind], selectedField = mapping?.fields[field];
  if (!mapping || !selectedField) return res.status(400).json({ error: 'نوع سجل أو وقت غير مدعوم للتعديل' });
  const record = db.prepare(mapping.query).get(id);
  if (!record) return res.status(404).json({ error: 'السجل غير موجود' });
  if (req.user.role !== 'admin' && record.by !== req.user.username)
    return res.status(403).json({ error: 'يمكنك تعديل السجلات التي سجلتها أنت فقط' });
  const otherField = {
    visitor: field === 'enter_ts' ? 'exit_ts' : 'enter_ts',
    clinic: field === 'start_ts' ? 'end_ts' : 'start_ts',
    leave: field === 'departed_at' ? 'returned_at' : field === 'returned_at' ? 'departed_at' : null,
    mission: field === 'start_ts' ? 'end_ts' : 'start_ts',
    squad: field === 'enter_ts' ? 'exit_ts' : 'enter_ts',
  }[kind];
  const pairedTimestamp = otherField && record[otherField];
  if (kind === 'visitor' && (field === 'enter_ts' ? pairedTimestamp && timestamp > pairedTimestamp : timestamp < record.enter_ts))
    return res.status(400).json({ error: 'يجب أن يكون وقت الدخول قبل وقت الخروج' });
  if (kind === 'clinic' && (field === 'start_ts' ? pairedTimestamp && timestamp > pairedTimestamp : timestamp < record.start_ts))
    return res.status(400).json({ error: 'يجب أن يكون وقت الخروج للعيادة قبل وقت العودة' });
  if (kind === 'leave' && field !== 'expected_return_at' &&
      (field === 'departed_at' ? pairedTimestamp && timestamp > pairedTimestamp : timestamp < record.departed_at))
    return res.status(400).json({ error: 'يجب أن يكون وقت الخروج للإجازة قبل وقت العودة' });
  if (kind === 'leave' && field === 'expected_return_at' && timestamp.slice(0, 10) < record.start_date)
    return res.status(400).json({ error: 'موعد العودة المتوقع لا يمكن أن يسبق بداية الإجازة' });
  if (kind === 'mission' && (field === 'start_ts' ? pairedTimestamp && timestamp > pairedTimestamp : timestamp < record.start_ts))
    return res.status(400).json({ error: 'يجب أن يكون وقت الخروج للمأمورية قبل وقت العودة' });
  if (kind === 'movement') {
    const identity = record.ref_id == null ? 'kind=? AND ref_id IS NULL AND ref=?' : 'kind=? AND ref_id=?';
    const reference = record.ref_id == null ? record.ref : record.ref_id;
    const previous = db.prepare(`SELECT ts FROM movements WHERE ${identity} AND id<? ORDER BY id DESC LIMIT 1`)
      .get(record.kind, reference, id);
    const next = db.prepare(`SELECT ts FROM movements WHERE ${identity} AND id>? ORDER BY id ASC LIMIT 1`)
      .get(record.kind, reference, id);
    if (previous && timestamp < previous.ts || next && timestamp > next.ts)
      return res.status(400).json({ error: 'يجب الحفاظ على الترتيب الزمني لحركات السجل' });
  }
  if (kind === 'squad') {
    const date = timestamp.slice(0, 10);
    const squad = db.prepare('SELECT start_date,end_date FROM squads WHERE id=?').get(record.squad_id);
    if (!squad || date < squad.start_date || date > squad.end_date)
      return res.status(400).json({ error: 'تاريخ الحضور يجب أن يقع ضمن مدة الفرقة' });
    if (field === 'enter_ts' && pairedTimestamp && (timestamp > pairedTimestamp || pairedTimestamp.slice(0, 10) !== date))
      return res.status(400).json({ error: 'وقت الحضور يجب أن يسبق الانصراف في اليوم نفسه' });
    if (field === 'exit_ts' && (timestamp < record.enter_ts || date !== record.enter_ts.slice(0, 10)))
      return res.status(400).json({ error: 'وقت الانصراف يجب أن يأتي بعد الحضور في اليوم نفسه' });
    if (field === 'enter_ts' && date !== record.enter_ts.slice(0, 10) &&
        db.prepare('SELECT 1 FROM squad_attendance WHERE squad_id=? AND person_id=? AND date=? AND id<>?')
          .get(record.squad_id, record.person_id, date, id))
      return res.status(409).json({ error: 'يوجد سجل حضور لهذا الفرد في التاريخ المحدد بالفعل' });
  }
  const update = db.transaction(() => {
    db.prepare(`UPDATE ${selectedField.table} SET ${selectedField.column}=? WHERE id=?`).run(timestamp, id);
    if (kind === 'leave' && field === 'expected_return_at')
      db.prepare('UPDATE leaves SET end_date=? WHERE id=?').run(timestamp.slice(0, 10), id);
    if (kind === 'mission' && field === 'start_ts')
      db.prepare('UPDATE missions SET start_date=? WHERE id=?').run(timestamp.slice(0, 10), id);
    if (kind === 'squad' && field === 'enter_ts')
      db.prepare('UPDATE squad_attendance SET date=? WHERE id=?').run(timestamp.slice(0, 10), id);
    if (kind === 'movement' && record.kind === 'personnel' && record.ref_id &&
        ['خروج إجازة', 'عودة إجازة'].includes(record.details)) {
      const leaveField = record.details === 'خروج إجازة' ? 'departed_at' : 'returned_at';
      db.prepare(`UPDATE leaves SET ${leaveField}=? WHERE person_id=? AND ${leaveField}=?`)
        .run(timestamp, record.ref_id, record.ts);
    }
    if (kind === 'leave' && record.person_id && ['departed_at', 'returned_at'].includes(field)) {
      const details = field === 'departed_at' ? 'خروج إجازة' : 'عودة إجازة';
      db.prepare(`UPDATE movements SET ts=? WHERE kind='personnel' AND ref_id=? AND details=? AND ts=?`)
        .run(timestamp, record.person_id, details, record[field]);
    }
  });
  update();
  res.json({ ok: true, value: timestamp });
});

// ---------- training squads ----------
app.get('/api/squads', (req, res) => {
  const d = day(req.query.date);
  res.json(db.prepare(`SELECT s.*, COUNT(sm.person_id) AS member_count
    FROM squads s LEFT JOIN squad_members sm ON sm.squad_id=s.id
    WHERE ? BETWEEN s.start_date AND s.end_date
    GROUP BY s.id ORDER BY s.start_date, s.name`).all(d));
});
app.get('/api/admin/squads', admin, (req, res) => {
  res.json(db.prepare(`SELECT s.*, COUNT(sm.person_id) AS member_count
    FROM squads s LEFT JOIN squad_members sm ON sm.squad_id=s.id
    GROUP BY s.id ORDER BY s.start_date DESC, s.name`).all());
});
app.post('/api/squads', admin, (req, res) => {
  const name = String(req.body.name || '').trim(), { start_date, end_date } = req.body;
  if (!name || !dateOk(start_date) || !dateOk(end_date) || end_date < start_date) return res.status(400).json({ error: 'أدخل اسم الفرقة وتاريخي بداية ونهاية صحيحين' });
  const result = db.prepare('INSERT INTO squads(name,start_date,end_date,by) VALUES (?,?,?,?)').run(name, start_date, end_date, req.user.username);
  res.json({ id: result.lastInsertRowid });
});
app.get('/api/squads/:id', (req, res) => {
  const d = day(req.query.date), squad = db.prepare('SELECT * FROM squads WHERE id=?').get(req.params.id);
  if (!squad) return res.status(404).json({ error: 'الفرقة غير موجودة' });
  const members = db.prepare(`SELECT p.id AS person_id, p.name, p.rank, p.mid, a.id AS attendance_id, a.by AS attendance_by, a.enter_ts, a.exit_ts,
      EXISTS(SELECT 1 FROM squad_absences sa WHERE sa.squad_id=sm.squad_id AND sa.person_id=sm.person_id AND sa.date=?) AS absent,
      (SELECT COUNT(*) FROM squad_attendance history WHERE history.squad_id=sm.squad_id AND history.person_id=sm.person_id AND history.date<=? AND history.enter_ts IS NOT NULL) AS attended_count,
      (SELECT COUNT(*) FROM squad_absences history WHERE history.squad_id=sm.squad_id AND history.person_id=sm.person_id AND history.date<=?) AS absent_count
    FROM squad_members sm JOIN personnel p ON p.id=sm.person_id
    LEFT JOIN squad_attendance a ON a.squad_id=sm.squad_id AND a.person_id=sm.person_id AND a.date=?
    WHERE sm.squad_id=? ORDER BY p.name`).all(d, d, d, d, squad.id);
  res.json({ ...squad, date: d, members });
});
app.get('/api/squads/:id/members/:personId/history', (req, res) => {
  const d = day(req.query.date);
  const squad = db.prepare('SELECT * FROM squads WHERE id=?').get(req.params.id);
  if (!squad) return res.status(404).json({ error: 'الفرقة غير موجودة' });
  const member = db.prepare(`SELECT p.id,p.name,p.rank,p.mid FROM squad_members sm
    JOIN personnel p ON p.id=sm.person_id WHERE sm.squad_id=? AND sm.person_id=?`).get(squad.id, req.params.personId);
  if (!member) return res.status(404).json({ error: 'الفرد غير مسجل في هذه الفرقة' });
  const endDate = d > squad.end_date ? squad.end_date : d;
  const attended = db.prepare(`SELECT date,enter_ts,exit_ts FROM squad_attendance
    WHERE squad_id=? AND person_id=? AND date BETWEEN ? AND ? ORDER BY date DESC`).all(squad.id, member.id, squad.start_date, endDate);
  const absences = db.prepare(`SELECT date,created_at,by FROM squad_absences
    WHERE squad_id=? AND person_id=? AND date BETWEEN ? AND ? ORDER BY date DESC`).all(squad.id, member.id, squad.start_date, endDate);
  res.json({
    squad: squad.name,
    start_date: squad.start_date,
    through_date: endDate,
    member,
    attended_count: attended.length,
    absent_count: absences.length,
    attended,
    absences,
  });
});
app.post('/api/squads/:id/members', admin, (req, res) => {
  const squad = db.prepare('SELECT id FROM squads WHERE id=?').get(req.params.id), person = personOf(req.body.person_id);
  if (!squad || !person) return res.status(400).json({ error: 'اختر فرقة وفرداً صحيحين' });
  try {
    db.prepare('INSERT INTO squad_members(squad_id,person_id) VALUES (?,?)').run(squad.id, person.id);
    res.json({ ok: true });
  } catch (error) {
    if (error.code === 'SQLITE_CONSTRAINT_PRIMARYKEY' || error.code === 'SQLITE_CONSTRAINT_UNIQUE')
      return res.status(409).json({ error: 'الفرد مسجل بالفعل في هذه الفرقة' });
    throw error;
  }
});
app.delete('/api/admin/squads/:id/members/:personId', admin, (req, res) => {
  db.prepare('DELETE FROM squad_members WHERE squad_id=? AND person_id=?').run(req.params.id, req.params.personId);
  res.json({ ok: true });
});
app.post('/api/squads/:id/attendance/:action', (req, res) => {
  const { id, action } = req.params, { person_id } = req.body, d = day(req.body.date);
  if (!['in', 'out', 'absent', 'leave'].includes(action)) return res.status(404).json({ error: 'إجراء غير معروف' });
  if (!dateOk(req.body.date)) return res.status(400).json({ error: 'تاريخ الحضور غير صحيح' });
  const squad = db.prepare('SELECT * FROM squads WHERE id=?').get(id);
  if (!squad || d < squad.start_date || d > squad.end_date) return res.status(400).json({ error: 'التاريخ خارج مدة الفرقة' });
  if (!db.prepare('SELECT 1 FROM squad_members WHERE squad_id=? AND person_id=?').get(id, person_id)) return res.status(400).json({ error: 'الفرد غير مسجل في هذه الفرقة' });
  if (action === 'absent') {
    if (db.prepare('SELECT 1 FROM squad_attendance WHERE squad_id=? AND person_id=? AND date=?').get(id, person_id, d))
      return res.status(409).json({ error: 'تم تسجيل حضور الفرد بالفعل في هذا اليوم' });
    try {
      db.prepare('INSERT INTO squad_absences(squad_id,person_id,date,by,created_at) VALUES (?,?,?,?,?)')
        .run(id, person_id, d, req.user.username, now());
      return res.json({ ok: true });
    } catch (error) {
      if (error.code === 'SQLITE_CONSTRAINT_PRIMARYKEY' || error.code === 'SQLITE_CONSTRAINT_UNIQUE')
        return res.status(409).json({ error: 'تم تسجيل غياب الفرد في هذا اليوم مسبقاً' });
      throw error;
    }
  }
  if (action === 'leave') {
    const result = db.prepare('DELETE FROM squad_absences WHERE squad_id=? AND person_id=? AND date=?').run(id, person_id, d);
    return result.changes ? res.json({ ok: true }) : res.status(409).json({ error: 'الفرد غير مسجل كغائب في هذا اليوم' });
  }
  if (db.prepare('SELECT 1 FROM squad_absences WHERE squad_id=? AND person_id=? AND date=?').get(id, person_id, d))
    return res.status(409).json({ error: 'ألغِ تسجيل الغياب أولاً قبل تسجيل الحضور' });
  if (action === 'in') {
    try {
      db.prepare('INSERT INTO squad_attendance(squad_id,person_id,date,enter_ts,by) VALUES (?,?,?,?,?)').run(id, person_id, d, now(), req.user.username);
      return res.json({ ok: true });
    } catch (error) {
      if (error.code === 'SQLITE_CONSTRAINT_UNIQUE')
        return res.status(409).json({ error: 'تم تسجيل حضور الفرد في هذا اليوم مسبقاً' });
      throw error;
    }
  }
  const result = db.prepare('UPDATE squad_attendance SET exit_ts=? WHERE squad_id=? AND person_id=? AND date=? AND exit_ts IS NULL').run(now(), id, person_id, d);
  result.changes ? res.json({ ok: true }) : res.status(409).json({ error: 'لا يوجد حضور مفتوح لتسجيل الانصراف' });
});
app.delete('/api/admin/squads/:id', admin, (req, res) => { db.prepare('DELETE FROM squads WHERE id=?').run(req.params.id); res.json({ ok: true }); });

// ---------- purchases (المشتريات) — amounts are stored in piasters (integers) ----------
app.post('/api/purchases', (req, res) => {
  const b = req.body, p = personOf(b.person_id);
  if (!p) return res.status(400).json({ error: 'اختر الفرد' });
  const lines = (Array.isArray(b.lines) ? b.lines : []).map(l => ({ item: String(l.item || '').trim(), price: Math.round(parseFloat(l.price) * 100), qty: parseInt(l.qty, 10) }));
  if (!lines.length || lines.length > 100 || lines.some(l => !l.item || !(l.price >= 0) || !(l.qty > 0))) return res.status(400).json({ error: 'تحقق من الأصناف: الاسم والسعر والعدد' });
  db.transaction(() => {
    const id = db.prepare('INSERT INTO purchases(person_id,person,date,notes,by) VALUES (?,?,?,?,?)').run(p.id, p.name, dateOk(b.date) ? b.date : now().slice(0, 10), String(b.notes || '').trim(), req.user.username).lastInsertRowid;
    const ins = db.prepare('INSERT INTO purchase_lines(purchase_id,item,price_cents,qty) VALUES (?,?,?,?)');
    lines.forEach(l => ins.run(id, l.item, l.price, l.qty));
  })();
  res.json({ ok: true });
});
function purchasesBetween(f, t) {
  if (f > t) [f, t] = [t, f];
  const recs = db.prepare('SELECT pu.*, p.rank FROM purchases pu LEFT JOIN personnel p ON p.id=pu.person_id WHERE pu.date BETWEEN ? AND ? ORDER BY pu.date DESC, pu.id DESC').all(f, t);
  const L = db.prepare('SELECT item, price_cents, qty FROM purchase_lines WHERE purchase_id=? ORDER BY id');
  recs.forEach(r => { r.lines = L.all(r.id); r.total_cents = r.lines.reduce((s, l) => s + l.price_cents * l.qty, 0); });
  return recs;
}
app.get('/api/purchases', (req, res) => {
  const recs = purchasesBetween(day(req.query.from), day(req.query.to));
  res.json({ records: recs, total_cents: recs.reduce((s, r) => s + r.total_cents, 0) });
});
app.get('/api/purchases/items', (q, res) => res.json(db.prepare('SELECT item, price_cents FROM purchase_lines WHERE id IN (SELECT MAX(id) FROM purchase_lines GROUP BY item) ORDER BY item').all()));
app.delete('/api/admin/purchases/:id', admin, (req, res) => { db.prepare('DELETE FROM purchases WHERE id=?').run(req.params.id); res.json({ ok: true }); });

// ---------- dashboard ----------
function dashboardData(d, userId) {
  const c = (sql, ...a) => db.prepare(sql).get(...a).c;
  const g = {
    officers: { total: 0, inside: 0 },
    ncos: { total: 0, inside: 0 },
    soldiers: { total: 0, inside: 0 },
    retired: { total: 0, inside: 0, present: 0 },
    civilians: { total: 0, inside: 0 },
  };
  db.prepare(`SELECT p.rank,p.inside,p.retired,
    EXISTS(SELECT 1 FROM attendance a WHERE a.person_id=p.id AND a.date=?) AS present
    FROM personnel p`).all(d).forEach(p => {
    const key = p.retired ? 'retired' : (RK[p.rank] || { grp: 'soldiers' }).grp;
    const x = g[key]; x.total++; if (p.inside) x.inside++;
    if (key === 'retired' && p.present) x.present++;
  });
  const squadAttendance = db.prepare(`SELECT COUNT(*) AS total,
    SUM(CASE WHEN a.id IS NOT NULL THEN 1 ELSE 0 END) AS present,
    SUM(CASE WHEN sa.person_id IS NOT NULL THEN 1 ELSE 0 END) AS absent
    FROM squads s JOIN squad_members sm ON sm.squad_id=s.id
    LEFT JOIN squad_attendance a ON a.squad_id=sm.squad_id AND a.person_id=sm.person_id AND a.date=?
    LEFT JOIN squad_absences sa ON sa.squad_id=sm.squad_id AND sa.person_id=sm.person_id AND sa.date=?
    WHERE ? BETWEEN s.start_date AND s.end_date`).get(d, d, d);
  const activeVehicles = db.prepare('SELECT plate,vehicle,NULL AS ts FROM vehicles WHERE inside=1 ORDER BY plate').all();
  const activeManualVehicles = db.prepare(`SELECT m.ref AS plate,COALESCE(m.details,'مركبة زائرة') AS vehicle,m.ts
    FROM movements m
    WHERE m.kind='vehicle' AND m.ref_id IS NULL AND m.type='IN'
      AND NOT EXISTS (SELECT 1 FROM movements last
        WHERE last.kind='vehicle' AND last.ref_id IS NULL AND last.ref=m.ref AND last.id>m.id)
    ORDER BY m.ts,m.id`).all();
  activeVehicles.push(...activeManualVehicles);
  const activeMissions = db.prepare("SELECT person,dest,start_date FROM missions WHERE status='active' ORDER BY start_date,id").all();
  const activeClinics = db.prepare("SELECT person,hospital,start_ts FROM clinics WHERE status='active' ORDER BY start_ts,id").all();
  const activeVisitors = db.prepare('SELECT name,rank,unit,enter_ts FROM visitors WHERE exit_ts IS NULL ORDER BY enter_ts,id').all();
  const unclosedSquadAttendance = db.prepare(`SELECT s.name AS squad_name,p.name AS person,a.enter_ts
    FROM squad_attendance a JOIN squads s ON s.id=a.squad_id
    JOIN personnel p ON p.id=a.person_id
    WHERE a.date=? AND a.exit_ts IS NULL AND ? BETWEEN s.start_date AND s.end_date
    ORDER BY s.name,p.name`).all(d, d);
  const pendingTasks = db.prepare(`SELECT t.title,t.details,a.id AS assignment_id
    FROM task_assignments a JOIN tasks t ON t.id=a.task_id
    WHERE a.recipient_id=? AND a.confirmed_at IS NULL ORDER BY t.id`).all(userId);
  const vehiclesInside = activeVehicles.length;
  const missionsActive = activeMissions.length;
  const clinicsActive = activeClinics.length;
  const visitors = {
    total: c('SELECT COUNT(*) c FROM visitors WHERE enter_ts LIKE ? OR exit_ts LIKE ?', d + '%', d + '%'),
    inside: activeVisitors.length,
  };
  const leavesActive = c('SELECT COUNT(*) c FROM leaves WHERE ? BETWEEN start_date AND end_date', d);
  const outstanding = [
    ...activeMissions.map(item => ({ type: 'mission', label: `${item.person} — ${item.dest}`, since: item.start_date })),
    ...activeClinics.map(item => ({ type: 'clinic', label: `${item.person} — ${item.hospital || 'عيادة'}`, since: item.start_ts })),
    ...activeVisitors.map(item => ({ type: 'visitor', label: `${item.name} — ${item.rank}${item.unit ? ` — ${item.unit}` : ''}`, since: item.enter_ts })),
    ...activeVehicles.map(item => ({ type: 'vehicle', label: `${item.plate} — ${item.vehicle}`, since: item.ts || '' })),
    ...unclosedSquadAttendance.map(item => ({ type: 'squad', label: `${item.person} — ${item.squad_name}`, since: item.enter_ts })),
    ...pendingTasks.map(item => ({ type: 'task', label: item.title, since: '' })),
  ];
  return {
    groups: g,
    retiredAttendance: { total: g.retired.total, present: g.retired.present },
    squadAttendance: {
      total: squadAttendance.total || 0,
      present: squadAttendance.present || 0,
      absent: squadAttendance.absent || 0,
      leave: Math.max(0, (squadAttendance.total || 0) - (squadAttendance.present || 0) - (squadAttendance.absent || 0)),
    },
    vehiclesInside,
    activeVehicles,
    entries: c("SELECT COUNT(*) c FROM movements WHERE ts LIKE ? AND type='IN'", d + '%'),
    exits: c("SELECT COUNT(*) c FROM movements WHERE ts LIKE ? AND type='OUT'", d + '%'),
    missionsActive,
    activeMissions,
    clinicsActive,
    activeClinics,
    leavesActive,
    visitors,
    activeVisitors,
    unclosedSquadAttendance,
    pendingTasks,
    outstanding,
    duty: db.prepare('SELECT id, person, role FROM duty WHERE ? BETWEEN start_date AND end_date').all(d)
      .map(record => ({ ...record, role_label: dutyRoleLabel(record.role) })),
    purchasesTotal: c('SELECT COALESCE(SUM(l.price_cents*l.qty),0) c FROM purchase_lines l JOIN purchases p ON p.id=l.purchase_id WHERE p.date=?', d),
    recent: db.prepare('SELECT * FROM movements WHERE ts LIKE ? ORDER BY id DESC LIMIT 8').all(d + '%'),
  };
}

const handoverSnapshot = data => ({
  retiredAttendance: data.retiredAttendance,
  retiredPresence: { total: data.groups.retired.total, inside: data.groups.retired.inside },
  squadAttendance: data.squadAttendance,
  entries: data.entries,
  exits: data.exits,
  vehiclesInside: data.vehiclesInside,
  missionsActive: data.missionsActive,
  clinicsActive: data.clinicsActive,
  leavesActive: data.leavesActive,
  visitors: data.visitors,
  outstanding: data.outstanding,
});
const handoverRecord = row => row && ({ ...row, summary: JSON.parse(row.summary_json) });

app.get('/api/dashboard', (req, res) => {
  const d = day(req.query.date);
  res.json(dashboardData(d, req.user.id));
});

app.get('/api/handovers', (req, res) => {
  const d = day(req.query.date);
  res.json(handoverRecord(db.prepare('SELECT * FROM daily_handovers WHERE date=?').get(d)) || null);
});

app.post('/api/handovers', (req, res) => {
  const d = String(req.body.date || '');
  const parsed = dateOk(d) ? new Date(`${d}T00:00:00`) : null;
  if (!parsed || Number.isNaN(parsed.getTime()) ||
      parsed.getFullYear() !== Number(d.slice(0, 4)) ||
      parsed.getMonth() + 1 !== Number(d.slice(5, 7)) ||
      parsed.getDate() !== Number(d.slice(8, 10))) {
    return res.status(400).json({ error: 'اختر تاريخاً صحيحاً للتسليم' });
  }
  const notes = String(req.body.notes || '').trim();
  if (notes.length > 5000) return res.status(400).json({ error: 'ملاحظات التسليم يجب ألا تتجاوز ٥٠٠٠ حرف' });
  const user = req.user.full_name || req.user.username;
  const data = dashboardData(d, req.user.id);
  const result = db.prepare(`INSERT OR IGNORE INTO daily_handovers
    (date,summary_json,notes,prepared_by,prepared_by_name,prepared_at)
    VALUES (?,?,?,?,?,?)`).run(d, JSON.stringify(handoverSnapshot(data)), notes, req.user.id, user, now());
  if (!result.changes) return res.status(409).json({ error: 'تم حفظ تسليم هذا اليوم بالفعل' });
  res.status(201).json(handoverRecord(db.prepare('SELECT * FROM daily_handovers WHERE date=?').get(d)));
});

app.post('/api/handovers/:date/acknowledge', (req, res) => {
  const d = String(req.params.date || '');
  if (!validCorrespondenceDate(d)) return res.status(400).json({ error: 'تاريخ التسليم غير صحيح' });
  const record = db.prepare('SELECT * FROM daily_handovers WHERE date=?').get(d);
  if (!record) return res.status(404).json({ error: 'لا يوجد تسليم محفوظ لهذا اليوم' });
  if (record.acknowledged_at) return res.status(409).json({ error: 'تم تأكيد استلام التسليم بالفعل' });
  if (record.prepared_by === req.user.id) return res.status(409).json({ error: 'يجب أن يؤكد مستخدم آخر استلام التسليم' });
  const user = req.user.full_name || req.user.username;
  db.prepare(`UPDATE daily_handovers SET acknowledged_by=?,acknowledged_by_name=?,acknowledged_at=?
    WHERE date=? AND acknowledged_at IS NULL`).run(req.user.id, user, now(), d);
  res.json(handoverRecord(db.prepare('SELECT * FROM daily_handovers WHERE date=?').get(d)));
});

function movementReport(req, res) {
  const from = String(req.query.from || ''), to = String(req.query.to || '');
  const gate = String(req.query.gate || ''), section = String(req.query.section || '');
  if (!dateOk(from) || !dateOk(to) || from > to) {
    res.status(400).json({ error: 'اختر فترة زمنية صحيحة' });
    return null;
  }
  if (section && !['personnel', 'vehicle'].includes(section)) {
    res.status(400).json({ error: 'القسم المحدد غير صحيح' });
    return null;
  }
  if (gate && !db.prepare("SELECT 1 FROM lookups WHERE kind='gate' AND name=?").get(gate)) {
    res.status(400).json({ error: 'البوابة المحددة غير موجودة' });
    return null;
  }
  const conditions = ['ts>=?', 'ts<=?'], params = [`${from} 00:00:00`, `${to} 23:59:59`];
  if (gate) { conditions.push('gate=?'); params.push(gate); }
  if (section) { conditions.push('kind=?'); params.push(section); }
  const rows = db.prepare(`SELECT * FROM movements WHERE ${conditions.join(' AND ')} ORDER BY ts DESC, id DESC`).all(...params);
  return { from, to, gate, section, rows };
}
const reportSectionNames = { personnel: 'أفراد', vehicle: 'مركبات' };
const reportMoveNames = { IN: 'دخول', OUT: 'خروج' };
const reportSummary = ({ rows, from, to, gate, section }) => ({
  from, to, gate: gate || 'كل البوابات', section: reportSectionNames[section] || 'كل الأقسام',
  total: rows.length,
  entries: rows.filter(row => row.type === 'IN').length,
  exits: rows.filter(row => row.type === 'OUT').length,
});
app.get('/api/reports/daily', (req, res) => {
  const d = String(req.query.date || '');
  if (!validCorrespondenceDate(d)) return res.status(400).json({ error: 'تاريخ التقرير غير صحيح' });
  const start = `${d} 00:00:00`, end = `${d} 23:59:59`;
  const leaves = {
    departures: db.prepare(`SELECT l.id,l.person,l.reason,l.departed_at,l.expected_return_at,l.end_date,p.rank
      FROM leaves l LEFT JOIN personnel p ON p.id=l.person_id
      WHERE l.departed_at BETWEEN ? AND ? ORDER BY l.departed_at,l.id`).all(start, end),
    returns: db.prepare(`SELECT l.id,l.person,l.departed_at,l.returned_at,p.rank
      FROM leaves l LEFT JOIN personnel p ON p.id=l.person_id
      WHERE l.returned_at BETWEEN ? AND ? ORDER BY l.returned_at,l.id`).all(start, end),
  };
  const visitors = db.prepare(`SELECT id,name,rank,enter_ts,exit_ts FROM visitors
    WHERE enter_ts BETWEEN ? AND ? OR exit_ts BETWEEN ? AND ? ORDER BY enter_ts,id`).all(start, end, start, end);
  const clinics = db.prepare(`SELECT c.id,c.person,c.case_details,c.hospital,c.start_ts,c.end_ts,p.rank FROM clinics c
    LEFT JOIN personnel p ON p.id=c.person_id
    WHERE c.start_ts BETWEEN ? AND ? OR c.end_ts BETWEEN ? AND ? ORDER BY c.start_ts,c.id`).all(start, end, start, end);
  const people = db.prepare(`SELECT p.id,p.name,p.rank,p.retired,
      (SELECT m.ts FROM movements m WHERE m.kind='personnel' AND m.ref_id=p.id AND m.type='IN' AND m.ts BETWEEN ? AND ? ORDER BY m.ts,m.id LIMIT 1) AS entry_ts,
      (SELECT m.id FROM movements m WHERE m.kind='personnel' AND m.ref_id=p.id AND m.type='IN' AND m.ts BETWEEN ? AND ? ORDER BY m.ts,m.id LIMIT 1) AS entry_id,
      (SELECT m.ts FROM movements m WHERE m.kind='personnel' AND m.ref_id=p.id AND m.type='OUT' AND m.ts BETWEEN ? AND ? ORDER BY m.ts,m.id LIMIT 1) AS exit_ts,
      (SELECT m.id FROM movements m WHERE m.kind='personnel' AND m.ref_id=p.id AND m.type='OUT' AND m.ts BETWEEN ? AND ? ORDER BY m.ts,m.id LIMIT 1) AS exit_id,
      (SELECT m.details FROM movements m WHERE m.kind='personnel' AND m.ref_id=p.id AND m.type='IN' AND m.ts BETWEEN ? AND ? ORDER BY m.ts,m.id LIMIT 1) AS entry_details
    FROM personnel p WHERE EXISTS(SELECT 1 FROM movements m WHERE m.kind='personnel' AND m.ref_id=p.id AND m.ts BETWEEN ? AND ?)
    ORDER BY COALESCE(entry_ts,exit_ts),p.name`).all(start, end, start, end, start, end, start, end, start, end, start, end);
  const duty = db.prepare(`SELECT d.person_id,d.person,d.role,p.rank FROM duty d LEFT JOIN personnel p ON p.id=d.person_id
    WHERE ? BETWEEN d.start_date AND d.end_date ORDER BY d.id`).all(d).map(item => ({ ...item, role_label: dutyRoleLabel(item.role) }));
  res.json({ date: d, leaves, visitors, clinics, people, duty });
});
app.get('/api/reports/movements', (req, res) => {
  const report = movementReport(req, res);
  if (!report) return;
  res.json({ ...reportSummary(report), rows: report.rows });
});
app.get('/api/reports/movements.xlsx', async (req, res, next) => {
  const report = movementReport(req, res);
  if (!report) return;
  try {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'التحكم في البوابات';
    workbook.created = new Date();
    const sheet = workbook.addWorksheet('سجل الحركة', { views: [{ rightToLeft: true, state: 'frozen', ySplit: 1 }] });
    sheet.columns = [
      { header: 'القسم', key: 'section', width: 16 },
      { header: 'المرجع', key: 'ref', width: 28 },
      { header: 'الحركة', key: 'type', width: 14 },
      { header: 'الوجهة', key: 'destination', width: 28 },
      { header: 'بيانات إضافية', key: 'details', width: 36 },
      { header: 'التاريخ والوقت', key: 'ts', width: 22 },
      { header: 'البوابة', key: 'gate', width: 24 },
      { header: 'بواسطة', key: 'by', width: 22 },
    ];
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5BC4' } };
    sheet.autoFilter = { from: 'A1', to: `H${report.rows.length + 1}` };
    report.rows.forEach(row => sheet.addRow({
      section: reportSectionNames[row.kind] || row.kind,
      ref: row.ref,
      type: reportMoveNames[row.type] || row.type,
      destination: row.destination || '',
      details: row.details || '',
      ts: row.ts,
      gate: row.gate || '',
      by: row.by || '',
    }));
    const summary = workbook.addWorksheet('ملخص', { views: [{ rightToLeft: true }] });
    summary.addRows([
      ['الفترة من', report.from],
      ['الفترة إلى', report.to],
      ['البوابة', report.gate || 'كل البوابات'],
      ['القسم', reportSectionNames[report.section] || 'كل الأقسام'],
      ['إجمالي الحركات', report.rows.length],
      ['الدخول', report.rows.filter(row => row.type === 'IN').length],
      ['الخروج', report.rows.filter(row => row.type === 'OUT').length],
    ]);
    summary.getColumn(1).width = 24;
    summary.getColumn(2).width = 30;
    summary.getColumn(1).font = { bold: true };
    const filename = `gate-movements-${report.from}-${report.to}.xlsx`;
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${filename}"`,
    });
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    next(error);
  }
});

// ---------- CSV export ----------
app.get('/api/export.csv', (req, res) => {
  const q = x => `"${String(x ?? '').replace(/"/g, '""')}"`, kind = req.query.kind;
  let name, rows;
  if (kind === 'purchases') {
    const f = day(req.query.from), t = day(req.query.to), recs = purchasesBetween(f, t); name = `purchases-${f}_${t}`;
    rows = [['التاريخ', 'الفرد', 'الرتبة', 'الصنف', 'السعر', 'العدد', 'الإجمالي']];
    recs.forEach(r => r.lines.forEach(l => rows.push([r.date, r.person, r.rank || '', l.item, l.price_cents / 100, l.qty, l.price_cents * l.qty / 100])));
    rows.push(['', '', '', '', '', 'الإجمالي العام', recs.reduce((s, r) => s + r.total_cents, 0) / 100]);
  } else {
    const d = day(req.query.date), k = ['personnel', 'vehicle'].includes(kind) ? kind : null; name = `movements-${d}`;
    const L = { personnel: 'أفراد', vehicle: 'مركبات' }, T = { IN: 'دخول', OUT: 'خروج' };
    rows = [['القسم', 'المرجع', 'الحركة', 'الوجهة', 'الوقت', 'البوابة', 'بواسطة'],
      ...db.prepare(`SELECT * FROM movements WHERE ts LIKE ? ${k ? 'AND kind=?' : ''} ORDER BY id`).all(...[d + '%', k].filter(Boolean)).map(m => [L[m.kind], m.ref, T[m.type], m.destination || '', m.ts, m.gate || '', m.by])];
  }
  res.set({ 'Content-Type': 'text/csv; charset=utf-8', 'Content-Disposition': `attachment; filename="${name}.csv"` });
  res.send('\uFEFF' + rows.map(r => r.map(q).join(',')).join('\r\n'));
});

const PORT = Number(process.env.PORT || 3000);
if (!Number.isInteger(PORT) || PORT < 1 || PORT > 65535) {
  throw new Error('PORT must be an integer between 1 and 65535');
}
const server = app.listen({ port: PORT, exclusive: true }, () => {
  console.log(`Gate Control running on http://localhost:${PORT}`);
  ensureDailyBackup().catch(error => console.error('Unable to create the daily database backup:', error));
  const backupCheck = setInterval(
    () => ensureDailyBackup().catch(error => console.error('Unable to create the daily database backup:', error)),
    60 * 60 * 1000
  );
  backupCheck.unref();
});
server.on('error', error => {
  if (error.code === 'EADDRINUSE') {
    console.error(`SERVER_PORT_IN_USE:${PORT}`);
    process.exitCode = 1;
    return;
  }
  console.error('Unable to start the HTTP server:', error);
  process.exitCode = 1;
});
