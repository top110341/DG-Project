import express from 'express';
import { createClient } from '@libsql/client';
import { handleUpload } from '@vercel/blob/client';
import path from 'path';
import cors from 'cors';
import crypto from 'crypto';
import { fileURLToPath } from 'url';

// ES module __dirname
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// ---- Cryptographic Hashing Helpers ----
function generateSalt() {
  return crypto.randomBytes(16).toString('hex');
}

function hashPassword(password, salt) {
  const hash = crypto.scryptSync(password, salt, 64).toString('hex');
  return `${salt}:${hash}`;
}

function verifyPassword(password, storedHash) {
  if (!storedHash || !storedHash.includes(':')) return false;
  const [salt, originalHash] = storedHash.split(':');
  const inputHash = crypto.scryptSync(password, salt, 64).toString('hex');
  return inputHash === originalHash;
}

// ---- Input Length Validation Helper ----
// Client-side maxlength is UX-only; direct API calls can bypass it, so re-check server-side.
function tooLong(value, max) {
  return typeof value === 'string' && value.length > max;
}

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
// ALLOWED_ORIGIN restricts CORS in production; unset (local dev) stays permissive.
app.use(cors(process.env.ALLOWED_ORIGIN ? { origin: process.env.ALLOWED_ORIGIN } : {}));
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ limit: '10mb', extended: true }));
// Vercel's "Express" framework preset routes *all* traffic (static included) through
// this one function rather than serving /public separately — so unlike app.listen()
// below, this must run in every environment, Vercel included.
app.use(express.static(path.join(__dirname, 'public')));

// File attachments upload directly from the browser to Vercel Blob (bypassing this
// server entirely) using a short-lived client token — Vercel Functions hard-cap
// request bodies at 4.5MB, well under this app's 20MB attachment limit, so files
// can never be routed through Express/multer here.
const BLOCKED_UPLOAD_EXTENSIONS = new Set([
  '.exe', '.msi', '.bat', '.cmd', '.com', '.scr', '.ps1', '.vbs', '.js',
  '.jar', '.dll', '.sh', '.app', '.apk'
]);

// Turso (libSQL) database — SQLite-compatible, works over the network on serverless
// platforms. TURSO_DATABASE_URL defaults to a local SQLite file for local dev/Docker.
const client = createClient({
  url: process.env.TURSO_DATABASE_URL || `file:${path.join(__dirname, 'database.db')}`,
  authToken: process.env.TURSO_AUTH_TOKEN
});

// On a persistent server this would finish long before real traffic arrives, but on
// serverless (Vercel) a request can hit a cold start the instant it begins — so every
// request must wait for schema init to finish before touching the database.
const dbReady = initializeDatabaseSchema().catch((err) => {
  console.error('Failed to initialize database:', err.message);
  throw err;
});

app.use(async (req, res, next) => {
  try {
    await dbReady;
    next();
  } catch (err) {
    res.status(503).json({ error: 'Database is not ready, please retry' });
  }
});

// dbRun/dbAll/dbGet keep the same signature the rest of this file already uses,
// so no route handler below needed to change when swapping sqlite3 -> libSQL.
async function dbRun(sql, params = []) {
  const result = await client.execute({ sql, args: params });
  return { lastID: result.lastInsertRowid, changes: result.rowsAffected };
}

async function dbAll(sql, params = []) {
  const result = await client.execute({ sql, args: params });
  return result.rows;
}

async function dbGet(sql, params = []) {
  const result = await client.execute({ sql, args: params });
  return result.rows[0];
}

async function migratePlaintextPasswords() {
  try {
    const users = await dbAll('SELECT id, password FROM users');
    for (const u of users) {
      if (u.password && !u.password.includes(':')) {
        const salt = generateSalt();
        const hashed = hashPassword(u.password, salt);
        await dbRun('UPDATE users SET password = ? WHERE id = ?', [hashed, u.id]);
        console.log(`Migrated password for user ID: ${u.id}`);
      }
    }
  } catch (e) {
    console.error('Password migration failed:', e.message);
  }
}

// ──────────────────────────────────────────────
// Database Schema
// ──────────────────────────────────────────────
async function initializeDatabaseSchema() {
  await dbRun(`CREATE TABLE IF NOT EXISTS users (
    id TEXT PRIMARY KEY,
    name TEXT,
    email TEXT UNIQUE,
    password TEXT,
    avatar_color TEXT,
    auth_provider TEXT DEFAULT 'local'
  )`);

  await dbRun(`CREATE TABLE IF NOT EXISTS sessions (
    token TEXT PRIMARY KEY,
    user_id TEXT,
    created_at TEXT,
    FOREIGN KEY (user_id) REFERENCES users(id)
  )`);

  await dbRun(`CREATE TABLE IF NOT EXISTS workspaces (
    id TEXT PRIMARY KEY,
    name TEXT,
    created_at TEXT,
    hourly_rate REAL DEFAULT 50
  )`);

  await dbRun(`CREATE TABLE IF NOT EXISTS projects (
    id TEXT PRIMARY KEY,
    workspace_id TEXT,
    name TEXT,
    desc TEXT,
    created_at TEXT,
    FOREIGN KEY (workspace_id) REFERENCES workspaces(id)
  )`);

  await dbRun(`CREATE TABLE IF NOT EXISTS tasks (
    id TEXT PRIMARY KEY,
    project_id TEXT,
    name TEXT,
    priority TEXT,
    due TEXT,
    status TEXT,
    milestone_id TEXT,
    recurring_pattern TEXT,
    FOREIGN KEY (project_id) REFERENCES projects(id)
  )`);

  await dbRun(`CREATE TABLE IF NOT EXISTS comments (
    id TEXT PRIMARY KEY,
    task_id TEXT,
    user_name TEXT,
    content TEXT,
    created_at TEXT,
    FOREIGN KEY (task_id) REFERENCES tasks(id)
  )`);

  await dbRun(`CREATE TABLE IF NOT EXISTS attachments (
    id TEXT PRIMARY KEY,
    task_id TEXT,
    filename TEXT,
    filepath TEXT,
    uploaded_at TEXT,
    FOREIGN KEY (task_id) REFERENCES tasks(id)
  )`);

  await dbRun(`CREATE TABLE IF NOT EXISTS notifications (
    id TEXT PRIMARY KEY,
    message TEXT,
    is_read INTEGER DEFAULT 0,
    created_at TEXT
  )`);

  await dbRun(`CREATE TABLE IF NOT EXISTS milestones (
    id TEXT PRIMARY KEY,
    project_id TEXT,
    name TEXT,
    date TEXT,
    completed INTEGER DEFAULT 0,
    FOREIGN KEY (project_id) REFERENCES projects(id)
  )`);

  await dbRun(`CREATE TABLE IF NOT EXISTS timesheets (
    id TEXT PRIMARY KEY,
    project_id TEXT,
    task_id TEXT,
    hours REAL,
    date TEXT,
    notes TEXT,
    user TEXT,
    billed INTEGER DEFAULT 0,
    invoice_id TEXT,
    FOREIGN KEY (project_id) REFERENCES projects(id),
    FOREIGN KEY (task_id) REFERENCES tasks(id)
  )`);

  await dbRun(`CREATE TABLE IF NOT EXISTS invoices (
    id TEXT PRIMARY KEY,
    project_id TEXT,
    invoice_number TEXT,
    client TEXT,
    client_address TEXT,
    date TEXT,
    due_date TEXT,
    status TEXT DEFAULT 'unpaid',
    amount REAL,
    tax_rate REAL DEFAULT 7,
    tax_amount REAL DEFAULT 0,
    discount REAL DEFAULT 0,
    total_amount REAL,
    hours REAL,
    items TEXT,
    notes TEXT,
    paid_date TEXT,
    FOREIGN KEY (project_id) REFERENCES projects(id)
  )`);

  // Safe migrations for existing databases (columns added after initial release) —
  // ALTER TABLE errors (column already exists) are expected on every run after the
  // first and are swallowed.
  const userCols = [
    "ADD COLUMN role TEXT DEFAULT 'member'",
    "ADD COLUMN manager_id TEXT DEFAULT ''",
    "ADD COLUMN department TEXT DEFAULT 'Engineering'",
    "ADD COLUMN title TEXT DEFAULT 'Team Member'",
    "ADD COLUMN avatar_url TEXT DEFAULT ''"
  ];
  for (const col of userCols) {
    await dbRun(`ALTER TABLE users ${col}`).catch(() => {});
  }

  const taskCols = [
    "ADD COLUMN assignee_id TEXT DEFAULT ''",
    "ADD COLUMN assignee_name TEXT DEFAULT 'Unassigned'",
    "ADD COLUMN assignee_avatar TEXT DEFAULT '#8b949e'",
    "ADD COLUMN start_date TEXT DEFAULT ''"
  ];
  for (const col of taskCols) {
    await dbRun(`ALTER TABLE tasks ${col}`).catch(() => {});
  }

  await dbRun("ALTER TABLE workspaces ADD COLUMN hourly_rate REAL DEFAULT 50").catch(() => {});

  console.log('Connected to database.');
  await migratePlaintextPasswords();
  await seedData();
}

// ──────────────────────────────────────────────
// Seed Data
// ──────────────────────────────────────────────
async function seedData() {
  try {
    console.log('Seeding database with demo data & hierarchy...');

    // Users
    const userCount = await dbGet('SELECT COUNT(*) as count FROM users');
    if (userCount.count === 0) {
      const demoUsers = [
        ['u1', 'Admin Demo', 'admin@demo.com', 'admin123', '#2684FF', 'local', 'admin', '', 'Executive', 'Head of Operations'],
        ['u2', 'Somchai Manager', 'somchai@demo.com', 'admin123', '#d29922', 'local', 'manager', 'u1', 'Project Management', 'Engineering Supervisor'],
        ['u3', 'Nong Dev', 'nong@demo.com', 'admin123', '#3fb950', 'local', 'member', 'u2', 'Frontend Dev', 'Senior Frontend Developer'],
        ['u4', 'Alice QA', 'alice@demo.com', 'admin123', '#bc8cff', 'local', 'member', 'u2', 'Quality Assurance', 'QA Lead'],
        ['u5', 'Bob Backend', 'bob@demo.com', 'admin123', '#58a6ff', 'local', 'member', 'u2', 'Backend Dev', 'Backend Specialist']
      ];
      for (const u of demoUsers) {
        const salt = generateSalt();
        u[3] = hashPassword(u[3], salt);
        await dbRun(
          `INSERT INTO users (id, name, email, password, avatar_color, auth_provider, role, manager_id, department, title) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
          u
        );
      }
    }

    const wsCount = await dbGet('SELECT COUNT(*) as count FROM workspaces');
    if (wsCount.count === 0) {
      // Workspaces
      await dbRun(`INSERT INTO workspaces (id, name, created_at, hourly_rate) VALUES (?, ?, ?, ?)`, ['ws1', 'Production Workspace', new Date().toISOString(), 50]);
      await dbRun(`INSERT INTO workspaces (id, name, created_at, hourly_rate) VALUES (?, ?, ?, ?)`, ['ws2', 'Development Lab', new Date().toISOString(), 50]);

      // Projects
      await dbRun(`INSERT INTO projects (id, workspace_id, name, desc, created_at) VALUES (?, ?, ?, ?, ?)`, ['p1', 'ws1', 'Corporate Website Redesign', 'Full redesign of the corporate website with modern UI/UX', new Date().toISOString()]);
      await dbRun(`INSERT INTO projects (id, workspace_id, name, desc, created_at) VALUES (?, ?, ?, ?, ?)`, ['p2', 'ws1', 'Mobile App Development', 'Cross-platform mobile application for client portal', new Date().toISOString()]);
      await dbRun(`INSERT INTO projects (id, workspace_id, name, desc, created_at) VALUES (?, ?, ?, ?, ?)`, ['p3', 'ws2', 'Marketing Campaign Q3', 'Q3 digital marketing campaign strategy and execution', new Date().toISOString()]);

      // Milestones
      await dbRun(`INSERT INTO milestones (id, project_id, name, date, completed) VALUES (?, ?, ?, ?, ?)`, ['m1', 'p1', 'Design Approved', '2026-07-10', 1]);
      await dbRun(`INSERT INTO milestones (id, project_id, name, date, completed) VALUES (?, ?, ?, ?, ?)`, ['m2', 'p1', 'Alpha Release', '2026-07-25', 0]);
      await dbRun(`INSERT INTO milestones (id, project_id, name, date, completed) VALUES (?, ?, ?, ?, ?)`, ['m3', 'p2', 'Beta Deployment', '2026-08-15', 0]);

      // Tasks
      await dbRun(`INSERT INTO tasks (id, project_id, name, priority, due, status, milestone_id, recurring_pattern, assignee_id, assignee_name, assignee_avatar) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`, ['t1', 'p1', 'Draft Figma Wireframes', 'high', '2026-07-09', 'completed', 'm1', 'none', 'u3', 'Nong Dev', '#3fb950']);
      await dbRun(`INSERT INTO tasks (id, project_id, name, priority, due, status, milestone_id, recurring_pattern, assignee_id, assignee_name, assignee_avatar) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`, ['t2', 'p1', 'Configure Design System', 'medium', '2026-07-15', 'in_progress', 'm2', 'none', 'u3', 'Nong Dev', '#3fb950']);
      await dbRun(`INSERT INTO tasks (id, project_id, name, priority, due, status, milestone_id, recurring_pattern, assignee_id, assignee_name, assignee_avatar) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`, ['t3', 'p1', 'Develop Landing Pages', 'low', '2026-07-20', 'todo', 'm2', 'none', 'u3', 'Nong Dev', '#3fb950']);
      await dbRun(`INSERT INTO tasks (id, project_id, name, priority, due, status, milestone_id, recurring_pattern, assignee_id, assignee_name, assignee_avatar) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`, ['t4', 'p1', 'Connect CMS API', 'high', '2026-07-24', 'todo', 'm2', 'none', 'u5', 'Bob Backend', '#58a6ff']);
      await dbRun(`INSERT INTO tasks (id, project_id, name, priority, due, status, milestone_id, recurring_pattern, assignee_id, assignee_name, assignee_avatar) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`, ['t5', 'p2', 'Design Mobile Screens', 'high', '2026-07-18', 'in_progress', 'm3', 'none', 'u3', 'Nong Dev', '#3fb950']);
      await dbRun(`INSERT INTO tasks (id, project_id, name, priority, due, status, milestone_id, recurring_pattern, assignee_id, assignee_name, assignee_avatar) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`, ['t6', 'p2', 'Setup Push Notifications', 'medium', '2026-07-30', 'todo', 'm3', 'weekly', 'u5', 'Bob Backend', '#58a6ff']);
      await dbRun(`INSERT INTO tasks (id, project_id, name, priority, due, status, milestone_id, recurring_pattern, assignee_id, assignee_name, assignee_avatar) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`, ['t7', 'p3', 'Website Technical Audit', 'medium', '2026-07-12', 'completed', '', 'none', 'u4', 'Alice QA', '#bc8cff']);

      // Timesheets
      await dbRun(`INSERT INTO timesheets (id, project_id, task_id, hours, date, notes, user, billed, invoice_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`, ['ts1', 'p1', 't1', 8, '2026-07-05', 'Completed wireframe drafts', 'Nong Dev', 1, 'inv1']);
      await dbRun(`INSERT INTO timesheets (id, project_id, task_id, hours, date, notes, user, billed, invoice_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`, ['ts2', 'p1', 't2', 4, '2026-07-06', 'Design system configuration', 'Nong Dev', 0, '']);
      await dbRun(`INSERT INTO timesheets (id, project_id, task_id, hours, date, notes, user, billed, invoice_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`, ['ts3', 'p2', 't5', 6, '2026-07-07', 'Mobile screen designs', 'Nong Dev', 0, '']);

      // Invoices
      const sampleItems = JSON.stringify([
        { desc: 'Wireframe design services (Task: Draft Figma Wireframes)', hours: 8, rate: 50, amount: 400 },
        { desc: 'UI Design System & Architecture Consultation', hours: 4, rate: 75, amount: 300 }
      ]);
      await dbRun(`INSERT INTO invoices (id, project_id, invoice_number, client, client_address, date, due_date, status, amount, tax_rate, tax_amount, discount, total_amount, hours, items, notes, paid_date) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`, ['inv1', 'p1', 'INV-2026-001', 'Acme Corp Thailand Co., Ltd.', '123 Sukhumvit Rd, Klongtoey, Bangkok 10110', '2026-07-05', '2026-07-20', 'paid', 700, 7, 49, 0, 749, 12, sampleItems, 'Payment received via Bank Transfer', '2026-07-10']);

      // Notifications
      await dbRun(`INSERT INTO notifications (id, message, is_read, created_at) VALUES (?, ?, ?, ?)`, ['n1', 'Welcome to AG Projects Enterprise!', 0, new Date().toISOString()]);
      await dbRun(`INSERT INTO notifications (id, message, is_read, created_at) VALUES (?, ?, ?, ?)`, ['n2', 'Team hierarchy and executive dashboard ready', 0, new Date().toISOString()]);
    } else {
      await dbRun(`UPDATE tasks SET assignee_id = 'u3', assignee_name = 'Nong Dev', assignee_avatar = '#3fb950' WHERE id IN ('t1', 't2', 't3', 't5') AND (assignee_id = '' OR assignee_id IS NULL)`);
      await dbRun(`UPDATE tasks SET assignee_id = 'u5', assignee_name = 'Bob Backend', assignee_avatar = '#58a6ff' WHERE id IN ('t4', 't6') AND (assignee_id = '' OR assignee_id IS NULL)`);
      await dbRun(`UPDATE tasks SET assignee_id = 'u4', assignee_name = 'Alice QA', assignee_avatar = '#bc8cff' WHERE id IN ('t7') AND (assignee_id = '' OR assignee_id IS NULL)`);
    }
    console.log('Database seeded and hierarchy configured successfully.');
  } catch (err) {
    console.error('Error seeding database:', err.message);
  }
}

// ──────────────────────────────────────────────
// Helper: Create Notification
// ──────────────────────────────────────────────
async function createNotification(message) {
  const id = 'n_' + Date.now() + Math.round(Math.random() * 1e4);
  const created_at = new Date().toISOString();
  await dbRun(
    `INSERT INTO notifications (id, message, is_read, created_at) VALUES (?, ?, 0, ?)`,
    [id, message, created_at]
  );
}

// ──────────────────────────────────────────────
// Helper: Handle Task Recurrence
// ──────────────────────────────────────────────
async function handleTaskRecurrence(task) {
  const validDue = (task.due && !isNaN(new Date(task.due))) ? new Date(task.due) : new Date();
  let nextDue;

  switch (task.recurring_pattern) {
    case 'daily':
      nextDue = new Date(validDue);
      nextDue.setDate(nextDue.getDate() + 1);
      break;
    case 'weekly':
      nextDue = new Date(validDue);
      nextDue.setDate(nextDue.getDate() + 7);
      break;
    case 'monthly':
      nextDue = new Date(validDue);
      nextDue.setMonth(nextDue.getMonth() + 1);
      break;
    default:
      return;
  }

  const newId = 't_' + Date.now() + Math.round(Math.random() * 1e4);
  const newName = task.name.includes('(Recurring)') ? task.name : task.name + ' (Recurring)';
  const nextDueStr = nextDue.toISOString().split('T')[0];

  // Preserve the task's original duration (days between start and due) on the next occurrence
  let nextStartDateStr = '';
  if (task.start_date && !isNaN(new Date(task.start_date))) {
    const durationMs = validDue - new Date(task.start_date);
    const nextStart = new Date(nextDue.getTime() - durationMs);
    nextStartDateStr = nextStart.toISOString().split('T')[0];
  }

  await dbRun(
    `INSERT INTO tasks (id, project_id, name, priority, due, status, milestone_id, recurring_pattern, assignee_id, assignee_name, assignee_avatar, start_date) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [newId, task.project_id, newName, task.priority, nextDueStr, 'todo', task.milestone_id, task.recurring_pattern, task.assignee_id || '', task.assignee_name || 'Unassigned', task.assignee_avatar || '#8b949e', nextStartDateStr]
  );

  await createNotification(`Recurring task created: ${newName}`);
}

// ──────────────────────────────────────────────
// Middleware: Token-based User Authentication
// ──────────────────────────────────────────────
async function authenticateUser(req, res, next) {
  try {
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({ error: 'Unauthorized: Session token missing or invalid' });
    }
    const token = authHeader.substring(7);
    const session = await dbGet('SELECT * FROM sessions WHERE token = ?', [token]);
    if (!session) {
      return res.status(401).json({ error: 'Unauthorized: Session expired or invalid' });
    }
    const user = await dbGet('SELECT * FROM users WHERE id = ?', [session.user_id]);
    if (!user) {
      return res.status(401).json({ error: 'Unauthorized: User not found' });
    }
    req.user = user;
    next();
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}

app.use((req, res, next) => {
  if (req.path === '/api/login' || req.path === '/api/login/microsoft') {
    return next();
  }
  if (req.path.startsWith('/api/')) {
    return authenticateUser(req, res, next);
  }
  next();
});

// ──────────────────────────────────────────────
// API Routes
// ──────────────────────────────────────────────

// ── Auth ──
app.post('/api/login', async (req, res) => {
  try {
    const { email, password } = req.body;
    if (!email || !password) {
      return res.status(400).json({ success: false, message: 'Email and password are required' });
    }
    const user = await dbGet('SELECT * FROM users WHERE email = ?', [email]);
    if (user && verifyPassword(password, user.password)) {
      const token = crypto.randomBytes(32).toString('hex');
      await dbRun('INSERT INTO sessions (token, user_id, created_at) VALUES (?, ?, ?)', [token, user.id, new Date().toISOString()]);
      res.json({
        success: true,
        token,
        user: {
          id: user.id,
          name: user.name,
          email: user.email,
          color: user.avatar_color,
          avatar_url: user.avatar_url || '',
          role: user.role || 'member',
          title: user.title || 'Team Member',
          department: user.department || 'Engineering',
          manager_id: user.manager_id || ''
        }
      });
    } else {
      res.status(401).json({ success: false, message: 'Invalid credentials' });
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/login/microsoft', async (req, res) => {
  try {
    const { name, email } = req.body;
    if (!email || !name) {
      return res.status(400).json({ success: false, message: 'Name and email are required' });
    }
    let user = await dbGet('SELECT * FROM users WHERE email = ?', [email]);

    if (!user) {
      const colors = ['#2684FF', '#3fb950', '#d29922', '#f85149', '#58a6ff', '#bc8cff'];
      const avatar_color = colors[Math.floor(Math.random() * colors.length)];
      const id = 'u_' + Date.now();

      await dbRun(
        `INSERT INTO users (id, name, email, password, avatar_color, auth_provider, role, manager_id, department, title, avatar_url) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '')`,
        [id, name, email, '', avatar_color, 'microsoft', 'member', '', 'Engineering', 'Team Member']
      );
      user = await dbGet('SELECT * FROM users WHERE id = ?', [id]);
    }

    const token = crypto.randomBytes(32).toString('hex');
    await dbRun('INSERT INTO sessions (token, user_id, created_at) VALUES (?, ?, ?)', [token, user.id, new Date().toISOString()]);

    res.json({
      success: true,
      token,
      user: {
        id: user.id,
        name: user.name,
        email: user.email,
        color: user.avatar_color,
        avatar_url: user.avatar_url || '',
        role: user.role || 'member',
        title: user.title || 'Team Member',
        department: user.department || 'Engineering',
        manager_id: user.manager_id || ''
      }
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/logout', async (req, res) => {
  try {
    const authHeader = req.headers.authorization;
    if (authHeader && authHeader.startsWith('Bearer ')) {
      const token = authHeader.substring(7);
      await dbRun('DELETE FROM sessions WHERE token = ?', [token]);
    }
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Profile ──
app.post('/api/profile/update', async (req, res) => {
  try {
    const { name, color, avatar_url, email } = req.body;
    await dbRun('UPDATE users SET name = ?, avatar_color = ?, avatar_url = ? WHERE email = ?', [name, color, avatar_url || '', email]);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Users & Hierarchy ──
app.get('/api/users', async (req, res) => {
  try {
    const users = await dbAll('SELECT id, name, email, avatar_color, avatar_url, role, manager_id, department, title FROM users');
    res.json(users);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/users', async (req, res) => {
  try {
    const { name, email, password, role, manager_id, department, title } = req.body;
    
    // Validations
    if (!name || !email) {
      return res.status(400).json({ error: 'Name and email are required' });
    }
    if (tooLong(name, 100) || tooLong(title, 80) || tooLong(department, 80)) {
      return res.status(400).json({ error: 'Name, title, or department is too long' });
    }
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email) || tooLong(email, 150)) {
      return res.status(400).json({ error: 'Invalid email address format' });
    }
    const rawPass = password || 'admin123';
    if (rawPass.length < 6) {
      return res.status(400).json({ error: 'Password must be at least 6 characters' });
    }

    const salt = generateSalt();
    const hashedPassword = hashPassword(rawPass, salt);

    const colors = ['#2684FF', '#3fb950', '#d29922', '#f85149', '#58a6ff', '#bc8cff'];
    const avatar_color = colors[Math.floor(Math.random() * colors.length)];
    const id = 'u_' + Date.now();
    await dbRun(
      `INSERT INTO users (id, name, email, password, avatar_color, auth_provider, role, manager_id, department, title) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [id, name, email, hashedPassword, avatar_color, 'local', role || 'member', manager_id || '', department || 'Engineering', title || 'Team Member']
    );
    const newUser = { id, name, email, avatar_color, role: role || 'member', manager_id: manager_id || '', department: department || 'Engineering', title: title || 'Team Member' };
    res.json({ success: true, user: newUser, ...newUser });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/users/:id/hierarchy', async (req, res) => {
  try {
    const { id } = req.params;
    const { role, manager_id, department, title } = req.body;
    const requesterRole = req.user.role;
    if (requesterRole === 'member') {
      return res.status(403).json({ error: 'Member cannot modify user hierarchy or roles' });
    }
    const target = await dbGet('SELECT role FROM users WHERE id = ?', [id]);
    if (requesterRole === 'manager' && target && (target.role === 'admin' || target.role === 'manager')) {
      return res.status(403).json({ error: 'Manager can only manage Member hierarchy' });
    }
    if (requesterRole === 'manager' && role === 'admin') {
      return res.status(403).json({ error: 'Manager cannot grant Admin role' });
    }
    await dbRun(
      `UPDATE users SET role = ?, manager_id = ?, department = ?, title = ? WHERE id = ?`,
      [role, manager_id || '', department || 'Engineering', title || 'Team Member', id]
    );
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/users/:id/password', async (req, res) => {
  try {
    const { id } = req.params;
    const { password } = req.body;
    const requesterRole = req.user.role;
    if (requesterRole === 'member') {
      return res.status(403).json({ error: 'Member cannot reset Manager or Admin passwords' });
    }
    const target = await dbGet('SELECT role FROM users WHERE id = ?', [id]);
    if (target && (target.role === 'admin' || target.role === 'manager') && requesterRole === 'manager') {
      return res.status(403).json({ error: 'Manager cannot reset Admin or Manager password' });
    }
    if (!password || password.trim().length < 6) {
      return res.status(400).json({ error: 'Password must be at least 6 characters' });
    }
    
    const salt = generateSalt();
    const hashedPassword = hashPassword(password.trim(), salt);
    await dbRun('UPDATE users SET password = ? WHERE id = ?', [hashedPassword, id]);
    res.json({ success: true, message: 'Password updated successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/users/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { name, email, password, role, manager_id, department, title } = req.body;
    const requesterRole = req.user.role;
    if (requesterRole === 'member') {
      return res.status(403).json({ error: 'Member cannot modify Manager or Admin accounts' });
    }
    const target = await dbGet('SELECT * FROM users WHERE id = ?', [id]);
    if (!target) {
      return res.status(404).json({ error: 'User not found' });
    }
    if ((target.role === 'admin' || target.role === 'manager') && requesterRole === 'manager') {
      return res.status(403).json({ error: 'Manager cannot modify Admin or Manager accounts' });
    }
    if (requesterRole === 'manager' && role === 'admin') {
      return res.status(403).json({ error: 'Manager cannot grant Admin role' });
    }

    const finalName = name !== undefined && name !== null ? name : target.name;
    const finalEmail = email !== undefined && email !== null ? email : target.email;
    const finalRole = role !== undefined && role !== null ? role : target.role;
    const finalManagerId = manager_id !== undefined && manager_id !== null ? manager_id : target.manager_id;
    const finalDept = department !== undefined && department !== null ? department : target.department;
    const finalTitle = title !== undefined && title !== null ? title : target.title;

    if (finalEmail) {
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailRegex.test(finalEmail) || tooLong(finalEmail, 150)) {
        return res.status(400).json({ error: 'Invalid email address format' });
      }
    }
    if (tooLong(finalName, 100) || tooLong(finalTitle, 80) || tooLong(finalDept, 80)) {
      return res.status(400).json({ error: 'Name, title, or department is too long' });
    }

    await dbRun(
      `UPDATE users SET name = ?, email = ?, role = ?, manager_id = ?, department = ?, title = ? WHERE id = ?`,
      [finalName, finalEmail, finalRole || 'member', finalManagerId || '', finalDept || 'Engineering', finalTitle || 'Team Member', id]
    );
    if (password && password.trim().length > 0) {
      if (password.trim().length < 6) {
        return res.status(400).json({ error: 'Password must be at least 6 characters' });
      }
      const salt = generateSalt();
      const hashedPassword = hashPassword(password.trim(), salt);
      await dbRun('UPDATE users SET password = ? WHERE id = ?', [hashedPassword, id]);
    }
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/users/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const requesterRole = req.user.role;
    if (requesterRole === 'member') {
      return res.status(403).json({ error: 'Member cannot delete users' });
    }
    if (id === req.user.id) {
      return res.status(400).json({ error: 'Cannot delete your own account' });
    }
    const target = await dbGet('SELECT role FROM users WHERE id = ?', [id]);
    if (!target) {
      return res.status(404).json({ error: 'User not found' });
    }
    if ((target.role === 'admin' || target.role === 'manager') && requesterRole === 'manager') {
      return res.status(403).json({ error: 'Manager cannot delete Admin or Manager accounts' });
    }
    await dbRun('DELETE FROM users WHERE id = ?', [id]);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Workspaces ──
app.get('/api/workspaces', async (req, res) => {
  try {
    const workspaces = await dbAll('SELECT * FROM workspaces ORDER BY created_at');
    res.json(workspaces);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/workspaces', async (req, res) => {
  try {
    const { name } = req.body;
    if (tooLong(name, 80)) {
      return res.status(400).json({ error: 'Workspace name must be 80 characters or fewer' });
    }
    const id = 'ws_' + Date.now();
    const created_at = new Date().toISOString();
    await dbRun('INSERT INTO workspaces (id, name, created_at, hourly_rate) VALUES (?, ?, ?, ?)', [id, name, created_at, 50]);
    await createNotification(`Workspace created: ${name}`);
    res.json({ id, name, created_at, hourly_rate: 50 });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/workspaces/:id/rate', async (req, res) => {
  try {
    const { id } = req.params;
    const { hourly_rate } = req.body;
    await dbRun('UPDATE workspaces SET hourly_rate = ? WHERE id = ?', [hourly_rate, id]);
    await createNotification(`Workspace hourly rate updated to $${hourly_rate}/hr`);
    res.json({ success: true, hourly_rate });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/workspaces/:workspace_id/executive-summary', async (req, res) => {
  try {
    const { workspace_id } = req.params;
    const projects = await dbAll('SELECT * FROM projects WHERE workspace_id = ?', [workspace_id]);
    const projectIds = projects.map(p => p.id);
    
    if (projectIds.length === 0) {
      return res.json({
        kpi: { total_projects: 0, total_tasks: 0, completed_tasks: 0, completion_percent: 0, total_hours: 0, total_revenue: 0, unpaid_revenue: 0 },
        project_health: [],
        workload_matrix: [],
        priority_feed: []
      });
    }

    const placeholders = projectIds.map(() => '?').join(',');
    const tasks = await dbAll(`SELECT * FROM tasks WHERE project_id IN (${placeholders})`, projectIds);
    const timesheets = await dbAll(`SELECT * FROM timesheets WHERE project_id IN (${placeholders})`, projectIds);
    const invoices = await dbAll(`SELECT * FROM invoices WHERE project_id IN (${placeholders})`, projectIds);
    const users = await dbAll('SELECT id, name, email, avatar_color, role, manager_id, department, title FROM users');

    const totalProjects = projects.length;
    const totalTasks = tasks.length;
    const completedTasks = tasks.filter(t => t.status === 'completed').length;
    const completionPercent = totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : 0;
    const totalHours = timesheets.reduce((sum, t) => sum + (t.hours || 0), 0);
    const validInvoices = invoices.filter(i => i.status !== 'cancelled');
    const totalRevenue = validInvoices.reduce((sum, i) => sum + (i.total_amount !== undefined && i.total_amount !== null ? i.total_amount : (i.amount || 0)), 0);
    const unpaidRevenue = validInvoices.filter(i => i.status !== 'paid').reduce((sum, i) => sum + (i.total_amount !== undefined && i.total_amount !== null ? i.total_amount : (i.amount || 0)), 0);

    const todayStr = new Date().toISOString().split('T')[0];

    const projectHealth = projects.map(p => {
      const pTasks = tasks.filter(t => t.project_id === p.id);
      const pCompleted = pTasks.filter(t => t.status === 'completed').length;
      const pInProgress = pTasks.filter(t => t.status === 'in_progress').length;
      const pTodo = pTasks.filter(t => t.status === 'todo').length;
      const pOverdue = pTasks.filter(t => t.status !== 'completed' && t.due && t.due < todayStr).length;
      const pPercent = pTasks.length > 0 ? Math.round((pCompleted / pTasks.length) * 100) : 0;
      const pHours = timesheets.filter(ts => ts.project_id === p.id).reduce((sum, ts) => sum + (ts.hours || 0), 0);
      const pRev = validInvoices.filter(inv => inv.project_id === p.id).reduce((sum, inv) => sum + (inv.total_amount !== undefined && inv.total_amount !== null ? inv.total_amount : (inv.amount || 0)), 0);

      const teamSize = new Set(pTasks.map(t => t.assignee_id).filter(Boolean)).size || 3;
      const healthStatus = pOverdue > 0 ? 'Needs Attention' : pPercent === 100 ? 'Completed' : 'Optimal';
      const statusBadge = pOverdue > 0 ? 'warning' : pPercent === 100 ? 'success' : 'progress';

      return {
        id: p.id,
        name: p.name,
        desc: p.desc,
        total_tasks: pTasks.length,
        completed_tasks: pCompleted,
        in_progress_tasks: pInProgress,
        todo_tasks: pTodo,
        overdue_tasks: pOverdue,
        progress_percent: pPercent,
        completion_percentage: pPercent,
        total_hours: pHours,
        logged_hours: pHours,
        total_revenue: pRev,
        revenue: pRev,
        team_size: teamSize,
        health_status: healthStatus,
        status_badge: statusBadge
      };
    });

    const workloadMatrix = users.filter(u => u.role !== 'admin' && (u.title || '').toLowerCase() !== 'executive director').map(u => {
      const manager = users.find(m => m.id === u.manager_id);
      const uTasks = tasks.filter(t => t.assignee_id === u.id || t.assignee_name === u.name);
      const activeTasks = uTasks.filter(t => t.status !== 'completed').length;
      const doneTasks = uTasks.filter(t => t.status === 'completed').length;
      const uHours = timesheets.filter(ts => ts.user === u.name).reduce((sum, ts) => sum + (ts.hours || 0), 0);

      return {
        id: u.id,
        name: u.name,
        email: u.email,
        avatar_color: u.avatar_color,
        avatar_url: u.avatar_url || '',
        role: u.role,
        title: u.title,
        department: u.department,
        manager_name: manager ? manager.name : 'Top Level / CEO',
        active_tasks: activeTasks,
        completed_tasks: doneTasks,
        total_tasks: uTasks.length,
        efficiency: uTasks.length > 0 ? Math.round((doneTasks / uTasks.length) * 100) : 100,
        logged_hours: uHours
      };
    });

    const priorityFeed = tasks
      .filter(t => t.status !== 'completed')
      .map(t => {
        const p = projects.find(proj => proj.id === t.project_id);
        const isOverdue = t.due && t.due < todayStr;
        return {
          ...t,
          project_name: p ? p.name : 'Unknown Project',
          is_overdue: isOverdue
        };
      })
      .sort((a, b) => {
        if (a.is_overdue && !b.is_overdue) return -1;
        if (!a.is_overdue && b.is_overdue) return 1;
        return (a.due || '9999') > (b.due || '9999') ? 1 : -1;
      })
      .slice(0, 8);

    const kpiObj = {
      total_projects: totalProjects,
      total_tasks: totalTasks,
      completed_tasks: completedTasks,
      completion_percent: completionPercent,
      overall_completion_percentage: completionPercent,
      total_hours: totalHours,
      total_logged_hours: totalHours,
      total_revenue: totalRevenue,
      unpaid_revenue: unpaidRevenue
    };

    res.json({
      kpi: kpiObj,
      kpis: kpiObj,
      project_health: projectHealth,
      workload_matrix: workloadMatrix,
      priority_feed: priorityFeed
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Projects ──
app.get('/api/projects', async (req, res) => {
  try {
    const { workspace_id } = req.query;
    let projects;
    if (workspace_id) {
      projects = await dbAll('SELECT * FROM projects WHERE workspace_id = ?', [workspace_id]);
    } else {
      projects = await dbAll('SELECT * FROM projects');
    }
    res.json(projects);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/projects', async (req, res) => {
  try {
    const { workspace_id, name, desc } = req.body;
    if (tooLong(name, 100) || tooLong(desc, 500)) {
      return res.status(400).json({ error: 'Project name or description is too long' });
    }
    const id = 'p_' + Date.now();
    const created_at = new Date().toISOString();
    await dbRun(
      'INSERT INTO projects (id, workspace_id, name, desc, created_at) VALUES (?, ?, ?, ?, ?)',
      [id, workspace_id, name, desc || '', created_at]
    );
    await createNotification(`Project created: ${name}`);
    res.json({ id, workspace_id, name, desc: desc || '', created_at });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/projects/:id', async (req, res) => {
  try {
    const { id } = req.params;
    await dbRun('DELETE FROM projects WHERE id = ?', [id]);
    await dbRun('DELETE FROM tasks WHERE project_id = ?', [id]);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Tasks ──
app.get('/api/tasks', async (req, res) => {
  try {
    const { project_id, workspace_id } = req.query;
    let tasks;
    if (project_id) {
      tasks = await dbAll('SELECT * FROM tasks WHERE project_id = ? ORDER BY due ASC', [project_id]);
    } else if (workspace_id) {
      tasks = await dbAll(
        `SELECT t.* FROM tasks t JOIN projects p ON t.project_id = p.id WHERE p.workspace_id = ? ORDER BY t.due ASC`,
        [workspace_id]
      );
    } else {
      tasks = await dbAll('SELECT * FROM tasks ORDER BY due ASC');
    }
    res.json(tasks);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/tasks', async (req, res) => {
  try {
    const { project_id, name, priority, due, milestone_id, recurring_pattern, assignee_id, assignee_name, assignee_avatar, start_date } = req.body;
    if (tooLong(name, 150)) {
      return res.status(400).json({ error: 'Task name must be 150 characters or fewer' });
    }
    if (start_date && due && start_date > due) {
      return res.status(400).json({ error: 'Start date cannot be after the due date' });
    }
    const id = 't_' + Date.now();
    await dbRun(
      'INSERT INTO tasks (id, project_id, name, priority, due, status, milestone_id, recurring_pattern, assignee_id, assignee_name, assignee_avatar, start_date) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
      [id, project_id, name, priority || 'medium', due || '', 'todo', milestone_id || '', recurring_pattern || 'none', assignee_id || '', assignee_name || 'Unassigned', assignee_avatar || '#8b949e', start_date || '']
    );
    await createNotification(`Task created: ${name} (Assigned to ${assignee_name || 'Unassigned'})`);
    res.json({ id, project_id, name, priority: priority || 'medium', due: due || '', status: 'todo', milestone_id: milestone_id || '', recurring_pattern: recurring_pattern || 'none', assignee_id: assignee_id || '', assignee_name: assignee_name || 'Unassigned', assignee_avatar: assignee_avatar || '#8b949e', start_date: start_date || '' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/tasks/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { name, priority, due, status, milestone_id, recurring_pattern, assignee_id, assignee_name, assignee_avatar, start_date } = req.body;
    const task = await dbGet('SELECT * FROM tasks WHERE id = ?', [id]);
    if (!task) return res.status(404).json({ error: 'Task not found' });

    if (tooLong(name, 150)) {
      return res.status(400).json({ error: 'Task name must be 150 characters or fewer' });
    }

    const finalStartDate = start_date !== undefined ? start_date : task.start_date;
    const finalDue = due !== undefined ? due : task.due;
    if (finalStartDate && finalDue && finalStartDate > finalDue) {
      return res.status(400).json({ error: 'Start date cannot be after the due date' });
    }

    await dbRun(
      `UPDATE tasks SET name = ?, priority = ?, due = ?, status = ?, milestone_id = ?, recurring_pattern = ?, assignee_id = ?, assignee_name = ?, assignee_avatar = ?, start_date = ? WHERE id = ?`,
      [
        name !== undefined ? name : task.name,
        priority !== undefined ? priority : task.priority,
        finalDue,
        status !== undefined ? status : task.status,
        milestone_id !== undefined ? milestone_id : task.milestone_id,
        recurring_pattern !== undefined ? recurring_pattern : task.recurring_pattern,
        assignee_id !== undefined ? assignee_id : task.assignee_id,
        assignee_name !== undefined ? assignee_name : task.assignee_name,
        assignee_avatar !== undefined ? assignee_avatar : task.assignee_avatar,
        finalStartDate,
        id
      ]
    );
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/tasks/:id', async (req, res) => {
  try {
    const { id } = req.params;
    await dbRun('DELETE FROM tasks WHERE id = ?', [id]);
    await dbRun('DELETE FROM comments WHERE task_id = ?', [id]);
    await dbRun('DELETE FROM attachments WHERE task_id = ?', [id]);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/tasks/:id/status', async (req, res) => {
  try {
    const { id } = req.params;
    const { status } = req.body;

    const task = await dbGet('SELECT * FROM tasks WHERE id = ?', [id]);
    if (!task) return res.status(404).json({ error: 'Task not found' });

    await dbRun('UPDATE tasks SET status = ? WHERE id = ?', [status, id]);

    if (status === 'completed' && task.recurring_pattern && task.recurring_pattern !== 'none') {
      await handleTaskRecurrence(task);
    }

    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Comments ──
app.get('/api/tasks/:id/comments', async (req, res) => {
  try {
    const { id } = req.params;
    const comments = await dbAll('SELECT * FROM comments WHERE task_id = ? ORDER BY created_at', [id]);
    res.json(comments);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/tasks/:id/comments', async (req, res) => {
  try {
    const { id } = req.params;
    const { user_name, content } = req.body;
    const commentId = 'c_' + Date.now() + Math.round(Math.random() * 1e4);
    const created_at = new Date().toISOString();

    await dbRun(
      'INSERT INTO comments (id, task_id, user_name, content, created_at) VALUES (?, ?, ?, ?, ?)',
      [commentId, id, user_name, content, created_at]
    );
    await createNotification(`New comment by ${user_name}`);
    res.json({ id: commentId, task_id: id, user_name, content, created_at });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Attachments ──
app.get('/api/tasks/:id/attachments', async (req, res) => {
  try {
    const { id } = req.params;
    const attachments = await dbAll('SELECT * FROM attachments WHERE task_id = ?', [id]);
    res.json(attachments);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Issues a short-lived, scoped token so the browser can upload the file bytes directly
// to Vercel Blob, bypassing this server (Vercel Functions cap request bodies at 4.5MB,
// well under this app's 20MB attachment limit, so files can never be routed through here).
app.post('/api/uploads/token', async (req, res) => {
  try {
    const jsonResponse = await handleUpload({
      body: req.body,
      request: req,
      onBeforeGenerateToken: async (pathname) => {
        const ext = path.extname(pathname).toLowerCase();
        if (BLOCKED_UPLOAD_EXTENSIONS.has(ext)) {
          throw new Error(`File type "${ext}" is not allowed`);
        }
        return {
          addRandomSuffix: true,
          maximumSizeInBytes: 20 * 1024 * 1024 // 20MB per file
        };
      }
    });
    res.json(jsonResponse);
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

// TEMPORARY diagnostic route — remove after debugging the "access denied" issue.
app.get('/api/uploads/diag', async (req, res) => {
  const rw = process.env.BLOB_READ_WRITE_TOKEN || '';
  const rwParts = rw.split('_');
  res.json({
    rwTokenPresent: !!rw,
    rwTokenPrefix: rwParts.slice(0, 3).join('_'), // e.g. "vercel_blob_rw" — no secret revealed
    storeIdFromRwToken: rwParts[3] || null,
    blobStoreIdEnv: process.env.BLOB_STORE_ID || null,
    match: (rwParts[3] || null) === (process.env.BLOB_STORE_ID || null)
  });
});

// Records an attachment already uploaded straight to Vercel Blob via the token above.
app.post('/api/tasks/:id/attachments', async (req, res) => {
  try {
    const { id } = req.params;
    const { filename, filepath } = req.body;
    if (!filename || !filepath) {
      return res.status(400).json({ error: 'filename and filepath are required' });
    }

    const attachmentId = 'a_' + Date.now() + Math.round(Math.random() * 1e4);
    const uploaded_at = new Date().toISOString();

    await dbRun(
      'INSERT INTO attachments (id, task_id, filename, filepath, uploaded_at) VALUES (?, ?, ?, ?, ?)',
      [attachmentId, id, filename, filepath, uploaded_at]
    );
    await createNotification(`File attached: ${filename}`);
    res.json({ id: attachmentId, task_id: id, filename, filepath, uploaded_at });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Notifications ──
app.get('/api/notifications', async (req, res) => {
  try {
    const notifications = await dbAll('SELECT * FROM notifications ORDER BY created_at DESC LIMIT 20');
    res.json(notifications);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/notifications/unread-count', async (req, res) => {
  try {
    const row = await dbGet('SELECT COUNT(*) as count FROM notifications WHERE is_read = 0');
    res.json({ count: row ? row.count : 0 });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/notifications/read', async (req, res) => {
  try {
    await dbRun('UPDATE notifications SET is_read = 1 WHERE is_read = 0');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Milestones ──
app.get('/api/milestones', async (req, res) => {
  try {
    const { project_id } = req.query;
    let milestones;
    if (project_id) {
      milestones = await dbAll('SELECT * FROM milestones WHERE project_id = ?', [project_id]);
    } else {
      milestones = await dbAll('SELECT * FROM milestones');
    }
    res.json(milestones);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/milestones', async (req, res) => {
  try {
    const { project_id, name, date } = req.body;
    if (tooLong(name, 100)) {
      return res.status(400).json({ error: 'Milestone name must be 100 characters or fewer' });
    }
    const id = 'm_' + Date.now();
    await dbRun(
      'INSERT INTO milestones (id, project_id, name, date, completed) VALUES (?, ?, ?, ?, 0)',
      [id, project_id, name, date]
    );
    res.json({ id, project_id, name, date, completed: 0 });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/milestones/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { name, date } = req.body;
    const milestone = await dbGet('SELECT * FROM milestones WHERE id = ?', [id]);
    if (!milestone) return res.status(404).json({ error: 'Milestone not found' });
    if (tooLong(name, 100)) {
      return res.status(400).json({ error: 'Milestone name must be 100 characters or fewer' });
    }
    await dbRun(
      'UPDATE milestones SET name = ?, date = ? WHERE id = ?',
      [name !== undefined ? name : milestone.name, date !== undefined ? date : milestone.date, id]
    );
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/milestones/:id/status', async (req, res) => {
  try {
    const { id } = req.params;
    const { completed } = req.body;
    await dbRun('UPDATE milestones SET completed = ? WHERE id = ?', [completed, id]);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Timesheets ──
app.get('/api/timesheets', async (req, res) => {
  try {
    const { project_id } = req.query;
    let timesheets;
    if (project_id) {
      timesheets = await dbAll('SELECT * FROM timesheets WHERE project_id = ?', [project_id]);
    } else {
      timesheets = await dbAll('SELECT * FROM timesheets');
    }
    res.json(timesheets);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/timesheets', async (req, res) => {
  try {
    const { project_id, task_id, hours, date, notes, user } = req.body;
    const id = 'ts_' + Date.now();
    await dbRun(
      'INSERT INTO timesheets (id, project_id, task_id, hours, date, notes, user, billed, invoice_id) VALUES (?, ?, ?, ?, ?, ?, ?, 0, ?)',
      [id, project_id, task_id, hours, date, notes || '', user || '', '']
    );
    res.json({ id, project_id, task_id, hours, date, notes: notes || '', user: user || '', billed: 0, invoice_id: '' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/timesheets/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const timesheet = await dbGet('SELECT * FROM timesheets WHERE id = ?', [id]);
    if (!timesheet) return res.status(404).json({ error: 'Time entry not found' });
    if (timesheet.billed) {
      return res.status(400).json({ error: 'Cannot edit a time entry that has already been billed on an invoice' });
    }
    const { task_id, hours, date, notes } = req.body;
    if (tooLong(notes, 400)) {
      return res.status(400).json({ error: 'Notes must be 400 characters or fewer' });
    }
    await dbRun(
      'UPDATE timesheets SET task_id = ?, hours = ?, date = ?, notes = ? WHERE id = ?',
      [
        task_id !== undefined ? task_id : timesheet.task_id,
        hours !== undefined ? hours : timesheet.hours,
        date !== undefined ? date : timesheet.date,
        notes !== undefined ? notes : timesheet.notes,
        id
      ]
    );
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Invoices ──
app.get('/api/invoices', async (req, res) => {
  try {
    const { project_id } = req.query;
    let invoices;
    if (project_id) {
      invoices = await dbAll('SELECT * FROM invoices WHERE project_id = ? ORDER BY date DESC', [project_id]);
    } else {
      invoices = await dbAll('SELECT * FROM invoices ORDER BY date DESC');
    }
    res.json(invoices);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/invoices/unbilled-timesheets', async (req, res) => {
  try {
    const { project_id } = req.query;
    const timesheets = await dbAll(
      `SELECT ts.*, t.name as task_name 
       FROM timesheets ts 
       LEFT JOIN tasks t ON ts.task_id = t.id 
       WHERE ts.project_id = ? AND ts.billed = 0 ORDER BY ts.date ASC`,
      [project_id]
    );
    res.json(timesheets);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/invoices', async (req, res) => {
  try {
    const { project_id, client, client_address, date, due_date, items, timesheet_ids, tax_rate, discount, notes } = req.body;

    let lineItems = items || [];
    let totalHours = 0;
    let subtotal = 0;

    const wsRow = await dbGet('SELECT w.hourly_rate FROM projects p LEFT JOIN workspaces w ON p.workspace_id = w.id WHERE p.id = ?', [project_id]);
    const wsRate = (wsRow && wsRow.hourly_rate) ? parseFloat(wsRow.hourly_rate) : 50;

    if (lineItems.length === 0 && timesheet_ids && timesheet_ids.length > 0) {
      // Auto build line items from timesheets
      for (const tsId of timesheet_ids) {
        const ts = await dbGet('SELECT ts.*, t.name as task_name FROM timesheets ts LEFT JOIN tasks t ON ts.task_id = t.id WHERE ts.id = ?', [tsId]);
        if (ts) {
          const hours = ts.hours || 0;
          const rate = wsRate;
          const amt = hours * rate;
          lineItems.push({ desc: `${ts.task_name || 'Project Service'} (${ts.date}) - ${ts.notes || ''}`, hours, rate, amount: amt });
        }
      }
    }

    if (lineItems.length === 0) {
      return res.status(400).json({ error: 'Please add at least one line item or select timesheets to bill.' });
    }

    lineItems.forEach(item => {
      totalHours += parseFloat(item.hours || 0);
      subtotal += parseFloat(item.amount || 0);
    });

    const disc = parseFloat(discount || 0);
    const taxR = parseFloat(tax_rate || 0);
    const taxable = Math.max(0, subtotal - disc);
    const taxAmt = taxable * (taxR / 100);
    const grandTotal = taxable + taxAmt;

    const id = 'inv_' + Date.now();
    const invoiceNumber = 'INV-' + new Date().getFullYear() + '-' + String(Date.now()).slice(-3);
    const issueDate = date || new Date().toISOString().split('T')[0];
    const dueDate = due_date || new Date(Date.now() + 15 * 86400000).toISOString().split('T')[0];

    await dbRun(
      `INSERT INTO invoices (id, project_id, invoice_number, client, client_address, date, due_date, status, amount, tax_rate, tax_amount, discount, total_amount, hours, items, notes, paid_date)
       VALUES (?, ?, ?, ?, ?, ?, ?, 'unpaid', ?, ?, ?, ?, ?, ?, ?, ?, '')`,
      [id, project_id, invoiceNumber, client || 'Valued Client', client_address || '', issueDate, dueDate, subtotal, taxR, taxAmt, disc, grandTotal, totalHours, JSON.stringify(lineItems), notes || '']
    );

    // Mark timesheets as billed if provided
    if (timesheet_ids && Array.isArray(timesheet_ids)) {
      for (const tsId of timesheet_ids) {
        await dbRun('UPDATE timesheets SET billed = 1, invoice_id = ? WHERE id = ?', [id, tsId]);
      }
    }

    await createNotification(`Invoice generated: ${invoiceNumber} (${client}) for $${grandTotal.toFixed(2)}`);

    res.json({ id, project_id, invoice_number: invoiceNumber, client, status: 'unpaid', amount: subtotal, total_amount: grandTotal });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/invoices/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const invoice = await dbGet('SELECT * FROM invoices WHERE id = ?', [id]);
    if (!invoice) return res.status(404).json({ error: 'Invoice not found' });

    const { client, client_address, date, due_date, items, tax_rate, discount, notes } = req.body;

    const lineItems = Array.isArray(items) ? items : JSON.parse(invoice.items || '[]');
    if (lineItems.length === 0) {
      return res.status(400).json({ error: 'Invoice must have at least one line item.' });
    }
    if (tooLong(client, 150) || tooLong(client_address, 300) || tooLong(notes, 500)) {
      return res.status(400).json({ error: 'Client name, address, or notes is too long' });
    }

    let totalHours = 0;
    let subtotal = 0;
    lineItems.forEach(item => {
      totalHours += parseFloat(item.hours || 0);
      subtotal += parseFloat(item.amount || 0);
    });

    const disc = parseFloat(discount !== undefined ? discount : invoice.discount) || 0;
    const taxR = parseFloat(tax_rate !== undefined ? tax_rate : invoice.tax_rate) || 0;
    const taxable = Math.max(0, subtotal - disc);
    const taxAmt = taxable * (taxR / 100);
    const grandTotal = taxable + taxAmt;

    await dbRun(
      `UPDATE invoices SET client = ?, client_address = ?, date = ?, due_date = ?, amount = ?, tax_rate = ?, tax_amount = ?, discount = ?, total_amount = ?, hours = ?, items = ?, notes = ? WHERE id = ?`,
      [
        client !== undefined ? client : invoice.client,
        client_address !== undefined ? client_address : invoice.client_address,
        date !== undefined ? date : invoice.date,
        due_date !== undefined ? due_date : invoice.due_date,
        subtotal,
        taxR,
        taxAmt,
        disc,
        grandTotal,
        totalHours,
        JSON.stringify(lineItems),
        notes !== undefined ? notes : invoice.notes,
        id
      ]
    );
    res.json({ success: true, amount: subtotal, total_amount: grandTotal });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/invoices/:id/status', async (req, res) => {
  try {
    const { id } = req.params;
    const { status } = req.body;
    const paidDate = status === 'paid' ? new Date().toISOString().split('T')[0] : '';
    await dbRun('UPDATE invoices SET status = ?, paid_date = ? WHERE id = ?', [status, paidDate, id]);
    res.json({ success: true, status, paid_date: paidDate });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/invoices/:id', async (req, res) => {
  try {
    const { id } = req.params;
    // Release linked timesheets
    await dbRun("UPDATE timesheets SET billed = 0, invoice_id = '' WHERE invoice_id = ?", [id]);
    await dbRun('DELETE FROM invoices WHERE id = ?', [id]);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const exportInvoicesHandler = async (req, res) => {
  try {
    const { project_id } = req.query;
    let query = 'SELECT * FROM invoices';
    const params = [];
    if (project_id) {
      query += ' WHERE project_id = ?';
      params.push(project_id);
    }
    query += ' ORDER BY date DESC';
    const invoices = await dbAll(query, params);
    let csv = '\uFEFFInvoice Number,Client,Issue Date,Due Date,Status,Subtotal ($),Discount ($),Tax Amount ($),Grand Total ($),Paid Date\n';
    for (const inv of invoices) {
      const row = [
        `"${inv.invoice_number}"`,
        `"${(inv.client || '').replace(/"/g, '""')}"`,
        `"${inv.date}"`,
        `"${inv.due_date || ''}"`,
        `"${inv.status || 'unpaid'}"`,
        inv.amount || 0,
        inv.discount || 0,
        inv.tax_amount || 0,
        inv.total_amount || inv.amount || 0,
        `"${inv.paid_date || ''}"`
      ];
      csv += row.join(',') + '\n';
    }
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', 'attachment; filename="invoices_export.csv"');
    res.send(csv);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
app.get('/api/export/invoices', exportInvoicesHandler);
app.get('/api/invoices/export', exportInvoicesHandler);

// ── Export Tasks as CSV ──
const exportTasksHandler = async (req, res) => {
  try {
    const { project_id } = req.query;
    let query = `SELECT t.id, t.name, t.priority, t.due, t.status, t.recurring_pattern, m.name as milestone_name
                 FROM tasks t
                 LEFT JOIN milestones m ON t.milestone_id = m.id`;
    const params = [];
    if (project_id) {
      query += ' WHERE t.project_id = ?';
      params.push(project_id);
    }
    const tasks = await dbAll(query, params);

    let csv = '\uFEFFID,Name,Priority,Due Date,Status,Milestone,Recurring\n';
    for (const t of tasks) {
      const row = [
        t.id,
        `"${(t.name || '').replace(/"/g, '""')}"`,
        t.priority,
        t.due,
        t.status,
        `"${(t.milestone_name || 'None').replace(/"/g, '""')}"`,
        t.recurring_pattern || 'none'
      ];
      csv += row.join(',') + '\n';
    }

    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', 'attachment; filename="tasks_export.csv"');
    res.send(csv);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
app.get('/api/export/tasks', exportTasksHandler);
app.get('/api/tasks/export', exportTasksHandler);

// ──────────────────────────────────────────────
// Start Server
// ──────────────────────────────────────────────
// On Vercel the platform itself invokes the exported app per-request (no listener);
// everywhere else (local dev, Docker/Northflank) it needs to listen on a real port.
if (process.env.VERCEL !== '1') {
  app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
  });
}

export default app;
