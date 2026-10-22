const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron');
const { spawn } = require('node:child_process');
const fs = require('node:fs/promises');
const net = require('node:net');
const os = require('node:os');
const path = require('node:path');
const { pathToFileURL } = require('node:url');

const STATUS_PAGE = path.join(__dirname, 'server-status.html');
const DEFAULT_PORT = Number(process.env.PORT || 3000);
let port = DEFAULT_PORT;
let statusWindow;
let serverProcess;
let shuttingDown = false;
let logStream;
let portConflict = false;

function networkAddresses(serverPort = port) {
  return [...new Set(Object.values(os.networkInterfaces()).flatMap(records =>
    (records || [])
      .filter(record => record.family === 'IPv4' && !record.internal)
      .map(record => `http://${record.address}:${serverPort}`)
  ))];
}

function portInUseError(serverPort) {
  const error = new Error(`المنفذ ${serverPort} مستخدم بالفعل. أدخل منفذًا آخر للمتابعة.`);
  error.code = 'EADDRINUSE';
  return error;
}

function validatePort(value) {
  const serverPort = Number(value);
  if (!Number.isInteger(serverPort) || serverPort < 1 || serverPort > 65535) {
    throw new Error('أدخل رقم منفذ صحيحًا بين 1 و65535.');
  }
  return serverPort;
}

function isPortListening(serverPort, host) {
  return new Promise((resolve, reject) => {
    const socket = net.createConnection({ port: serverPort, host });
    socket.once('connect', () => {
      socket.destroy();
      resolve(true);
    });
    socket.once('timeout', () => {
      socket.destroy();
      resolve(false);
    });
    socket.once('error', error => {
      socket.destroy();
      if (['ECONNREFUSED', 'EHOSTUNREACH', 'ENETUNREACH', 'EADDRNOTAVAIL', 'ETIMEDOUT'].includes(error.code)) {
        resolve(false);
      } else {
        reject(error);
      }
    });
    socket.setTimeout(500);
  });
}

async function isPortAvailable(serverPort) {
  const hosts = new Set([
    '127.0.0.1',
    '::1',
    ...Object.values(os.networkInterfaces()).flatMap(records =>
      (records || [])
        .filter(record => record.family === 'IPv4' && !record.internal)
        .map(record => record.address)
    ),
  ]);
  for (const host of hosts) {
    if (await isPortListening(serverPort, host)) return false;
  }
  return new Promise((resolve, reject) => {
    const tester = net.createServer();
    tester.once('error', error => {
      if (error.code === 'EADDRINUSE') resolve(false);
      else reject(error);
    });
    tester.listen({ port: serverPort, host: '0.0.0.0', exclusive: true }, () => {
      tester.close(error => error ? reject(error) : resolve(true));
    });
  });
}

function portSettingsPath() {
  return path.join(app.getPath('userData'), 'server-settings.json');
}

async function readConfiguredPort() {
  if (process.env.PORT) return validatePort(process.env.PORT);
  try {
    const settings = JSON.parse(await fs.readFile(portSettingsPath(), 'utf8'));
    return validatePort(settings.port);
  } catch (error) {
    if (error.code === 'ENOENT') return DEFAULT_PORT;
    console.error('Unable to read saved server port:', error);
    return DEFAULT_PORT;
  }
}

async function saveConfiguredPort(serverPort) {
  const file = portSettingsPath();
  await fs.mkdir(path.dirname(file), { recursive: true });
  const temporary = `${file}.${process.pid}.tmp`;
  await fs.writeFile(temporary, JSON.stringify({ port: serverPort }, null, 2), 'utf8');
  await fs.rename(temporary, file);
}

function updateStatus(status, message) {
  if (statusWindow && !statusWindow.isDestroyed()) {
    statusWindow.webContents.send('server:status', {
      status,
      message,
      port,
      addresses: networkAddresses(),
    });
  }
}

function ensureStatusSender(event) {
  if (!statusWindow || event.sender !== statusWindow.webContents ||
      !event.senderFrame.url.startsWith(pathToFileURL(STATUS_PAGE).href)) {
    throw new Error('طلب غير مصرح به');
  }
}

ipcMain.handle('server:get-status', event => {
  ensureStatusSender(event);
  return {
    status: serverProcess && serverProcess.exitCode === null ? 'starting' : 'stopped',
    port,
    addresses: networkAddresses(),
  };
});

ipcMain.handle('server:open-address', async (event, address) => {
  ensureStatusSender(event);
  const allowed = new Set(['http://localhost:' + port, 'http://127.0.0.1:' + port, ...networkAddresses()]);
  if (!allowed.has(address)) throw new Error('عنوان الاتصال غير صحيح');
  await shell.openExternal(address);
});

ipcMain.handle('server:retry-port', async (event, value) => {
  ensureStatusSender(event);
  const requestedPort = validatePort(value);
  try {
    await startServer(requestedPort);
    await saveConfiguredPort(requestedPort);
    return { ok: true };
  } catch (error) {
    updateStatus(error.code === 'EADDRINUSE' ? 'port-in-use' : 'stopped', error.message);
    if (error.code === 'EADDRINUSE') return { ok: false };
    throw error;
  }
});

function createStatusWindow() {
  const window = new BrowserWindow({
    width: 620,
    height: 590,
    minWidth: 500,
    minHeight: 480,
    title: 'خادم التحكم في البوابات',
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'server-preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      webSecurity: true,
    },
  });
  statusWindow = window;
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  window.on('closed', () => {
    if (statusWindow === window) statusWindow = null;
  });
  const addresses = networkAddresses();
  window.loadFile(STATUS_PAGE, {
    query: { port: String(port), addresses: JSON.stringify(addresses) },
  });
  return window;
}

async function waitUntilReady(serverPort) {
  const healthUrl = `http://127.0.0.1:${serverPort}/api/auth/status`;
  for (let attempt = 0; attempt < 40; attempt++) {
    if (!serverProcess || serverProcess.exitCode !== null) {
      if (portConflict) throw portInUseError(serverPort);
      throw new Error(`انتهى الخادم قبل بدء الاستماع. راجع سجل الخادم في مجلد بيانات التطبيق.`);
    }
    try {
      const response = await fetch(healthUrl, { signal: AbortSignal.timeout(1000) });
      if (response.ok) return;
    } catch {}
    await new Promise(resolve => setTimeout(resolve, 500));
  }
  if (portConflict) throw portInUseError(serverPort);
  throw new Error(`لم يستجب الخادم على المنفذ ${serverPort}. راجع سجل الخادم في مجلد بيانات التطبيق.`);
}

async function startServer(serverPort) {
  if (serverProcess && serverProcess.exitCode === null) {
    throw new Error('الخادم يعمل بالفعل.');
  }
  port = validatePort(serverPort);
  portConflict = false;
  if (!await isPortAvailable(port)) throw portInUseError(port);
  updateStatus('starting', '');
  const dataDirectory = app.getPath('userData');
  const logDirectory = path.join(dataDirectory, 'logs');
  await fs.mkdir(logDirectory, { recursive: true });
  logStream = require('node:fs').createWriteStream(path.join(logDirectory, 'server.log'), { flags: 'a' });
  const executable = path.join(process.resourcesPath, 'node-runtime', 'node.exe');
  const runner = path.join(process.resourcesPath, 'app.asar.unpacked', 'desktop', 'server-runner.js');
  const databasePath = path.join(dataDirectory, 'gate.db');
  await fs.access(executable);
  await fs.access(runner);
  serverProcess = spawn(executable, [runner], {
    env: {
      ...process.env,
      PORT: String(port),
      GATE_DB_PATH: databasePath,
      PORTABLE_EXECUTABLE_DIR: process.env.PORTABLE_EXECUTABLE_DIR || path.dirname(process.execPath),
    },
    windowsHide: true,
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  serverProcess.stdout.pipe(logStream, { end: false });
  serverProcess.stderr.on('data', chunk => {
    logStream.write(chunk);
    if (chunk.toString().includes(`SERVER_PORT_IN_USE:${port}`)) portConflict = true;
  });
  serverProcess.once('error', error => {
    console.error('Unable to launch server process:', error);
    updateStatus('stopped', error.message);
  });
  serverProcess.once('exit', (code, signal) => {
    if (!shuttingDown) {
      const message = portConflict
        ? `المنفذ ${port} مستخدم بالفعل. أدخل منفذًا آخر للمتابعة.`
        : `توقف الخادم (الرمز ${code ?? signal}). راجع سجل الخادم في مجلد بيانات التطبيق.`;
      console.error(message);
      updateStatus(portConflict ? 'port-in-use' : 'stopped', message);
    }
  });
  await waitUntilReady(port);
  updateStatus('running', '');
}

app.whenReady().then(async () => {
  const window = createStatusWindow();
  try {
    port = await readConfiguredPort();
    await startServer(port);
  } catch (error) {
    console.error('Unable to start the central gate server:', error);
    updateStatus(error.code === 'EADDRINUSE' ? 'port-in-use' : 'stopped', error.message);
    if (error.code !== 'EADDRINUSE' && !window.isDestroyed()) {
      dialog.showMessageBox(window, {
        type: 'error',
        title: 'تعذّر تشغيل الخادم',
        message: error.message,
      }).catch(dialogError => console.error('Unable to show server error:', dialogError));
    }
  }
}).catch(error => {
  console.error('Unable to launch the server application:', error);
  dialog.showErrorBox('تعذّر تشغيل الخادم', error.message);
  app.quit();
});

app.on('before-quit', () => {
  shuttingDown = true;
  if (serverProcess && serverProcess.exitCode === null) serverProcess.kill();
  if (logStream) logStream.end();
});

app.on('window-all-closed', () => app.quit());
