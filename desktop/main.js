const { app, BrowserWindow, dialog, ipcMain, Menu } = require('electron');
const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');

const CONFIG_PAGE = path.join(__dirname, 'connection.html');
const SETTINGS_FILE = () => path.join(app.getPath('userData'), 'settings.json');
let configWindow;
let appWindow;

function normalizeServerUrl(value) {
  let url;
  try {
    url = new URL(String(value || '').trim());
  } catch {
    throw new Error('أدخل عنوان الخادم بصيغة صحيحة، مثل http://192.168.1.10:3000');
  }
  if (!['http:', 'https:'].includes(url.protocol) || !url.hostname ||
      url.username || url.password || !['', '/'].includes(url.pathname) || url.search || url.hash) {
    throw new Error('استخدم عنوان خادم http أو https فقط دون مسار إضافي');
  }
  return url.origin;
}

async function readSettings() {
  let contents;
  try {
    contents = await fs.readFile(SETTINGS_FILE(), 'utf8');
  } catch (error) {
    if (error.code === 'ENOENT') return { serverUrl: '' };
    throw error;
  }
  try {
    const settings = JSON.parse(contents);
    return { serverUrl: normalizeServerUrl(settings.serverUrl) };
  } catch (error) {
    console.error('Desktop server settings are invalid:', error);
    return { serverUrl: '' };
  }
}

async function writeSettings(serverUrl) {
  const file = SETTINGS_FILE();
  await fs.mkdir(path.dirname(file), { recursive: true });
  const temporary = `${file}.${process.pid}.tmp`;
  await fs.writeFile(temporary, JSON.stringify({ serverUrl }, null, 2), 'utf8');
  await fs.rename(temporary, file);
}

function ensureConfigSender(event) {
  if (!configWindow || event.sender !== configWindow.webContents ||
      !event.senderFrame.url.startsWith(pathToFileURL(CONFIG_PAGE).href)) {
    throw new Error('طلب غير مصرح به');
  }
}

ipcMain.handle('desktop:get-settings', async event => {
  ensureConfigSender(event);
  return readSettings();
});

ipcMain.handle('desktop:save-server', async (event, value) => {
  ensureConfigSender(event);
  const serverUrl = normalizeServerUrl(value);
  await writeSettings(serverUrl);
  if (configWindow && !configWindow.isDestroyed()) configWindow.close();
  await createAppWindow(serverUrl);
  return { ok: true };
});

function createConfigWindow() {
  if (appWindow && !appWindow.isDestroyed()) {
    appWindow.close();
    appWindow = null;
  }
  if (configWindow && !configWindow.isDestroyed()) {
    configWindow.focus();
    return;
  }
  const window = new BrowserWindow({
    width: 560,
    height: 500,
    minWidth: 440,
    minHeight: 420,
    title: 'إعداد الاتصال بالخادم',
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      webSecurity: true,
    },
  });
  configWindow = window;
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  window.on('closed', () => {
    if (configWindow === window) configWindow = null;
  });
  window.loadFile(CONFIG_PAGE);
}

async function createAppWindow(serverUrl) {
  const trustedOrigin = normalizeServerUrl(serverUrl);
  if (appWindow && !appWindow.isDestroyed()) appWindow.close();
  const window = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 900,
    minHeight: 640,
    show: false,
    title: 'التحكم في البوابات',
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      webSecurity: true,
      allowRunningInsecureContent: false,
    },
  });
  appWindow = window;
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  const allowTrustedNavigation = (event, target) => {
    try {
      if (new URL(target).origin !== trustedOrigin) event.preventDefault();
    } catch {
      event.preventDefault();
    }
  };
  window.webContents.on('will-navigate', allowTrustedNavigation);
  window.webContents.on('will-redirect', allowTrustedNavigation);
  window.webContents.on('did-fail-load', (_event, code, description, _url, isMainFrame) => {
    if (isMainFrame && code !== -3) {
      console.error(`Unable to load central gate server (${code}): ${description}`);
      createConfigWindow();
    }
  });
  window.on('closed', () => {
    if (appWindow === window) appWindow = null;
  });
  await window.loadURL(trustedOrigin);
  if (!window.isDestroyed()) window.show();
}

function installMenu() {
  const menu = Menu.buildFromTemplate([
    {
      label: 'التطبيق',
      submenu: [
        { label: 'تغيير عنوان الخادم…', click: createConfigWindow },
        { type: 'separator' },
        { role: 'quit', label: 'إغلاق' },
      ],
    },
    {
      label: 'عرض',
      submenu: [
        { role: 'reload', label: 'إعادة تحميل' },
        { role: 'togglefullscreen', label: 'ملء الشاشة' },
      ],
    },
  ]);
  Menu.setApplicationMenu(menu);
}

app.whenReady().then(async () => {
  installMenu();
  const settings = await readSettings();
  if (settings.serverUrl) {
    try {
      await createAppWindow(settings.serverUrl);
    } catch (error) {
      console.error('Unable to connect to the configured gate server:', error);
      createConfigWindow();
    }
  } else {
    createConfigWindow();
  }
  app.on('activate', async () => {
    if (!BrowserWindow.getAllWindows().length) {
      const current = await readSettings();
      if (current.serverUrl) await createAppWindow(current.serverUrl);
      else createConfigWindow();
    }
  });
}).catch(error => {
  console.error('Unable to start Gate Control Desktop:', error);
  dialog.showErrorBox('تعذّر تشغيل التطبيق', error.message);
  app.quit();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
