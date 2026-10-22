const Database = require('better-sqlite3');
const fs = require('node:fs');
const path = require('node:path');

const appRoot = path.resolve(__dirname, '..');
const databasePath = path.resolve(process.env.GATE_DB_PATH);
const portableDirectory = process.env.PORTABLE_EXECUTABLE_DIR;

function existingDatabaseCandidates() {
  const directories = [
    portableDirectory,
    portableDirectory && path.dirname(portableDirectory),
    process.cwd(),
  ].filter(Boolean);
  return [...new Set(directories.map(directory => path.resolve(directory, 'gate.db')))]
    .filter(file => file !== databasePath && fs.existsSync(file));
}

function importExistingDatabase() {
  if (fs.existsSync(databasePath)) return;
  const sourcePath = existingDatabaseCandidates()[0];
  if (!sourcePath) return;
  const source = new Database(sourcePath, { readonly: true, fileMustExist: true });
  try {
    source.backup(databasePath);
    console.log(`Imported the existing gate database from ${sourcePath}`);
  } finally {
    source.close();
  }
}

try {
  fs.mkdirSync(path.dirname(databasePath), { recursive: true });
  importExistingDatabase();
  process.env.GATE_DB_PATH = databasePath;
  process.env.GATE_BACKUP_DIR = path.join(path.dirname(databasePath), 'backups');
  process.env.PORT = process.env.PORT || '3000';
  require(path.join(appRoot, 'server.js'));
} catch (error) {
  console.error('Unable to start the central gate server:', error);
  process.exitCode = 1;
}
