const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');

let mainWindow;
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true
    }
  });
  mainWindow.loadFile('index.html');
}

app.whenReady().then(() => {
  createWindow();
  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});

// IPC handlers for saving files
ipcMain.handle('show-save-dialog', async (event, defaultPath) => {
  const res = await dialog.showSaveDialog({ defaultPath });
  return res;
});

ipcMain.handle('write-file', async (event, filepath, buffer) => {
  try {
    fs.writeFileSync(filepath, Buffer.from(buffer));
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
});

// New: choose output folder
ipcMain.handle('choose-folder', async (event) => {
  try {
    const res = await dialog.showOpenDialog({ properties: ['openDirectory'] });
    if (res.canceled) return { canceled: true };
    return { canceled: false, folder: res.filePaths && res.filePaths[0] };
  } catch (err) {
    return { canceled: true, error: err.message };
  }
});

// New: write file into a chosen folder (main joins path safely)
ipcMain.handle('write-file-in-folder', async (event, folderPath, filename, buffer) => {
  try {
    const full = path.join(folderPath, filename);
    fs.writeFileSync(full, Buffer.from(buffer));
    return { ok: true, path: full };
  } catch (err) {
    return { ok: false, error: err.message };
  }
});
