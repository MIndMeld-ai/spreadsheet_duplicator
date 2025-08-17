const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electron', {
  showSaveDialog: (defaultPath) => ipcRenderer.invoke('show-save-dialog', defaultPath),
  writeFile: (filepath, buffer) => ipcRenderer.invoke('write-file', filepath, buffer),
  chooseFolder: () => ipcRenderer.invoke('choose-folder'),
  writeFileInFolder: (folderPath, filename, buffer) => ipcRenderer.invoke('write-file-in-folder', folderPath, filename, buffer)
});
