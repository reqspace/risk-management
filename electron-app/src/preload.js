const { contextBridge, ipcRenderer } = require('electron');

// Expose Microsoft Auth functions to renderer
contextBridge.exposeInMainWorld('electronAPI', {
  // Microsoft Auth
  signInWithMicrosoft: () => ipcRenderer.invoke('microsoft-sign-in'),
  signOutMicrosoft: () => ipcRenderer.invoke('microsoft-sign-out'),
  getMicrosoftAccount: () => ipcRenderer.invoke('microsoft-get-account'),
  refreshMicrosoftToken: () => ipcRenderer.invoke('microsoft-refresh-token'),

  // App info
  getAppVersion: () => ipcRenderer.invoke('get-app-version'),

  // Check if running in Electron
  isElectron: true
});
