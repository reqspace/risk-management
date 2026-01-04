const { contextBridge, ipcRenderer } = require('electron');

// Expose Auth functions to renderer
contextBridge.exposeInMainWorld('electronAPI', {
  // Microsoft Auth
  signInWithMicrosoft: () => ipcRenderer.invoke('microsoft-sign-in'),
  signOutMicrosoft: () => ipcRenderer.invoke('microsoft-sign-out'),
  getMicrosoftAccount: () => ipcRenderer.invoke('microsoft-get-account'),
  refreshMicrosoftToken: () => ipcRenderer.invoke('microsoft-refresh-token'),

  // Google Auth
  signInWithGoogle: () => ipcRenderer.invoke('google-sign-in'),
  signOutGoogle: () => ipcRenderer.invoke('google-sign-out'),
  getGoogleAccount: () => ipcRenderer.invoke('google-get-account'),

  // App info
  getAppVersion: () => ipcRenderer.invoke('get-app-version'),

  // Check if running in Electron
  isElectron: true
});
