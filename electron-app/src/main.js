const { app, BrowserWindow, dialog, shell } = require('electron');
const { autoUpdater } = require('electron-updater');
const log = require('electron-log');
const path = require('path');
const { spawn } = require('child_process');
const http = require('http');

// Configure logging
log.transports.file.level = 'info';
autoUpdater.logger = log;

let mainWindow;
let pythonProcess;
const PYTHON_PORT = 5001;

// Get the path to the Python executable
function getPythonPath() {
  const isProd = app.isPackaged;

  if (isProd) {
    // In production, Python is bundled in resources
    const resourcesPath = process.resourcesPath;
    if (process.platform === 'darwin') {
      return path.join(resourcesPath, 'python', 'risk_server');
    } else if (process.platform === 'win32') {
      return path.join(resourcesPath, 'python', 'risk_server.exe');
    }
  } else {
    // In development, use the Python script directly
    return null; // Will use system Python
  }
  return null;
}

// Start the Python backend server
function startPythonServer() {
  return new Promise((resolve, reject) => {
    const pythonPath = getPythonPath();

    if (pythonPath) {
      // Production: Run bundled executable
      log.info('Starting bundled Python server:', pythonPath);
      pythonProcess = spawn(pythonPath, [], {
        env: { ...process.env, PORT: PYTHON_PORT.toString() }
      });
    } else {
      // Development: Run Python script
      const scriptPath = path.join(__dirname, '..', '..', 'server.py');
      log.info('Starting Python server in dev mode:', scriptPath);
      pythonProcess = spawn('python3', [scriptPath], {
        env: { ...process.env, PORT: PYTHON_PORT.toString() },
        cwd: path.join(__dirname, '..', '..')
      });
    }

    pythonProcess.stdout.on('data', (data) => {
      log.info(`Python: ${data}`);
    });

    pythonProcess.stderr.on('data', (data) => {
      log.error(`Python Error: ${data}`);
    });

    pythonProcess.on('error', (err) => {
      log.error('Failed to start Python server:', err);
      reject(err);
    });

    pythonProcess.on('close', (code) => {
      log.info(`Python server exited with code ${code}`);
    });

    // Wait for server to be ready
    waitForServer(resolve, reject);
  });
}

// Poll until the Python server is responding
function waitForServer(resolve, reject, attempts = 0) {
  const maxAttempts = 30;

  http.get(`http://localhost:${PYTHON_PORT}/health`, (res) => {
    if (res.statusCode === 200) {
      log.info('Python server is ready');
      resolve();
    } else {
      retry();
    }
  }).on('error', () => {
    retry();
  });

  function retry() {
    if (attempts < maxAttempts) {
      setTimeout(() => waitForServer(resolve, reject, attempts + 1), 500);
    } else {
      reject(new Error('Python server failed to start'));
    }
  }
}

// Create the main application window
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1000,
    minHeight: 700,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true
    },
    titleBarStyle: 'hiddenInset',
    show: false
  });

  // Load the React app
  const isProd = app.isPackaged;
  if (isProd) {
    mainWindow.loadFile(path.join(__dirname, '..', 'build', 'webapp', 'index.html'));
  } else {
    // In dev, load from Vite dev server
    mainWindow.loadURL('http://localhost:3000');
  }

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();

    // Check for updates after window shows
    if (app.isPackaged) {
      autoUpdater.checkForUpdatesAndNotify();
    }
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  // Open external links in browser
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
}

// Auto-updater events
autoUpdater.on('checking-for-update', () => {
  log.info('Checking for update...');
});

autoUpdater.on('update-available', (info) => {
  log.info('Update available:', info.version);
  dialog.showMessageBox(mainWindow, {
    type: 'info',
    title: 'Update Available',
    message: `Version ${info.version} is available. It will be downloaded in the background.`,
    buttons: ['OK']
  });
});

autoUpdater.on('update-not-available', () => {
  log.info('Update not available - running latest version');
});

autoUpdater.on('download-progress', (progressObj) => {
  log.info(`Download speed: ${progressObj.bytesPerSecond} - Downloaded ${progressObj.percent}%`);
});

autoUpdater.on('update-downloaded', (info) => {
  log.info('Update downloaded:', info.version);
  dialog.showMessageBox(mainWindow, {
    type: 'info',
    title: 'Update Ready',
    message: `Version ${info.version} has been downloaded. Restart now to install?`,
    buttons: ['Restart', 'Later']
  }).then((result) => {
    if (result.response === 0) {
      autoUpdater.quitAndInstall();
    }
  });
});

autoUpdater.on('error', (err) => {
  log.error('Auto-updater error:', err);
});

// App lifecycle
app.whenReady().then(async () => {
  try {
    // Start Python server first
    await startPythonServer();

    // Then create window
    createWindow();
  } catch (err) {
    log.error('Failed to start application:', err);
    dialog.showErrorBox('Startup Error',
      'Failed to start the application server. Please try again or contact support.');
    app.quit();
  }
});

app.on('window-all-closed', () => {
  // Kill Python server
  if (pythonProcess) {
    pythonProcess.kill();
  }

  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

app.on('before-quit', () => {
  if (pythonProcess) {
    pythonProcess.kill();
  }
});
