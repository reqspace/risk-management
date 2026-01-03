const { app, BrowserWindow, dialog, shell } = require('electron');
const { autoUpdater } = require('electron-updater');
const log = require('electron-log');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const http = require('http');

// Configure logging
log.transports.file.level = 'info';
autoUpdater.logger = log;

let mainWindow;
let pythonProcess;
const PYTHON_PORT = 5001;

// Get user data directory for storing config and data
function getUserDataPath() {
  return app.getPath('userData');
}

// Check if .env file exists, if not, prompt user to create one
async function ensureEnvFile() {
  const userDataPath = getUserDataPath();
  const envPath = path.join(userDataPath, '.env');

  if (!fs.existsSync(envPath)) {
    // Copy example env if it exists
    const exampleEnvPath = app.isPackaged
      ? path.join(process.resourcesPath, '.env.example')
      : path.join(__dirname, '..', '..', '.env.example');

    if (fs.existsSync(exampleEnvPath)) {
      fs.copyFileSync(exampleEnvPath, envPath);
    } else {
      // Create minimal .env
      fs.writeFileSync(envPath, 'ANTHROPIC_API_KEY=\n');
    }

    // Show setup dialog
    const result = await dialog.showMessageBox({
      type: 'info',
      title: 'First Run Setup',
      message: 'Welcome to Risk Management!',
      detail: `Please configure your API key.\n\nConfig file location:\n${envPath}\n\nYou need to add your Anthropic API key to this file.`,
      buttons: ['Open Config File', 'Continue Anyway']
    });

    if (result.response === 0) {
      shell.openPath(envPath);
    }
  }

  return envPath;
}

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
function startPythonServer(envPath) {
  return new Promise((resolve, reject) => {
    const pythonPath = getPythonPath();
    const userDataPath = getUserDataPath();

    // Environment variables for Python server
    const pythonEnv = {
      ...process.env,
      PORT: PYTHON_PORT.toString(),
      USER_DATA_PATH: userDataPath,
      DOTENV_PATH: envPath
    };

    if (pythonPath) {
      // Production: Run bundled executable
      log.info('Starting bundled Python server:', pythonPath);
      log.info('User data path:', userDataPath);
      pythonProcess = spawn(pythonPath, [], {
        env: pythonEnv,
        cwd: userDataPath,
        stdio: ['ignore', 'pipe', 'pipe'],
        detached: false
      });
    } else {
      // Development: Run Python script
      const scriptPath = path.join(__dirname, '..', '..', 'server.py');
      log.info('Starting Python server in dev mode:', scriptPath);
      pythonProcess = spawn('python3', [scriptPath], {
        env: pythonEnv,
        cwd: path.join(__dirname, '..', '..'),
        stdio: ['ignore', 'pipe', 'pipe']
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
  const maxAttempts = 60; // Wait up to 60 seconds for bundled Python to start

  // Use 127.0.0.1 explicitly to avoid IPv6 resolution issues
  http.get(`http://127.0.0.1:${PYTHON_PORT}/health`, (res) => {
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
      if (attempts % 10 === 0) {
        log.info(`Waiting for Python server... (attempt ${attempts + 1}/${maxAttempts})`);
      }
      setTimeout(() => waitForServer(resolve, reject, attempts + 1), 1000);
    } else {
      reject(new Error('Python server failed to start after 60 seconds'));
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
    // Ensure .env file exists (first run setup)
    const envPath = await ensureEnvFile();
    log.info('Using env file:', envPath);

    // Start Python server first
    await startPythonServer(envPath);

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
