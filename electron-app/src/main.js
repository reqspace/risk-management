const { app, BrowserWindow, dialog, shell, ipcMain } = require('electron');
const { autoUpdater } = require('electron-updater');
const log = require('electron-log');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const http = require('http');
const { PublicClientApplication } = require('@azure/msal-node');
const { google } = require('googleapis');

// Configure logging
log.transports.file.level = 'info';
autoUpdater.logger = log;

// Microsoft Auth Configuration
const MSAL_CONFIG = {
  auth: {
    clientId: 'd9260854-0354-44f3-b12a-6aab224803ff',
    authority: 'https://login.microsoftonline.com/14c97b03-9bb9-44ba-a4a6-e44e181ab35e',
    redirectUri: 'http://localhost'
  }
};

const MSAL_SCOPES = ['User.Read', 'Calendars.ReadWrite', 'Mail.Read', 'Mail.Send'];

let msalClient = null;
let msalTokenCache = null;

// Initialize MSAL
function initializeMsal() {
  msalClient = new PublicClientApplication(MSAL_CONFIG);
  log.info('MSAL initialized');
}

// Get cached accounts
async function getMsalAccounts() {
  if (!msalClient) return [];
  const cache = msalClient.getTokenCache();
  const accounts = await cache.getAllAccounts();
  return accounts;
}

// Sign in with Microsoft
async function signInWithMicrosoft() {
  if (!msalClient) initializeMsal();

  try {
    // Try silent auth first if we have cached accounts
    const accounts = await getMsalAccounts();
    if (accounts.length > 0) {
      try {
        const silentResult = await msalClient.acquireTokenSilent({
          account: accounts[0],
          scopes: MSAL_SCOPES
        });
        log.info('Silent auth successful for:', silentResult.account.username);
        return {
          success: true,
          account: silentResult.account,
          accessToken: silentResult.accessToken
        };
      } catch (silentError) {
        log.info('Silent auth failed, will try interactive');
      }
    }

    // Interactive auth with device code flow (works well for desktop apps)
    // Copy code to clipboard immediately
    const { clipboard } = require('electron');

    const deviceCodeRequest = {
      scopes: MSAL_SCOPES,
      deviceCodeCallback: (response) => {
        // Copy code to clipboard immediately
        clipboard.writeText(response.userCode);
        log.info('Device code:', response.userCode);

        // Show dialog to user with the code - keep showing until they act
        dialog.showMessageBox(mainWindow, {
          type: 'info',
          title: 'Sign in with Microsoft',
          message: `Your code: ${response.userCode}`,
          detail: `This code has been copied to your clipboard.\n\n1. Click "Open Browser" below\n2. Paste the code: ${response.userCode}\n3. Sign in with your Microsoft account\n4. Come back here - you'll be signed in automatically`,
          buttons: ['Open Browser', 'Copy Code Again', 'Cancel']
        }).then((result) => {
          if (result.response === 0) {
            // Open browser
            shell.openExternal(response.verificationUri);
          } else if (result.response === 1) {
            // Copy code again
            clipboard.writeText(response.userCode);
            shell.openExternal(response.verificationUri);
          }
        });
      }
    };

    const result = await msalClient.acquireTokenByDeviceCode(deviceCodeRequest);
    log.info('Microsoft sign-in successful:', result.account.username);

    // Save tokens to config for Python backend to use
    await saveMicrosoftTokens(result);

    return {
      success: true,
      account: result.account,
      accessToken: result.accessToken
    };
  } catch (error) {
    log.error('Microsoft sign-in error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// Save tokens to config file for Python backend
async function saveMicrosoftTokens(authResult) {
  const configPath = path.join(getAppConfigPath(), 'ms-tokens.json');
  const tokenData = {
    accessToken: authResult.accessToken,
    account: {
      username: authResult.account.username,
      name: authResult.account.name,
      homeAccountId: authResult.account.homeAccountId
    },
    expiresOn: authResult.expiresOn,
    scopes: authResult.scopes
  };
  fs.writeFileSync(configPath, JSON.stringify(tokenData, null, 2));
  log.info('Microsoft tokens saved');

  // Also update the .env file with the user email for Graph API
  const dataFolder = getDataFolderPath();
  const envPath = path.join(dataFolder, '.env');
  if (fs.existsSync(envPath)) {
    let envContent = fs.readFileSync(envPath, 'utf8');
    // Update MS_USER_EMAIL
    if (envContent.includes('MS_USER_EMAIL=')) {
      envContent = envContent.replace(/MS_USER_EMAIL=.*/, `MS_USER_EMAIL=${authResult.account.username}`);
    } else {
      envContent += `\nMS_USER_EMAIL=${authResult.account.username}`;
    }
    fs.writeFileSync(envPath, envContent);
  }
}

// Get current Microsoft account
async function getMicrosoftAccount() {
  const configPath = path.join(getAppConfigPath(), 'ms-tokens.json');
  if (fs.existsSync(configPath)) {
    try {
      const tokenData = JSON.parse(fs.readFileSync(configPath, 'utf8'));
      return {
        success: true,
        account: tokenData.account,
        isExpired: new Date(tokenData.expiresOn) < new Date()
      };
    } catch (e) {
      return { success: false };
    }
  }
  return { success: false };
}

// Sign out from Microsoft
async function signOutMicrosoft() {
  try {
    // Clear token cache
    const configPath = path.join(getAppConfigPath(), 'ms-tokens.json');
    if (fs.existsSync(configPath)) {
      fs.unlinkSync(configPath);
    }

    // Clear MSAL cache
    if (msalClient) {
      const cache = msalClient.getTokenCache();
      const accounts = await cache.getAllAccounts();
      for (const account of accounts) {
        await cache.removeAccount(account);
      }
    }

    log.info('Microsoft sign-out complete');
    return { success: true };
  } catch (error) {
    log.error('Sign-out error:', error);
    return { success: false, error: error.message };
  }
}

// Refresh Microsoft token
async function refreshMicrosoftToken() {
  if (!msalClient) initializeMsal();

  try {
    const accounts = await getMsalAccounts();
    if (accounts.length === 0) {
      return { success: false, error: 'No account found' };
    }

    const result = await msalClient.acquireTokenSilent({
      account: accounts[0],
      scopes: MSAL_SCOPES
    });

    await saveMicrosoftTokens(result);
    return {
      success: true,
      accessToken: result.accessToken
    };
  } catch (error) {
    log.error('Token refresh error:', error);
    return { success: false, error: error.message };
  }
}

// ============================================
// Google OAuth Configuration
// ============================================
const GOOGLE_CONFIG = {
  clientId: '10018350263-leqmet8aekiu05b42bq8qk6el7r87fr6.apps.googleusercontent.com',
  clientSecret: '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',
  redirectUri: 'http://localhost:8089/oauth2callback'
};

const GOOGLE_SCOPES = [
  'https://www.googleapis.com/auth/gmail.readonly',
  'https://www.googleapis.com/auth/gmail.send',
  'https://www.googleapis.com/auth/userinfo.email',
  'https://www.googleapis.com/auth/userinfo.profile'
];

let googleOAuth2Client = null;

// Initialize Google OAuth client
function initializeGoogle() {
  googleOAuth2Client = new google.auth.OAuth2(
    GOOGLE_CONFIG.clientId,
    GOOGLE_CONFIG.clientSecret,
    GOOGLE_CONFIG.redirectUri
  );
  log.info('Google OAuth initialized');
}

// Sign in with Google
async function signInWithGoogle() {
  if (!googleOAuth2Client) initializeGoogle();

  // Kill any existing server on the port first
  try {
    const { execSync } = require('child_process');
    execSync('lsof -ti:8089 | xargs kill -9 2>/dev/null || true');
  } catch (e) {
    // Ignore errors
  }

  return new Promise((resolve) => {
    // Create a local server to handle the OAuth callback
    const server = http.createServer(async (req, res) => {
      try {
        const url = new URL(req.url, 'http://localhost:8089');
        if (url.pathname === '/oauth2callback') {
          const code = url.searchParams.get('code');

          if (code) {
            // Exchange code for tokens
            const { tokens } = await googleOAuth2Client.getToken(code);
            googleOAuth2Client.setCredentials(tokens);

            // Get user info
            const oauth2 = google.oauth2({ version: 'v2', auth: googleOAuth2Client });
            const userInfo = await oauth2.userinfo.get();

            // Save tokens
            await saveGoogleTokens(tokens, userInfo.data);

            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(`
              <html>
                <body style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: #f0f0f0;">
                  <div style="text-align: center; padding: 40px; background: white; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
                    <h1 style="color: #4285f4;">✓ Signed in successfully!</h1>
                    <p>You can close this window and return to the app.</p>
                  </div>
                </body>
              </html>
            `);

            server.close();
            resolve({
              success: true,
              account: {
                email: userInfo.data.email,
                name: userInfo.data.name,
                picture: userInfo.data.picture
              }
            });
          } else {
            res.writeHead(400, { 'Content-Type': 'text/html' });
            res.end('<h1>Error: No authorization code received</h1>');
            server.close();
            resolve({ success: false, error: 'No authorization code' });
          }
        }
      } catch (error) {
        log.error('Google OAuth callback error:', error);
        res.writeHead(500, { 'Content-Type': 'text/html' });
        res.end(`<h1>Error: ${error.message}</h1>`);
        server.close();
        resolve({ success: false, error: error.message });
      }
    });

    server.listen(8089, () => {
      log.info('Google OAuth callback server listening on port 8089');

      // Generate auth URL and open in browser
      const authUrl = googleOAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: GOOGLE_SCOPES,
        prompt: 'consent'
      });

      shell.openExternal(authUrl);
    });

    // Timeout after 5 minutes
    setTimeout(() => {
      server.close();
      resolve({ success: false, error: 'Sign-in timeout' });
    }, 300000);
  });
}

// Save Google tokens
async function saveGoogleTokens(tokens, userInfo) {
  const configPath = path.join(getAppConfigPath(), 'google-tokens.json');
  const tokenData = {
    accessToken: tokens.access_token,
    refreshToken: tokens.refresh_token,
    expiryDate: tokens.expiry_date,
    account: {
      email: userInfo.email,
      name: userInfo.name,
      picture: userInfo.picture
    }
  };
  fs.writeFileSync(configPath, JSON.stringify(tokenData, null, 2));
  log.info('Google tokens saved for:', userInfo.email);
}

// Get current Google account
async function getGoogleAccount() {
  const configPath = path.join(getAppConfigPath(), 'google-tokens.json');
  if (fs.existsSync(configPath)) {
    try {
      const tokenData = JSON.parse(fs.readFileSync(configPath, 'utf8'));
      return {
        success: true,
        account: tokenData.account,
        isExpired: tokenData.expiryDate && Date.now() > tokenData.expiryDate
      };
    } catch (e) {
      return { success: false };
    }
  }
  return { success: false };
}

// Sign out from Google
async function signOutGoogle() {
  try {
    const configPath = path.join(getAppConfigPath(), 'google-tokens.json');
    if (fs.existsSync(configPath)) {
      fs.unlinkSync(configPath);
    }
    googleOAuth2Client = null;
    log.info('Google sign-out complete');
    return { success: true };
  } catch (error) {
    log.error('Google sign-out error:', error);
    return { success: false, error: error.message };
  }
}

let mainWindow;
let pythonProcess;
const PYTHON_PORT = 5001;

// Prevent multiple instances of the app
const gotTheLock = app.requestSingleInstanceLock();
if (!gotTheLock) {
  app.quit();
} else {
  // Focus existing window when second instance is launched
  app.on('second-instance', () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
    }
  });
}

// Get the app config directory (for storing settings only)
function getAppConfigPath() {
  return app.getPath('userData');
}

// Get the data folder path (user-selectable, can be OneDrive, etc.)
function getDataFolderPath() {
  const configPath = path.join(getAppConfigPath(), 'config.json');
  if (fs.existsSync(configPath)) {
    try {
      const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
      if (config.dataFolder && fs.existsSync(config.dataFolder)) {
        return config.dataFolder;
      }
    } catch (e) {
      log.error('Error reading config:', e);
    }
  }
  // Default to userData if no custom folder set
  return getAppConfigPath();
}

// Save data folder path to config
function saveDataFolderPath(folderPath) {
  const configPath = path.join(getAppConfigPath(), 'config.json');
  let config = {};
  if (fs.existsSync(configPath)) {
    try {
      config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
    } catch (e) {}
  }
  config.dataFolder = folderPath;
  fs.writeFileSync(configPath, JSON.stringify(config, null, 2));
  log.info('Saved data folder path:', folderPath);
}

// Check if this is the first run
function isFirstRun() {
  const configPath = path.join(getAppConfigPath(), 'config.json');
  return !fs.existsSync(configPath);
}

// Run the first-time setup wizard
async function runSetupWizard() {
  // Step 1: Welcome and explain what we need
  const welcomeResult = await dialog.showMessageBox({
    type: 'info',
    title: 'Welcome to Risk Management!',
    message: 'First-Time Setup',
    detail: `This wizard will help you set up the Risk Management app.

You'll need to:
1. Choose where to store your data files
   (You can use OneDrive, Dropbox, or any synced folder!)

2. Add your Anthropic API key
   (Get one free at console.anthropic.com)

3. Configure your project names

4. (Optional) Set up integrations

Click "Start Setup" to begin.`,
    buttons: ['Start Setup', 'Quit']
  });

  if (welcomeResult.response === 1) {
    app.quit();
    return null;
  }

  // Step 2: Choose data folder
  const defaultFolder = path.join(app.getPath('home'), 'Risk Management');
  const folderResult = await dialog.showMessageBox({
    type: 'question',
    title: 'Step 1: Choose Data Location',
    message: 'Where should we store your data?',
    detail: `Your risk registers, reports, and project files will be stored here.

TIP: Choose a OneDrive or Dropbox folder to automatically sync your data across devices!

Default location:
${defaultFolder}`,
    buttons: ['Browse for Folder...', 'Use Default Location', 'Cancel Setup']
  });

  let dataFolder;
  if (folderResult.response === 2) {
    app.quit();
    return null;
  } else if (folderResult.response === 0) {
    // User wants to browse
    const browseResult = await dialog.showOpenDialog({
      title: 'Select Data Folder',
      defaultPath: app.getPath('home'),
      properties: ['openDirectory', 'createDirectory'],
      message: 'Choose a folder for Risk Management data (e.g., OneDrive folder)'
    });

    if (browseResult.canceled || browseResult.filePaths.length === 0) {
      // Cancelled, use default
      dataFolder = defaultFolder;
    } else {
      dataFolder = browseResult.filePaths[0];
    }
  } else {
    // Use default
    dataFolder = defaultFolder;
  }

  // Ensure the folder exists
  if (!fs.existsSync(dataFolder)) {
    fs.mkdirSync(dataFolder, { recursive: true });
  }

  // Save the folder path
  saveDataFolderPath(dataFolder);
  log.info('Data folder set to:', dataFolder);

  // Step 3: Create .env file in the data folder
  const envPath = path.join(dataFolder, '.env');
  if (!fs.existsSync(envPath)) {
    const envTemplate = `# Risk Management System Configuration
# Data Location: ${dataFolder}

# ===========================================
# REQUIRED: Anthropic API Key
# ===========================================
# Get yours at: https://console.anthropic.com/
ANTHROPIC_API_KEY=

# ===========================================
# REQUIRED: Project Names
# ===========================================
# Comma-separated list of your project codes
# Example: PROJECT_NAMES=RH,NSD,HERC,BRGO,HB
PROJECT_NAMES=

# ===========================================
# OPTIONAL: ngrok (for remote access)
# ===========================================
# Get token at: https://dashboard.ngrok.com/
NGROK_AUTH_TOKEN=

# ===========================================
# OPTIONAL: Email Reports (Daily Digest)
# ===========================================
SMTP_SERVER=smtp.office365.com
SMTP_PORT=587
EMAIL_FROM=
EMAIL_TO=
EMAIL_PASSWORD=

# ===========================================
# OPTIONAL: Microsoft Graph API
# ===========================================
# For Outlook Calendar integration
# Setup: Azure Portal > App Registrations
MS_CLIENT_ID=
MS_TENANT_ID=
MS_CLIENT_SECRET=
MS_USER_EMAIL=

# ===========================================
# OPTIONAL: Email Reading (IMAP)
# ===========================================
# For processing email attachments
IMAP_SERVER=outlook.office365.com
IMAP_PORT=993
IMAP_PASSWORD=
`;
    fs.writeFileSync(envPath, envTemplate);
  }

  // Step 4: Prompt for API key and show next steps
  const setupComplete = await dialog.showMessageBox({
    type: 'info',
    title: 'Step 2: Configure Settings',
    message: 'Data folder ready!',
    detail: `Your data will be stored in:
${dataFolder}

NEXT STEPS:
1. Open the app's Settings page
2. Enter your Anthropic API key
3. Add your project names
4. (Optional) Configure integrations

You can also edit the config file directly if you prefer.`,
    buttons: ['Open Config File', 'Continue to App']
  });

  if (setupComplete.response === 0) {
    shell.openPath(envPath);
  }

  return envPath;
}

// Check if .env file exists, if not, run setup wizard
async function ensureEnvFile() {
  // Check if this is first run
  if (isFirstRun()) {
    return await runSetupWizard();
  }

  // Get the data folder path
  const dataFolder = getDataFolderPath();
  const envPath = path.join(dataFolder, '.env');

  // If .env doesn't exist in data folder, check app config folder for migration
  if (!fs.existsSync(envPath)) {
    const oldEnvPath = path.join(getAppConfigPath(), '.env');
    if (fs.existsSync(oldEnvPath)) {
      // Migrate old .env to new location
      fs.copyFileSync(oldEnvPath, envPath);
      log.info('Migrated .env from', oldEnvPath, 'to', envPath);
    } else {
      // No env file at all, run setup
      return await runSetupWizard();
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
    const dataFolderPath = getDataFolderPath();

    // Environment variables for Python server
    const pythonEnv = {
      ...process.env,
      PORT: PYTHON_PORT.toString(),
      USER_DATA_PATH: dataFolderPath,  // This is the user-selected data folder
      DOTENV_PATH: envPath
    };

    if (pythonPath) {
      // Production: Run bundled executable
      log.info('Starting bundled Python server:', pythonPath);
      log.info('Data folder path:', dataFolderPath);
      pythonProcess = spawn(pythonPath, [], {
        env: pythonEnv,
        cwd: dataFolderPath,
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
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    },
    // Use default title bar for proper Mac window behavior (dragging, minimize, etc.)
    titleBarStyle: 'default',
    // Add proper window title
    title: 'Risk Management',
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

// IPC Handlers for Microsoft Auth
ipcMain.handle('microsoft-sign-in', async () => {
  return await signInWithMicrosoft();
});

ipcMain.handle('microsoft-sign-out', async () => {
  return await signOutMicrosoft();
});

ipcMain.handle('microsoft-get-account', async () => {
  return await getMicrosoftAccount();
});

ipcMain.handle('microsoft-refresh-token', async () => {
  return await refreshMicrosoftToken();
});

ipcMain.handle('get-app-version', () => {
  return app.getVersion();
});

// IPC Handlers for Google Auth
ipcMain.handle('google-sign-in', async () => {
  return await signInWithGoogle();
});

ipcMain.handle('google-sign-out', async () => {
  return await signOutGoogle();
});

ipcMain.handle('google-get-account', async () => {
  return await getGoogleAccount();
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
