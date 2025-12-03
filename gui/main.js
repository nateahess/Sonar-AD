const { app, BrowserWindow, ipcMain, shell } = require('electron');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    minWidth: 600,
    minHeight: 500,
    backgroundColor: '#2c2c2c',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true
    },
    icon: path.join(__dirname, 'assets/icon.png')
  });

  mainWindow.loadFile('index.html');

  // Open DevTools in development (comment out for production)
  // mainWindow.webContents.openDevTools();

  mainWindow.on('closed', function () {
    mainWindow = null;
  });
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

// IPC Handler: Get AD Domain
ipcMain.handle('get-domain', async () => {
  return new Promise((resolve, reject) => {
    const psCommand = 'Get-ADDomain | Select-Object -ExpandProperty DNSRoot';

    const ps = spawn('powershell.exe', [
      '-NoProfile',
      '-NonInteractive',
      '-Command',
      psCommand
    ]);

    let output = '';
    let errorOutput = '';

    ps.stdout.on('data', (data) => {
      output += data.toString();
    });

    ps.stderr.on('data', (data) => {
      errorOutput += data.toString();
    });

    ps.on('close', (code) => {
      if (code === 0 && output.trim()) {
        resolve({ success: true, domain: output.trim() });
      } else {
        const errorMsg = errorOutput || 'Unable to retrieve domain information';
        resolve({
          success: false,
          error: errorMsg,
          message: 'Make sure you are connected to a domain and have the Active Directory module installed.'
        });
      }
    });

    ps.on('error', (err) => {
      reject({
        success: false,
        error: err.message,
        message: 'Failed to execute PowerShell command'
      });
    });
  });
});

// IPC Handler: Generate Report
ipcMain.handle('generate-report', async () => {
  return new Promise((resolve, reject) => {
    // Determine script path based on whether app is packaged or in development
    let scriptPath;
    if (app.isPackaged) {
      // In production, SonarAD.ps1 is in the resources folder
      scriptPath = path.join(process.resourcesPath, 'SonarAD.ps1');
    } else {
      // In development, script is in parent directory
      scriptPath = path.join(__dirname, '..', 'SonarAD.ps1');
    }

    // Verify script exists
    if (!fs.existsSync(scriptPath)) {
      resolve({
        success: false,
        error: `Script not found at: ${scriptPath}`,
        message: 'SonarAD.ps1 script is missing'
      });
      return;
    }

    // Determine working directory (where the report will be saved)
    const workingDir = app.isPackaged
      ? path.dirname(app.getPath('exe'))
      : path.join(__dirname, '..');

    const outputPath = path.join(workingDir, 'ADMetricsReport.html');

    const ps = spawn('powershell.exe', [
      '-NoProfile',
      '-ExecutionPolicy', 'Bypass',
      '-File', scriptPath,
      '-OutputPath', outputPath
    ], {
      cwd: workingDir
    });

    let output = '';
    let errorOutput = '';

    ps.stdout.on('data', (data) => {
      const text = data.toString();
      output += text;
      // Send real-time output to renderer
      if (mainWindow) {
        mainWindow.webContents.send('script-output', text);
      }
    });

    ps.stderr.on('data', (data) => {
      const text = data.toString();
      errorOutput += text;
      // Send error output to renderer
      if (mainWindow) {
        mainWindow.webContents.send('script-output', text);
      }
    });

    ps.on('close', (code) => {
      if (code === 0) {
        // Check if report file was created
        if (fs.existsSync(outputPath)) {
          resolve({
            success: true,
            message: 'Report generated successfully!',
            reportPath: outputPath,
            output: output
          });
        } else {
          resolve({
            success: false,
            error: 'Report file was not created',
            message: 'Script completed but report file not found',
            output: output
          });
        }
      } else {
        resolve({
          success: false,
          error: `PowerShell exited with code ${code}`,
          message: errorOutput || 'Script execution failed',
          output: output + '\n' + errorOutput
        });
      }
    });

    ps.on('error', (err) => {
      reject({
        success: false,
        error: err.message,
        message: 'Failed to execute PowerShell script'
      });
    });
  });
});

// IPC Handler: Open Report
ipcMain.handle('open-report', async (event, reportPath) => {
  try {
    if (fs.existsSync(reportPath)) {
      await shell.openPath(reportPath);
      return { success: true };
    } else {
      return {
        success: false,
        error: 'Report file not found'
      };
    }
  } catch (err) {
    return {
      success: false,
      error: err.message
    };
  }
});
