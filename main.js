const { app, BrowserWindow, ipcMain, dialog, clipboard, nativeImage, shell, desktopCapturer } = require('electron');
const path = require('path');
const fs = require('fs').promises;
const Store = require('electron-store');
const { exec } = require('child_process');
const { promisify } = require('util');
require('dotenv').config();

const execAsync = promisify(exec);
const store = new Store();

// Company API key from environment variable
const COMPANY_OPENAI_KEY = process.env.OPENAI_API_KEY;

// JSON storage paths in user folder
const USER_DATA_PATH = app.getPath('userData');
const FOLDER_STRUCTURE_PATH = path.join(USER_DATA_PATH, 'folder-structure.json');
const LAST_PROCESSED_PATH = path.join(USER_DATA_PATH, 'last-processed.json');

// Initialize JSON files
async function initializeJSONFiles() {
  try {
    // Create folder-structure.json if it doesn't exist
    try {
      await fs.access(FOLDER_STRUCTURE_PATH);
    } catch {
      await fs.writeFile(FOLDER_STRUCTURE_PATH, JSON.stringify({ folders: [] }, null, 2));
    }

    // Create last-processed.json if it doesn't exist
    try {
      await fs.access(LAST_PROCESSED_PATH);
    } catch {
      await fs.writeFile(LAST_PROCESSED_PATH, JSON.stringify({ lastFolder: null, lastFiles: [] }, null, 2));
    }
  } catch (error) {
    console.error('Error initializing JSON files:', error);
  }
}

// Read folder structure
async function readFolderStructure() {
  try {
    const data = await fs.readFile(FOLDER_STRUCTURE_PATH, 'utf-8');
    return JSON.parse(data);
  } catch (error) {
    console.error('Error reading folder structure:', error);
    return { folders: [] };
  }
}

// Write folder structure
async function writeFolderStructure(structure) {
  try {
    await fs.writeFile(FOLDER_STRUCTURE_PATH, JSON.stringify(structure, null, 2));
  } catch (error) {
    console.error('Error writing folder structure:', error);
  }
}

// Read last processed
async function readLastProcessed() {
  try {
    const data = await fs.readFile(LAST_PROCESSED_PATH, 'utf-8');
    return JSON.parse(data);
  } catch (error) {
    console.error('Error reading last processed:', error);
    return { lastFolder: null, lastFiles: [] };
  }
}

// Write last processed
async function writeLastProcessed(data) {
  try {
    await fs.writeFile(LAST_PROCESSED_PATH, JSON.stringify(data, null, 2));
  } catch (error) {
    console.error('Error writing last processed:', error);
  }
}

// Add folder to structure
async function addFolderToStructure(folderPath, folderName, files) {
  const structure = await readFolderStructure();

  const folderEntry = {
    path: folderPath,
    name: folderName,
    created: new Date().toISOString(),
    files: files.map(f => ({
      name: f,
      path: path.join(folderPath, f),
      created: new Date().toISOString()
    }))
  };

  // Check if folder already exists, update it
  const existingIndex = structure.folders.findIndex(f => f.path === folderPath);
  if (existingIndex >= 0) {
    structure.folders[existingIndex] = folderEntry;
  } else {
    structure.folders.unshift(folderEntry);
  }

  await writeFolderStructure(structure);

  // Update last processed
  await writeLastProcessed({
    lastFolder: folderPath,
    lastFolderName: folderName,
    lastFiles: files.map(f => path.join(folderPath, f)),
    timestamp: new Date().toISOString()
  });
}

let mainWindow;
let viewerWindow;
let settingsWindow;

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 1000,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
    resizable: true,
    center: true,
    icon: path.join(__dirname, 'assets/icon.png')
  });

  mainWindow.loadFile('index.html');

  // Open DevTools in development
  // mainWindow.webContents.openDevTools();

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

function createSettingsWindow() {
  if (settingsWindow) {
    settingsWindow.focus();
    return;
  }

  settingsWindow = new BrowserWindow({
    width: 700,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
    parent: mainWindow,
    modal: true,
    resizable: false
  });

  settingsWindow.loadFile('settings.html');

  settingsWindow.on('closed', () => {
    settingsWindow = null;
  });
}

function createViewerWindow() {
  if (viewerWindow) {
    viewerWindow.focus();
    return;
  }

  viewerWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
    parent: mainWindow
  });

  viewerWindow.loadFile('viewer.html');

  viewerWindow.on('closed', () => {
    viewerWindow = null;
  });
}

app.whenReady().then(async () => {
  await initializeJSONFiles();
  createMainWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createMainWindow();
  }
});

// IPC Handlers
ipcMain.handle('get-settings', () => {
  return {
    rootFolder: store.get('rootFolder', path.join(app.getPath('documents'), 'ImageDrop')),
    useOpenAI: store.get('useOpenAI', !!COMPANY_OPENAI_KEY), // Auto-enable if company key exists
    openAIKey: COMPANY_OPENAI_KEY || store.get('openAIKey', ''), // Use company key first
    hasCompanyKey: !!COMPANY_OPENAI_KEY // Let UI know if company key is set
  };
});

ipcMain.handle('save-settings', (event, settings) => {
  store.set('rootFolder', settings.rootFolder);
  // Only allow toggling OpenAI if no company key is set
  if (!COMPANY_OPENAI_KEY) {
    store.set('useOpenAI', settings.useOpenAI);
    store.set('openAIKey', settings.openAIKey);
  }
  return { success: true };
});

ipcMain.handle('select-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory', 'createDirectory']
  });

  if (!result.canceled && result.filePaths.length > 0) {
    return result.filePaths[0];
  }
  return null;
});

ipcMain.handle('open-settings', () => {
  createSettingsWindow();
});

ipcMain.handle('open-viewer', () => {
  createViewerWindow();
});

ipcMain.handle('capture-screen', async () => {
  try {
    const sources = await desktopCapturer.getSources({
      types: ['screen'],
      thumbnailSize: { width: 1920, height: 1080 }
    });

    if (sources.length > 0) {
      // Get the first screen (primary display)
      const source = sources[0];
      return source.thumbnail.toPNG().toString('base64');
    }
    throw new Error('No screen source available');
  } catch (error) {
    console.error('Screenshot error:', error);
    throw error;
  }
});

ipcMain.handle('open-snipping-tool', async (event, rootFolder) => {
  try {
    // Launch Windows Snipping Tool and monitor clipboard
    if (process.platform === 'win32') {
      exec('snippingtool /clip');
      return { success: true };
    }
    throw new Error('Snipping tool only available on Windows');
  } catch (error) {
    console.error('Snipping tool error:', error);
    throw error;
  }
});

ipcMain.handle('get-clipboard-image', () => {
  try {
    const image = clipboard.readImage();
    if (!image.isEmpty()) {
      return image.toPNG().toString('base64');
    }
    return null;
  } catch (error) {
    console.error('Clipboard error:', error);
    return null;
  }
});

ipcMain.handle('get-folders', async (event, rootFolder) => {
  try {
    await fs.mkdir(rootFolder, { recursive: true });
    const entries = await fs.readdir(rootFolder, { withFileTypes: true });
    const folders = [];

    for (const entry of entries) {
      if (entry.isDirectory()) {
        const folderPath = path.join(rootFolder, entry.name);
        const files = await fs.readdir(folderPath);
        const imageFiles = files.filter(f =>
          f.endsWith('.png') || f.endsWith('.jpg') || f.endsWith('.jpeg')
        );

        const stats = await fs.stat(folderPath);

        // Get first image for thumbnail
        let firstImage = null;
        if (imageFiles.length > 0) {
          firstImage = path.join(folderPath, imageFiles[0]);
        }

        folders.push({
          name: entry.name,
          path: folderPath,
          imageCount: imageFiles.length,
          created: stats.birthtime,
          firstImage: firstImage
        });
      }
    }

    // Sort by creation date, newest first
    folders.sort((a, b) => b.created - a.created);

    return folders;
  } catch (error) {
    console.error('Get folders error:', error);
    return [];
  }
});

ipcMain.handle('get-folder-images', async (event, folderPath) => {
  try {
    const files = await fs.readdir(folderPath);
    const imageFiles = files.filter(f =>
      f.endsWith('.png') || f.endsWith('.jpg') || f.endsWith('.jpeg')
    );

    const images = [];
    for (const file of imageFiles) {
      const filePath = path.join(folderPath, file);
      const stats = await fs.stat(filePath);
      images.push({
        name: file,
        path: filePath,
        created: stats.birthtime
      });
    }

    // Sort by creation date
    images.sort((a, b) => b.created - a.created);

    // Check for description file
    let description = null;
    try {
      const descPath = path.join(folderPath, 'description.txt');
      description = await fs.readFile(descPath, 'utf-8');
    } catch (e) {
      // No description file
    }

    return { images, description };
  } catch (error) {
    console.error('Get folder images error:', error);
    return { images: [], description: null };
  }
});

ipcMain.handle('delete-folder', async (event, folderPath) => {
  try {
    await fs.rm(folderPath, { recursive: true, force: true });
    return { success: true };
  } catch (error) {
    console.error('Delete folder error:', error);
    throw error;
  }
});

ipcMain.handle('open-image-external', async (event, imagePath) => {
  try {
    await shell.openPath(imagePath);
    return { success: true };
  } catch (error) {
    console.error('Open image error:', error);
    throw error;
  }
});

ipcMain.handle('update-folder-structure', async (event, folderPath, folderName, files) => {
  try {
    await addFolderToStructure(folderPath, folderName, files);
    return { success: true };
  } catch (error) {
    console.error('Update folder structure error:', error);
    throw error;
  }
});

ipcMain.handle('get-last-processed', async () => {
  try {
    return await readLastProcessed();
  } catch (error) {
    console.error('Get last processed error:', error);
    return { lastFolder: null, lastFiles: [] };
  }
});

ipcMain.handle('get-folder-structure', async () => {
  try {
    return await readFolderStructure();
  } catch (error) {
    console.error('Get folder structure error:', error);
    return { folders: [] };
  }
});

ipcMain.handle('save-etcim-json', async (event, data) => {
  try {
    const rootFolder = store.get('rootFolder', path.join(app.getPath('documents'), 'ImageDrop'));
    const etcimPath = path.join(rootFolder, 'etcim.json');
    await fs.writeFile(etcimPath, JSON.stringify(data, null, 2));

    // Send message to ETCIM20 window
    try {
      await sendToEtcimWindow(etcimPath);
    } catch (error) {
      console.error('Error sending to ETCIM20:', error);
      // Continue even if sending fails
    }

    return etcimPath;
  } catch (error) {
    console.error('Save etcim.json error:', error);
    throw error;
  }
});

// Function to send Windows message to ETCIM20
async function sendToEtcimWindow(jsonPath) {
  if (process.platform !== 'win32') {
    throw new Error('SendMessage only works on Windows');
  }

  // Escape the path for PowerShell
  const escapedPath = jsonPath.replace(/\\/g, '\\\\');

  // Create PowerShell script with proper COPYDATASTRUCT
  const psScript = `
Add-Type @"
  using System;
  using System.Runtime.InteropServices;

  [StructLayout(LayoutKind.Sequential)]
  public struct COPYDATASTRUCT {
    public IntPtr dwData;
    public int cbData;
    public IntPtr lpData;
  }

  public class WinAPI {
    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, ref COPYDATASTRUCT lParam);
  }
"@

try {
  $$hwnd = [WinAPI]::FindWindow($$null, "ETCIM20")
  if ($$hwnd -eq [IntPtr]::Zero) {
    Write-Error "ETCIM20 window not found"
    exit 1
  }

  $$path = "${escapedPath}"
  $$bytes = [System.Text.Encoding]::Unicode.GetBytes($$path)
  $$ptr = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($$bytes.Length)
  [System.Runtime.InteropServices.Marshal]::Copy($$bytes, 0, $$ptr, $$bytes.Length)

  $$cds = New-Object COPYDATASTRUCT
  $$cds.dwData = [IntPtr]::Zero
  $$cds.cbData = $$bytes.Length
  $$cds.lpData = $$ptr

  $$WM_COPYDATA = 0x004A
  $$result = [WinAPI]::SendMessage($$hwnd, $$WM_COPYDATA, [IntPtr]::Zero, [ref]$$cds)

  [System.Runtime.InteropServices.Marshal]::FreeHGlobal($$ptr)

  Write-Output "Message sent to ETCIM20 (result: $$result)"
} catch {
  Write-Error $$_.Exception.Message
  exit 1
}
`;

  return new Promise((resolve, reject) => {
    exec(`powershell -Command "${psScript.replace(/\$/g, '$')}"`, (error, stdout, stderr) => {
      if (error) {
        reject(new Error(`Failed to send to ETCIM20: ${stderr || error.message}`));
      } else {
        resolve(stdout.trim());
      }
    });
  });
}
