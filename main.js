const { app, BrowserWindow, ipcMain, dialog, clipboard, nativeImage, shell, desktopCapturer } = require('electron');
const path = require('path');
const fs = require('fs').promises;
const Store = require('electron-store');
const { exec } = require('child_process');
const { promisify } = require('util');

const execAsync = promisify(exec);
const store = new Store();

let mainWindow;
let viewerWindow;

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 600,
    height: 700,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    },
    resizable: true,
    icon: path.join(__dirname, 'assets/icon.png')
  });

  mainWindow.loadFile('index.html');

  // Open DevTools in development
  // mainWindow.webContents.openDevTools();

  mainWindow.on('closed', () => {
    mainWindow = null;
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

app.whenReady().then(createMainWindow);

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
    useOpenAI: store.get('useOpenAI', false),
    openAIKey: store.get('openAIKey', '')
  };
});

ipcMain.handle('save-settings', (event, settings) => {
  store.set('rootFolder', settings.rootFolder);
  store.set('useOpenAI', settings.useOpenAI);
  store.set('openAIKey', settings.openAIKey);
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

        folders.push({
          name: entry.name,
          path: folderPath,
          imageCount: imageFiles.length,
          created: stats.birthtime
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
