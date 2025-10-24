const { ipcRenderer } = require('electron');
const path = require('path');
const fs = require('fs').promises;
const sharp = require('sharp');
const { Poppler } = require('node-poppler');
const mammoth = require('mammoth');
const ExcelJS = require('exceljs');
const OpenAI = require('openai');

const poppler = new Poppler();

let currentSettings = {};

// Load settings on startup
loadSettings();

async function loadSettings() {
    currentSettings = await ipcRenderer.invoke('get-settings');
}

// Settings button
document.getElementById('openSettings').addEventListener('click', () => {
    ipcRenderer.invoke('open-settings');
});

// Dropzone handlers
const dropzone = document.getElementById('dropzone');

dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
});

dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
});

dropzone.addEventListener('drop', async (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');

    const files = Array.from(e.dataTransfer.files);
    await processFiles(files);
});

dropzone.addEventListener('click', () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = true;
    input.accept = '.pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg';
    input.onchange = async (e) => {
        const files = Array.from(e.target.files);
        await processFiles(files);
    };
    input.click();
});

// Clipboard paste
document.addEventListener('paste', async (e) => {
    const items = e.clipboardData.items;
    for (const item of items) {
        if (item.type.indexOf('image') !== -1) {
            const blob = item.getAsFile();
            await processClipboardImage(blob);
        }
    }
});

// Global keyboard shortcut
document.addEventListener('keydown', async (e) => {
    if (e.ctrlKey && e.key === 'v') {
        const base64Image = await ipcRenderer.invoke('get-clipboard-image');
        if (base64Image) {
            const buffer = Buffer.from(base64Image, 'base64');
            await saveImage(buffer, 'clipboard');
        }
    }
});

// Action buttons
document.getElementById('captureScreen').addEventListener('click', async () => {
    try {
        showStatus('Capturing screen...', 'info');
        const base64Image = await ipcRenderer.invoke('capture-screen');
        const buffer = Buffer.from(base64Image, 'base64');
        await saveImage(buffer, 'screenshot');
        showStatus('Screen captured successfully!', 'success');
    } catch (error) {
        showStatus('Error capturing screen: ' + error.message, 'error');
    }
});

document.getElementById('snippingTool').addEventListener('click', async () => {
    try {
        showStatus('Starten snippet tool... even geduld', 'info');

        // Capture current clipboard state BEFORE opening snipping tool
        const initialClipboard = await ipcRenderer.invoke('get-clipboard-image');

        await ipcRenderer.invoke('open-snipping-tool');

        showStatus('Selecteer een gebied om te knippen...', 'info');

        let timeoutHandle = null;
        let captureCompleted = false;

        // Start monitoring clipboard for NEW captured image
        const checkInterval = setInterval(async () => {
            const base64Image = await ipcRenderer.invoke('get-clipboard-image');
            // Only process if there's a new image different from initial state
            if (base64Image && base64Image !== initialClipboard) {
                captureCompleted = true;
                clearInterval(checkInterval);
                if (timeoutHandle) clearTimeout(timeoutHandle);

                // Close snipping tool
                await ipcRenderer.invoke('close-snipping-tool');

                showStatus('Verwerken snippet...', 'info');
                const buffer = Buffer.from(base64Image, 'base64');
                await saveImage(buffer, 'snippet-tool');
                showStatus('Snippet succesvol verwerkt!', 'success');
            }
        }, 500);

        // Stop checking after 30 seconds (only show error if no capture happened)
        timeoutHandle = setTimeout(() => {
            clearInterval(checkInterval);
            if (!captureCompleted) {
                showStatus('Snippet tool timeout - geen capture gedetecteerd', 'error');
            }
        }, 30000);
    } catch (error) {
        showStatus('Error: ' + error.message, 'error');
    }
});

document.getElementById('openViewer').addEventListener('click', () => {
    ipcRenderer.invoke('open-viewer');
});

// Process files
async function processFiles(files) {
    showStatus(`Processing ${files.length} file(s)...`, 'info');

    try {
        for (const file of files) {
            await processFile(file);
        }
        showStatus('All files processed successfully!', 'success');
    } catch (error) {
        showStatus('Error processing files: ' + error.message, 'error');
    }
}

async function processFile(file) {
    const ext = path.extname(file.name).toLowerCase();

    switch (ext) {
        case '.png':
        case '.jpg':
        case '.jpeg':
            const buffer = await file.arrayBuffer();
            await saveImage(Buffer.from(buffer), path.basename(file.name, ext));
            break;
        case '.pdf':
            await processPDF(file.path);
            break;
        case '.doc':
        case '.docx':
            await processWord(file.path);
            break;
        case '.xls':
        case '.xlsx':
            await processExcel(file.path);
            break;
        default:
            throw new Error(`Unsupported file type: ${ext}`);
    }
}

async function processPDF(filePath) {
    const outputDir = path.join(require('os').tmpdir(), 'pdf-convert-' + Date.now());
    await fs.mkdir(outputDir, { recursive: true });

    try {
        const outputFile = path.join(outputDir, 'page');

        // Convert PDF to PNG using node-poppler
        await poppler.pdfToCairo(filePath, outputFile, {
            pngFile: true,
            singleFile: false
        });

        const files = await fs.readdir(outputDir);
        for (const file of files) {
            if (file.endsWith('.png')) {
                const buffer = await fs.readFile(path.join(outputDir, file));
                await saveImage(buffer, 'pdf-page');
            }
        }
    } catch (error) {
        console.error('PDF conversion error:', error);
        showStatus('PDF conversion requires Poppler to be installed. See README.', 'error');
    } finally {
        await fs.rm(outputDir, { recursive: true });
    }
}

async function processWord(filePath) {
    const result = await mammoth.convertToHtml({ path: filePath });
    const html = result.value;

    // Create a simple HTML to image conversion
    // For now, save the HTML and use screenshot
    // This is simplified - in production, you might want to use puppeteer or similar
    const buffer = await createImageFromText(html);
    await saveImage(buffer, 'word-doc');
}

async function processExcel(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    let sheetIndex = 0;
    for (const worksheet of workbook.worksheets) {
        const data = [];
        worksheet.eachRow((row, rowNumber) => {
            const rowData = [];
            row.eachCell((cell, colNumber) => {
                rowData.push(cell.value?.toString() || '');
            });
            data.push(rowData.join('\t'));
        });

        const text = data.join('\n');
        const buffer = await createImageFromText(text);
        await saveImage(buffer, `excel-sheet-${sheetIndex++}`);
    }
}

async function createImageFromText(text) {
    // Create a simple text image using sharp
    const width = 800;
    const lines = text.substring(0, 5000).split('\n').slice(0, 50); // Limit text
    const height = Math.max(600, lines.length * 20);

    const svg = `
        <svg width="${width}" height="${height}">
            <rect width="${width}" height="${height}" fill="white"/>
            <text x="20" y="30" font-family="Arial" font-size="14" fill="black">
                ${lines.map((line, i) => `<tspan x="20" dy="20">${escapeXml(line.substring(0, 100))}</tspan>`).join('')}
            </text>
        </svg>
    `;

    return await sharp(Buffer.from(svg)).png().toBuffer();
}

function escapeXml(text) {
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

async function processClipboardImage(blob) {
    const buffer = Buffer.from(await blob.arrayBuffer());
    await saveImage(buffer, 'clipboard');
    showStatus('Image pasted successfully!', 'success');
}

async function saveImage(buffer, baseName) {
    // Convert to PNG if needed
    const pngBuffer = await sharp(buffer).png().toBuffer();

    // Create folder name
    if (currentSettings.useOpenAI && currentSettings.openAIKey) {
        showStatus('Generating AI folder name...', 'info');
    }
    const folderName = await getFolderName(pngBuffer);
    const folderPath = path.join(currentSettings.rootFolder, folderName);
    await fs.mkdir(folderPath, { recursive: true });

    // Create filename
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').replace('T', '_').split('Z')[0];
    const fileName = `${baseName}_${timestamp}.png`;
    const filePath = path.join(folderPath, fileName);

    // Save image
    await fs.writeFile(filePath, pngBuffer);

    // Save description if using OpenAI
    if (currentSettings.useOpenAI && currentSettings.openAIKey) {
        const descPath = path.join(folderPath, 'description.txt');
        try {
            const existingDesc = await fs.readFile(descPath, 'utf-8').catch(() => null);
            if (!existingDesc) {
                showStatus('Generating AI description...', 'info');
                const description = await generateDescription(pngBuffer);
                await fs.writeFile(descPath, description);
            }
        } catch (error) {
            console.error('Error saving description:', error);
        }
    }

    // Update JSON tracking
    try {
        const files = await fs.readdir(folderPath);
        const imageFiles = files.filter(f => f.endsWith('.png') || f.endsWith('.jpg') || f.endsWith('.jpeg'));
        await ipcRenderer.invoke('update-folder-structure', folderPath, folderName, imageFiles);
    } catch (error) {
        console.error('Error updating folder structure:', error);
    }
}

async function getFolderName(imageBuffer) {
    if (currentSettings.useOpenAI && currentSettings.openAIKey) {
        try {
            const title = await generateTitle(imageBuffer);
            // Sanitize folder name
            return title.replace(/[<>:"/\\|?*]/g, '_').substring(0, 100);
        } catch (error) {
            console.error('Error generating title:', error);
        }
    }

    // Default to timestamp
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}-${String(now.getSeconds()).padStart(2, '0')}`;
}

async function generateTitle(imageBuffer) {
    const openai = new OpenAI({
        apiKey: currentSettings.openAIKey,
        dangerouslyAllowBrowser: true
    });

    const base64Image = imageBuffer.toString('base64');

    const response = await openai.chat.completions.create({
        model: "gpt-4o-mini",
        messages: [
            {
                role: "user",
                content: [
                    {
                        type: "text",
                        text: "Provide a short, descriptive title for this image (5-10 words max). Only return the title, nothing else."
                    },
                    {
                        type: "image_url",
                        image_url: {
                            url: `data:image/png;base64,${base64Image}`
                        }
                    }
                ]
            }
        ],
        max_tokens: 50
    });

    return response.choices[0].message.content.trim();
}

async function generateDescription(imageBuffer) {
    const openai = new OpenAI({
        apiKey: currentSettings.openAIKey,
        dangerouslyAllowBrowser: true
    });

    const base64Image = imageBuffer.toString('base64');

    const response = await openai.chat.completions.create({
        model: "gpt-4o-mini",
        messages: [
            {
                role: "user",
                content: [
                    {
                        type: "text",
                        text: "Provide a brief description of this image (2-3 sentences)."
                    },
                    {
                        type: "image_url",
                        image_url: {
                            url: `data:image/png;base64,${base64Image}`
                        }
                    }
                ]
            }
        ],
        max_tokens: 150
    });

    return response.choices[0].message.content.trim();
}

function showStatus(message, type) {
    const status = document.getElementById('status');
    status.textContent = message;
    status.className = `status ${type}`;

    if (type === 'success') {
        setTimeout(() => {
            status.className = 'status';
        }, 3000);
    }
}

