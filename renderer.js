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

// Check clipboard on mouse enter dropzone
dropzone.addEventListener('mouseenter', async () => {
    const base64Image = await ipcRenderer.invoke('get-clipboard-image');
    const pasteBtn = document.getElementById('pasteClipboardBtn');

    if (base64Image) {
        pasteBtn.style.display = 'inline-block';
    }
});

dropzone.addEventListener('mouseleave', () => {
    document.getElementById('pasteClipboardBtn').style.display = 'none';
});

dropzone.addEventListener('dragover', async (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
    e.dataTransfer.effectAllowed = 'copy';
    dropzone.classList.add('dragover');

    // Show "Drop It..." overlay
    document.getElementById('dropOverlay').style.display = 'flex';

    // Don't show paste button when dragging
    document.getElementById('pasteClipboardBtn').style.display = 'none';
});

dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
    document.getElementById('dropOverlay').style.display = 'none';
});

dropzone.addEventListener('drop', async (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    document.getElementById('dropOverlay').style.display = 'none';
    document.getElementById('pasteClipboardBtn').style.display = 'none';

    const files = Array.from(e.dataTransfer.files);
    await processFiles(files, true); // true = real drop
});

// Paste clipboard button handler
document.getElementById('pasteClipboardBtn').addEventListener('click', async (e) => {
    e.stopPropagation(); // Prevent dropzone click event

    const base64Image = await ipcRenderer.invoke('get-clipboard-image');
    if (base64Image) {
        // Clear clipboard immediately after reading (before AI processing)
        await ipcRenderer.invoke('clear-clipboard');

        // Hide button immediately
        document.getElementById('pasteClipboardBtn').style.display = 'none';

        showStatus('Processing image from clipboard...', 'info');
        const buffer = Buffer.from(base64Image, 'base64');
        await saveImage(buffer, 'clipboard');
        showStatus('Clipboard image saved successfully!', 'success');
    }
});

dropzone.addEventListener('click', () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = true;
    input.accept = '.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.png,.jpg,.jpeg,.svg';
    input.onchange = async (e) => {
        const files = Array.from(e.target.files);
        await processFiles(files, true); // true = real file selection
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
        // Clear clipboard immediately after receiving
        await ipcRenderer.invoke('clear-clipboard');
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

                // Clear clipboard immediately after receiving
                await ipcRenderer.invoke('clear-clipboard');

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
async function processFiles(files, isRealDrop = false) {
    if (files.length === 0) {
        showStatus('Geen bestanden gedetecteerd. Mogelijk niet-ondersteund bestandstype.', 'error');
        return;
    }

    showStatus(`Processing ${files.length} file(s)...`, 'info');

    try {
        let processedCount = 0;
        let skippedCount = 0;

        for (const file of files) {
            const result = await processFile(file, isRealDrop);
            if (result === false) {
                skippedCount++;
            } else {
                processedCount++;
            }
        }

        if (processedCount > 0) {
            showStatus(`${processedCount} file(s) processed successfully!${skippedCount > 0 ? ` (${skippedCount} skipped)` : ''}`, 'success');
        } else if (skippedCount > 0) {
            showStatus(`${skippedCount} file(s) skipped (unsupported file types)`, 'error');
        }
    } catch (error) {
        showStatus('Error processing files: ' + error.message, 'error');
    }
}

async function processFile(file, isRealDrop = false) {
    const ext = path.extname(file.name).toLowerCase();
    const baseFileName = path.basename(file.name, ext);

    console.log('processFile called:', {
        fileName: file.name,
        ext: ext,
        isRealDrop: isRealDrop
    });

    // Check for supported file types
    const supportedTypes = ['.png', '.jpg', '.jpeg', '.svg', '.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx'];

    if (!supportedTypes.includes(ext)) {
        showStatus(`Bestandstype "${ext}" wordt niet ondersteund. Ondersteunde types: PDF, Word, Excel, PowerPoint, Afbeeldingen (PNG, JPG, JPEG, SVG)`, 'error');
        return false; // Return false to indicate file was skipped
    }

    switch (ext) {
        case '.png':
        case '.jpg':
        case '.jpeg':
            const buffer = await file.arrayBuffer();
            if (isRealDrop) {
                await saveImageWithOriginalName(Buffer.from(buffer), baseFileName, ext);
            } else {
                await saveImage(Buffer.from(buffer), baseFileName);
            }
            break;
        case '.svg':
            const svgBuffer = await file.arrayBuffer();
            await processSVG(Buffer.from(svgBuffer), file.name, isRealDrop);
            break;
        case '.pdf':
            // Get file path - if not available (file input), save to temp first
            const pdfPath = file.path || await saveToTemp(file);
            await processPDF(pdfPath, isRealDrop, baseFileName);
            break;
        case '.doc':
        case '.docx':
            const docPath = file.path || await saveToTemp(file);
            await processWord(docPath, file.name, isRealDrop);
            break;
        case '.xls':
        case '.xlsx':
            const xlsPath = file.path || await saveToTemp(file);
            await processExcel(xlsPath, file.name, isRealDrop);
            break;
        case '.ppt':
        case '.pptx':
            const pptPath = file.path || await saveToTemp(file);
            await processPowerPoint(pptPath, file.name, isRealDrop);
            break;
    }
}

// Helper function to save file to temp directory
async function saveToTemp(file) {
    const tmpDir = require('os').tmpdir();
    const tmpPath = path.join(tmpDir, `upload-${Date.now()}-${file.name}`);
    const buffer = await file.arrayBuffer();
    await fs.writeFile(tmpPath, Buffer.from(buffer));
    return tmpPath;
}

async function processPDF(filePath, isRealDrop = false, baseFileName = null) {
    const outputDir = path.join(require('os').tmpdir(), 'pdf-convert-' + Date.now());
    await fs.mkdir(outputDir, { recursive: true });

    try {
        const outputFile = path.join(outputDir, 'page');

        // Convert PDF to PNG using node-poppler
        await poppler.pdfToCairo(filePath, outputFile, {
            pngFile: true,
            singleFile: false
        });

        // Read original PDF
        const pdfBuffer = await fs.readFile(filePath);

        // Create folder name for this PDF
        const pdfBaseName = baseFileName || path.basename(filePath, '.pdf');
        const folderName = await getFolderName(null, pdfBaseName);
        const folderPath = path.join(currentSettings.rootFolder, folderName);
        await fs.mkdir(folderPath, { recursive: true });

        // Save original PDF to folder
        const pdfFileName = isRealDrop ? `${pdfBaseName}.pdf` : `${pdfBaseName}_${getTimestamp()}.pdf`;
        await fs.writeFile(path.join(folderPath, pdfFileName), pdfBuffer);

        // Save all PNG pages to the same folder
        const files = await fs.readdir(outputDir);
        const pngFiles = files.filter(f => f.endsWith('.png')).sort();
        const totalPages = pngFiles.length;
        const savedPngFiles = [];

        for (let i = 0; i < pngFiles.length; i++) {
            const file = pngFiles[i];
            const buffer = await fs.readFile(path.join(outputDir, file));
            const pngBuffer = await sharp(buffer).png().toBuffer();

            const pageNum = i + 1;
            const pngFileName = isRealDrop
                ? `${pdfBaseName} ${pageNum}-${totalPages}.png`
                : `${pdfBaseName}_page_${pageNum}_${getTimestamp()}.png`;
            await fs.writeFile(path.join(folderPath, pngFileName), pngBuffer);
            savedPngFiles.push(pngFileName);
        }

        // Update folder structure tracking
        const allFiles = [pdfFileName, ...savedPngFiles];
        await ipcRenderer.invoke('update-folder-structure', folderPath, folderName, allFiles);

        showStatus(`PDF processed: ${pdfFileName} + ${savedPngFiles.length} pages`, 'success');

    } catch (error) {
        console.error('PDF conversion error:', error);
        showStatus('PDF conversion requires Poppler to be installed. See README.', 'error');
    } finally {
        await fs.rm(outputDir, { recursive: true });
    }
}

async function processSVG(svgBuffer, fileName, isRealDrop = false) {
    try {
        showStatus('SVG omzetten naar PNG...', 'info');

        const baseName = path.basename(fileName, '.svg');
        const folderName = await getFolderName(null, baseName);
        const folderPath = path.join(currentSettings.rootFolder, folderName);
        await fs.mkdir(folderPath, { recursive: true });

        // Save original SVG
        const savedSvgFileName = isRealDrop ? fileName : `${baseName}_${getTimestamp()}.svg`;
        await fs.writeFile(path.join(folderPath, savedSvgFileName), svgBuffer);

        // Convert SVG to PNG with max width 1080 and transparency
        const pngBuffer = await sharp(svgBuffer)
            .resize(1080, null, {
                fit: 'inside',
                withoutEnlargement: true
            })
            .png()
            .toBuffer();

        // Save PNG
        const pngFileName = isRealDrop ? `${baseName}.png` : `${baseName}_${getTimestamp()}.png`;
        await fs.writeFile(path.join(folderPath, pngFileName), pngBuffer);

        const allFiles = [savedSvgFileName, pngFileName];
        await ipcRenderer.invoke('update-folder-structure', folderPath, folderName, allFiles);
        showStatus(`SVG verwerkt: ${savedSvgFileName} + ${pngFileName}`, 'success');

    } catch (error) {
        console.error('SVG processing error:', error);
        showStatus('Error processing SVG: ' + error.message, 'error');
    }
}

async function processWord(filePath, fileName, isRealDrop = false) {
    try {
        // Save original Word file
        const fileBuffer = await fs.readFile(filePath);
        const baseName = path.basename(fileName, path.extname(fileName));
        const folderName = await getFolderName(null, baseName);
        const folderPath = path.join(currentSettings.rootFolder, folderName);
        await fs.mkdir(folderPath, { recursive: true });

        const savedFileName = isRealDrop ? fileName : `${baseName}_${getTimestamp()}${path.extname(fileName)}`;
        await fs.writeFile(path.join(folderPath, savedFileName), fileBuffer);

        // Try Word to PDF conversion
        showStatus('Word omzetten naar PDF...', 'info');

        try {
            const pdfPath = await convertWordToPDF(filePath);

            if (pdfPath) {
                // Check if PDF was created
                const pdfExists = await fs.access(pdfPath).then(() => true).catch(() => false);

                if (!pdfExists) {
                    throw new Error('PDF niet aangemaakt');
                }

                // Save PDF to folder
                const pdfFileName = isRealDrop ? `${baseName}.pdf` : `${baseName}_${getTimestamp()}.pdf`;
                const savedPdfPath = path.join(folderPath, pdfFileName);
                await fs.copyFile(pdfPath, savedPdfPath);

                showStatus('PDF omzetten naar PNG...', 'info');
                const pngFiles = await convertPDFToPNG(pdfPath, folderPath, baseName, isRealDrop);
                const allFiles = [savedFileName, pdfFileName, ...pngFiles];
                await ipcRenderer.invoke('update-folder-structure', folderPath, folderName, allFiles);
                showStatus(`Word verwerkt: ${savedFileName} + ${pngFiles.length} pagina's`, 'success');
            } else {
                throw new Error('Geen PDF output');
            }
        } catch (conversionError) {
            console.log('Word conversie gefaald:', conversionError.message);
            showStatus(`Word conversie gefaald: ${conversionError.message}`, 'error');
            await ipcRenderer.invoke('update-folder-structure', folderPath, folderName, [savedFileName]);
            showStatus(`Word document opgeslagen: ${savedFileName}`, 'success');
        }
    } catch (error) {
        console.error('Word processing error:', error);
        showStatus('Error processing Word document: ' + error.message, 'error');
    }
}

async function processExcel(filePath, fileName, isRealDrop = false) {
    try {
        // Save original Excel file
        const fileBuffer = await fs.readFile(filePath);
        const baseName = path.basename(fileName, path.extname(fileName));
        const folderName = await getFolderName(null, baseName);
        const folderPath = path.join(currentSettings.rootFolder, folderName);
        await fs.mkdir(folderPath, { recursive: true });

        const savedFileName = isRealDrop ? fileName : `${baseName}_${getTimestamp()}${path.extname(fileName)}`;
        await fs.writeFile(path.join(folderPath, savedFileName), fileBuffer);

        // Try Excel to PDF conversion
        showStatus('Excel omzetten naar PDF...', 'info');

        try {
            const pdfPath = await convertExcelToPDF(filePath);

            if (pdfPath) {
                // Check if PDF was created
                const pdfExists = await fs.access(pdfPath).then(() => true).catch(() => false);

                if (!pdfExists) {
                    throw new Error('PDF niet aangemaakt');
                }

                // Save PDF to folder
                const pdfFileName = isRealDrop ? `${baseName}.pdf` : `${baseName}_${getTimestamp()}.pdf`;
                const savedPdfPath = path.join(folderPath, pdfFileName);
                await fs.copyFile(pdfPath, savedPdfPath);

                showStatus('PDF omzetten naar PNG...', 'info');
                const pngFiles = await convertPDFToPNG(pdfPath, folderPath, baseName, isRealDrop);
                const allFiles = [savedFileName, pdfFileName, ...pngFiles];
                await ipcRenderer.invoke('update-folder-structure', folderPath, folderName, allFiles);
                showStatus(`Excel verwerkt: ${savedFileName} + ${pngFiles.length} sheets`, 'success');
            } else {
                throw new Error('Geen PDF output');
            }
        } catch (conversionError) {
            console.log('Excel conversie gefaald:', conversionError.message);
            showStatus(`Excel conversie gefaald: ${conversionError.message}`, 'error');
            await ipcRenderer.invoke('update-folder-structure', folderPath, folderName, [savedFileName]);
            showStatus(`Excel document opgeslagen: ${savedFileName}`, 'success');
        }
    } catch (error) {
        console.error('Excel processing error:', error);
        showStatus('Error processing Excel document: ' + error.message, 'error');
    }
}

async function processPowerPoint(filePath, fileName, isRealDrop = false) {
    try {
        // Save original PowerPoint file
        const fileBuffer = await fs.readFile(filePath);
        const baseName = path.basename(fileName, path.extname(fileName));
        const folderName = await getFolderName(null, baseName);
        const folderPath = path.join(currentSettings.rootFolder, folderName);
        await fs.mkdir(folderPath, { recursive: true });

        const savedFileName = isRealDrop ? fileName : `${baseName}_${getTimestamp()}${path.extname(fileName)}`;
        await fs.writeFile(path.join(folderPath, savedFileName), fileBuffer);

        // Try PowerPoint to PDF conversion
        showStatus('PowerPoint omzetten naar PDF...', 'info');

        try {
            const pdfPath = await convertPowerPointToPDF(filePath);

            if (pdfPath) {
                // Check if PDF was created
                const pdfExists = await fs.access(pdfPath).then(() => true).catch(() => false);

                if (!pdfExists) {
                    throw new Error('PDF niet aangemaakt');
                }

                // Save PDF to folder
                const pdfFileName = isRealDrop ? `${baseName}.pdf` : `${baseName}_${getTimestamp()}.pdf`;
                const savedPdfPath = path.join(folderPath, pdfFileName);
                await fs.copyFile(pdfPath, savedPdfPath);

                showStatus('PDF omzetten naar PNG...', 'info');
                const pngFiles = await convertPDFToPNG(pdfPath, folderPath, baseName, isRealDrop);
                const allFiles = [savedFileName, pdfFileName, ...pngFiles];
                await ipcRenderer.invoke('update-folder-structure', folderPath, folderName, allFiles);
                showStatus(`PowerPoint verwerkt: ${savedFileName} + ${pngFiles.length} slides`, 'success');
            } else {
                throw new Error('Geen PDF output');
            }
        } catch (conversionError) {
            console.log('PowerPoint conversie gefaald:', conversionError.message);
            showStatus(`PowerPoint conversie gefaald: ${conversionError.message}`, 'error');
            await ipcRenderer.invoke('update-folder-structure', folderPath, folderName, [savedFileName]);
            showStatus(`PowerPoint document opgeslagen: ${savedFileName}`, 'success');
        }
    } catch (error) {
        console.error('PowerPoint processing error:', error);
        showStatus('Error processing PowerPoint: ' + error.message, 'error');
    }
}

async function convertWordToPDF(filePath) {
    try {
        const outputDir = path.join(require('os').tmpdir(), `word-to-pdf-${Date.now()}`);
        await fs.mkdir(outputDir, { recursive: true });

        const pdfPath = path.join(outputDir, 'output.pdf');
        const scriptPath = path.join(__dirname, 'scripts', 'word-to-pdf.ps1');

        const { exec } = require('child_process');
        const { promisify } = require('util');
        const execAsync = promisify(exec);

        console.log('Word to PDF conversion:');
        console.log('Input:', filePath);
        console.log('Output:', pdfPath);

        const command = `powershell -ExecutionPolicy Bypass -File "${scriptPath}" -inputPath "${filePath}" -outputPath "${pdfPath}"`;
        console.log('Command:', command);

        const { stdout, stderr } = await execAsync(command, { timeout: 30000 });

        if (stderr) {
            console.error('PowerShell stderr:', stderr);
        }
        if (stdout) {
            console.log('PowerShell stdout:', stdout);
        }

        return pdfPath;
    } catch (error) {
        console.error('Word to PDF conversion error:', error);
        throw error;
    }
}

async function convertExcelToPDF(filePath) {
    try {
        const outputDir = path.join(require('os').tmpdir(), `excel-to-pdf-${Date.now()}`);
        await fs.mkdir(outputDir, { recursive: true });

        const pdfPath = path.join(outputDir, 'output.pdf');
        const scriptPath = path.join(__dirname, 'scripts', 'excel-to-pdf.ps1');

        const { exec } = require('child_process');
        const { promisify } = require('util');
        const execAsync = promisify(exec);

        console.log('Excel to PDF conversion:');
        console.log('Input:', filePath);
        console.log('Output:', pdfPath);

        const command = `powershell -ExecutionPolicy Bypass -File "${scriptPath}" -inputPath "${filePath}" -outputPath "${pdfPath}"`;
        console.log('Command:', command);

        const { stdout, stderr } = await execAsync(command, { timeout: 30000 });

        if (stderr) {
            console.error('PowerShell stderr:', stderr);
        }
        if (stdout) {
            console.log('PowerShell stdout:', stdout);
        }

        return pdfPath;
    } catch (error) {
        console.error('Excel to PDF conversion error:', error);
        throw error;
    }
}

async function convertPowerPointToPDF(filePath) {
    try {
        const outputDir = path.join(require('os').tmpdir(), `powerpoint-to-pdf-${Date.now()}`);
        await fs.mkdir(outputDir, { recursive: true });

        const pdfPath = path.join(outputDir, 'output.pdf');
        const scriptPath = path.join(__dirname, 'scripts', 'powerpoint-to-pdf.ps1');

        const { exec } = require('child_process');
        const { promisify } = require('util');
        const execAsync = promisify(exec);

        console.log('PowerPoint to PDF conversion:');
        console.log('Input:', filePath);
        console.log('Output:', pdfPath);

        const command = `powershell -ExecutionPolicy Bypass -File "${scriptPath}" -inputPath "${filePath}" -outputPath "${pdfPath}"`;
        console.log('Command:', command);

        const { stdout, stderr } = await execAsync(command, { timeout: 30000 });

        if (stderr) {
            console.error('PowerShell stderr:', stderr);
        }
        if (stdout) {
            console.log('PowerShell stdout:', stdout);
        }

        return pdfPath;
    } catch (error) {
        console.error('PowerPoint to PDF conversion error:', error);
        throw error;
    }
}

async function convertOfficeToPDF(filePath, type) {
    // Use PowerShell script to convert Office files to PDF or PNG
    // Returns PDF path for Word/Excel, PNG directory for PowerPoint
    try {
        const outputDir = path.join(require('os').tmpdir(), `office-convert-${Date.now()}`);
        await fs.mkdir(outputDir, { recursive: true });

        const scriptPath = path.join(__dirname, 'scripts', 'convert-office.ps1');
        const { exec } = require('child_process');
        const { promisify } = require('util');
        const execAsync = promisify(exec);

        console.log('PowerShell script path:', scriptPath);
        console.log('Input file:', filePath);
        console.log('Output dir:', outputDir);

        // Run PowerShell script
        const command = `powershell -ExecutionPolicy Bypass -File "${scriptPath}" -InputFile "${filePath}" -OutputFolder "${outputDir}"`;
        console.log('Executing command:', command);

        const { stdout, stderr } = await execAsync(command, { timeout: 60000 });

        console.log('PowerShell stdout:', stdout);
        console.log('PowerShell stderr:', stderr);

        if (stderr && !stdout) {
            console.error('PowerShell conversion error:', stderr);
            return null;
        }

        const output = stdout.trim();
        console.log('Output path:', output);

        if (output && await fs.access(output).then(() => true).catch(() => false)) {
            console.log('Output file exists:', output);
            return output;
        }

        console.log('Output file does not exist or no output');
        return null;
    } catch (error) {
        console.error('Office to PDF conversion error:', error);
        console.error('Error details:', error.message);
        if (error.stderr) console.error('stderr:', error.stderr);
        if (error.stdout) console.error('stdout:', error.stdout);
        return null;
    }
}

async function convertPDFToPNG(pdfPath, targetFolder, baseName, isRealDrop = false) {
    const outputDir = path.join(require('os').tmpdir(), 'pdf-to-png-' + Date.now());
    await fs.mkdir(outputDir, { recursive: true });

    try {
        const outputFile = path.join(outputDir, 'page');

        await poppler.pdfToCairo(pdfPath, outputFile, {
            pngFile: true,
            singleFile: false
        });

        const files = await fs.readdir(outputDir);
        const pngFiles = files.filter(f => f.endsWith('.png')).sort();
        const totalPages = pngFiles.length;
        const savedPngFiles = [];

        for (let i = 0; i < pngFiles.length; i++) {
            const file = pngFiles[i];
            const buffer = await fs.readFile(path.join(outputDir, file));
            const pngBuffer = await sharp(buffer).png().toBuffer();

            const pageNum = i + 1;
            const pngFileName = isRealDrop
                ? `${baseName} ${pageNum}-${totalPages}.png`
                : `${baseName}_page_${pageNum}_${getTimestamp()}.png`;
            await fs.writeFile(path.join(targetFolder, pngFileName), pngBuffer);
            savedPngFiles.push(pngFileName);
        }

        return savedPngFiles;
    } finally {
        await fs.rm(outputDir, { recursive: true });
        // Clean up source PDF if it was temporary
        try {
            await fs.rm(path.dirname(pdfPath), { recursive: true });
        } catch {
            // Ignore cleanup errors
        }
    }
}

async function copyPowerPointPNGs(pngDirPath, targetFolder, baseName, isRealDrop = false) {
    try {
        const files = await fs.readdir(pngDirPath);
        const pngFilenames = files.filter(f => f.toLowerCase().endsWith('.png')).sort();
        const totalSlides = pngFilenames.length;
        const savedPngFiles = [];

        for (let i = 0; i < pngFilenames.length; i++) {
            const file = pngFilenames[i];
            const buffer = await fs.readFile(path.join(pngDirPath, file));
            const pngBuffer = await sharp(buffer).png().toBuffer();

            const slideNum = i + 1;
            const pngFileName = isRealDrop
                ? `${baseName} ${slideNum}-${totalSlides}.png`
                : `${baseName}_slide_${slideNum}_${getTimestamp()}.png`;
            await fs.writeFile(path.join(targetFolder, pngFileName), pngBuffer);
            savedPngFiles.push(pngFileName);
        }

        // Clean up temp directory
        await fs.rm(pngDirPath, { recursive: true });

        return savedPngFiles;
    } catch (error) {
        console.error('Error copying PowerPoint PNGs:', error);
        return [];
    }
}


async function processClipboardImage(blob) {
    const buffer = Buffer.from(await blob.arrayBuffer());
    await saveImage(buffer, 'clipboard');
    showStatus('Image pasted successfully!', 'success');
}

function getTimestamp() {
    return new Date().toISOString().replace(/[:.]/g, '-').replace('T', '_').split('Z')[0];
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

    // Create filename with timestamp (for clipboard/screenshot operations)
    const fileName = `${baseName}_${getTimestamp()}.png`;
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

async function saveImageWithOriginalName(buffer, baseName, ext) {
    // Convert to PNG if needed
    const pngBuffer = await sharp(buffer).png().toBuffer();

    // Create folder name
    if (currentSettings.useOpenAI && currentSettings.openAIKey) {
        showStatus('Generating AI folder name...', 'info');
    }
    const folderName = await getFolderName(pngBuffer);
    const folderPath = path.join(currentSettings.rootFolder, folderName);
    await fs.mkdir(folderPath, { recursive: true });

    // Use original filename (for real drops)
    const fileName = `${baseName}.png`;
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

async function getFolderName(imageBuffer, fallbackName = null) {
    let baseName = '';

    if (imageBuffer && currentSettings.useOpenAI && currentSettings.openAIKey) {
        try {
            const title = await generateTitle(imageBuffer);
            // Sanitize folder name
            baseName = title.replace(/[<>:"/\\|?*]/g, '_').substring(0, 100);
        } catch (error) {
            console.error('Error generating title:', error);
        }
    }

    // Use fallback name if provided (e.g., PDF filename)
    if (!baseName && fallbackName) {
        baseName = fallbackName.replace(/[<>:"/\\|?*]/g, '_').substring(0, 100);
    }

    // Default to timestamp
    if (!baseName) {
        const now = new Date();
        baseName = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}-${String(now.getSeconds()).padStart(2, '0')}`;
    }

    // Check if folder exists and add index if needed
    let folderName = baseName;
    let index = 2;
    while (await fs.access(path.join(currentSettings.rootFolder, folderName)).then(() => true).catch(() => false)) {
        folderName = `${baseName} ${index}`;
        index++;
    }

    return folderName;
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

    // Auto-hide after 5 seconds for all message types
    setTimeout(() => {
        status.className = 'status';
        status.textContent = '';
    }, 5000);
}

