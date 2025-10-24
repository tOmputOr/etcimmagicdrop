const { ipcRenderer } = require('electron');
const path = require('path');

let currentSettings = {};
let currentFolderPath = null;
let isListView = false;

// Load settings and folders on startup
loadViewer();

async function loadViewer() {
    currentSettings = await ipcRenderer.invoke('get-settings');
    await loadFolders();
}

// Toggle view button
document.getElementById('toggleView').addEventListener('click', () => {
    isListView = !isListView;
    const grid = document.getElementById('foldersGrid');
    const btn = document.getElementById('toggleView');
    const icon = document.getElementById('viewIcon');

    if (isListView) {
        grid.classList.add('list-view');
        icon.textContent = '▦';
        btn.innerHTML = '<span id="viewIcon">▦</span> Grid View';
    } else {
        grid.classList.remove('list-view');
        icon.textContent = '☰';
        btn.innerHTML = '<span id="viewIcon">☰</span> List View';
    }
});

async function loadFolders() {
    const folders = await ipcRenderer.invoke('get-folders', currentSettings.rootFolder);
    displayFolders(folders);
}

function displayFolders(folders) {
    const grid = document.getElementById('foldersGrid');
    grid.innerHTML = '';

    if (folders.length === 0) {
        grid.innerHTML = '<p style="color: #666; text-align: center; grid-column: 1 / -1;">No folders yet. Start dropping images!</p>';
        return;
    }

    folders.forEach(folder => {
        const card = document.createElement('div');
        card.className = 'folder-card';

        // Create thumbnail or placeholder
        let thumbnailHtml = '';
        if (folder.firstImage) {
            thumbnailHtml = `<div class="folder-thumbnail"><img src="file:///${folder.firstImage.replace(/\\/g, '/')}" alt="Preview"></div>`;
        } else {
            thumbnailHtml = `<div class="folder-thumbnail">📁</div>`;
        }

        card.innerHTML = `
            <div class="folder-send-icon" title="Send to Etcim">
                <img src="assets/icons/toEtcim.svg" alt="Send to Etcim" />
            </div>
            <div class="folder-delete-icon" title="Delete folder">
                <img src="assets/icons/delete_icon.svg" alt="Delete folder" />
            </div>
            ${thumbnailHtml}
            <div class="folder-card-content">
                <div class="folder-title-row">
                    <h3>${escapeHtml(folder.name)}</h3>
                    <span class="folder-meta">${folder.imageCount} image(s) • ${formatDate(folder.created)}</span>
                </div>
            </div>
        `;

        // Add click handler for opening folder
        card.addEventListener('click', (e) => {
            if (!e.target.classList.contains('folder-delete-icon') &&
                !e.target.classList.contains('folder-send-icon')) {
                openFolder(folder);
            }
        });

        // Add send to etcim handler
        const sendIcon = card.querySelector('.folder-send-icon');
        sendIcon.addEventListener('click', async (e) => {
            e.stopPropagation();
            await sendFolderToEtcim(folder);
        });

        // Add delete handler
        const deleteIcon = card.querySelector('.folder-delete-icon');
        deleteIcon.addEventListener('click', async (e) => {
            e.stopPropagation();
            if (confirm(`Delete folder "${folder.name}" and all its contents?`)) {
                try {
                    await ipcRenderer.invoke('delete-folder', folder.path);
                    loadFolders();
                } catch (error) {
                    alert('Error deleting folder: ' + error.message);
                }
            }
        });

        grid.appendChild(card);
    });
}

async function openFolder(folder) {
    currentFolderPath = folder.path;
    const fs = require('fs').promises;

    // Get all files in folder, not just images
    const files = await fs.readdir(folder.path);
    const fileStats = [];

    for (const file of files) {
        if (file === 'description.txt') continue; // Skip description file
        const filePath = path.join(folder.path, file);
        const stats = await fs.stat(filePath);
        fileStats.push({
            name: file,
            path: filePath,
            created: stats.birthtime,
            ext: path.extname(file).toLowerCase()
        });
    }

    // Sort by creation date, oldest first (so original comes before pages)
    fileStats.sort((a, b) => a.created - b.created);

    // Switch views
    document.getElementById('foldersView').style.display = 'none';
    document.getElementById('imagesView').classList.add('active');

    // Update header
    document.getElementById('folderName').textContent = folder.name;

    // Show description if exists
    const descBox = document.getElementById('descriptionBox');
    try {
        const descPath = path.join(folder.path, 'description.txt');
        const description = await fs.readFile(descPath, 'utf-8');
        descBox.textContent = description;
        descBox.style.display = 'block';
    } catch {
        descBox.style.display = 'none';
    }

    // Display files
    displayFiles(fileStats);
}

function getFileIcon(ext) {
    const icons = {
        '.png': '🖼️',
        '.jpg': '🖼️',
        '.jpeg': '🖼️',
        '.pdf': '📄',
        '.doc': '📝',
        '.docx': '📝',
        '.xls': '📊',
        '.xlsx': '📊',
        '.ppt': '📊',
        '.pptx': '📊'
    };
    return icons[ext] || '📎';
}

function isImageFile(ext) {
    return ['.png', '.jpg', '.jpeg'].includes(ext);
}

function displayFiles(files) {
    const grid = document.getElementById('imagesGrid');
    grid.innerHTML = '';

    if (files.length === 0) {
        grid.innerHTML = '<p style="color: #666; text-align: center; grid-column: 1 / -1;">No files in this folder.</p>';
        return;
    }

    files.forEach(file => {
        const card = document.createElement('div');
        card.className = 'image-card';

        if (isImageFile(file.ext)) {
            // Image files show thumbnail
            card.innerHTML = `
                <img src="file:///${file.path.replace(/\\/g, '/')}" alt="${file.name}">
                <div class="image-info">
                    <div class="image-name" title="${file.name}">${escapeHtml(file.name)}</div>
                </div>
            `;
            card.addEventListener('click', () => openImageModal(file.path));
        } else {
            // Non-image files show icon
            card.innerHTML = `
                <div class="file-icon-preview">
                    <div class="file-type-icon">${getFileIcon(file.ext)}</div>
                    <div class="file-type-label">${file.ext.substring(1).toUpperCase()}</div>
                </div>
                <div class="image-info">
                    <div class="image-name" title="${file.name}">${escapeHtml(file.name)}</div>
                </div>
            `;
            card.addEventListener('click', () => openFile(file.path));
        }

        grid.appendChild(card);
    });
}

async function openFile(filePath) {
    try {
        await ipcRenderer.invoke('open-image-external', filePath);
    } catch (error) {
        alert('Error opening file: ' + error.message);
    }
}

function openImageModal(imagePath) {
    const modal = document.getElementById('imageModal');
    const modalImage = document.getElementById('modalImage');
    modalImage.src = `file:///${imagePath.replace(/\\/g, '/')}`;
    modal.classList.add('active');
}

function closeImageModal() {
    const modal = document.getElementById('imageModal');
    modal.classList.remove('active');
}

function backToFolders() {
    document.getElementById('imagesView').classList.remove('active');
    document.getElementById('foldersView').style.display = 'block';
    currentFolderPath = null;
    loadFolders();
}

async function deleteCurrentFolder() {
    if (!currentFolderPath) return;

    const confirmation = confirm('Are you sure you want to delete this folder and all its images?');
    if (!confirmation) return;

    try {
        await ipcRenderer.invoke('delete-folder', currentFolderPath);
        backToFolders();
    } catch (error) {
        alert('Error deleting folder: ' + error.message);
    }
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function formatDate(date) {
    const d = new Date(date);
    return d.toLocaleDateString() + ' ' + d.toLocaleTimeString();
}

async function sendFolderToEtcim(folder) {
    try {
        // Get the full folder data including files
        const { images } = await ipcRenderer.invoke('get-folder-images', folder.path);

        const etcimData = {
            timestamp: new Date().toISOString(),
            folder: {
                name: folder.name,
                path: folder.path,
                created: folder.created,
                fileCount: folder.imageCount,
                files: images.map(img => ({
                    name: img.name,
                    path: img.path,
                    created: img.created
                }))
            }
        };

        // Save to etcim.json in root folder (overwrites)
        const result = await ipcRenderer.invoke('save-etcim-json', etcimData);
        console.log('Send to ETCIM20 result:', result);

        // Show brief success message
        showToast('Doorgestuurd naar ETCIM');
    } catch (error) {
        console.error('Error sending to Etcim:', error);
        showToast('Fout bij doorsturen naar ETCIM', 'error');
    }
}

function showToast(message, type = 'success') {
    // Create toast element
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    document.body.appendChild(toast);

    // Show toast
    setTimeout(() => toast.classList.add('show'), 10);

    // Hide and remove after 3 seconds
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

async function openCurrentFolder() {
    if (!currentFolderPath) return;

    try {
        await ipcRenderer.invoke('open-folder', currentFolderPath);
    } catch (error) {
        alert('Error opening folder: ' + error.message);
    }
}

// Event listeners
document.getElementById('refreshFolders').addEventListener('click', loadFolders);
document.getElementById('backToFolders').addEventListener('click', backToFolders);
document.getElementById('openFolder').addEventListener('click', openCurrentFolder);
document.getElementById('processFolder').addEventListener('click', () => {
    // TODO: Add process functionality
    console.log('Process folder clicked');
});
document.getElementById('closeModal').addEventListener('click', closeImageModal);
document.getElementById('imageModal').addEventListener('click', (e) => {
    if (e.target.id === 'imageModal') {
        closeImageModal();
    }
});

// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
        const modal = document.getElementById('imageModal');
        if (modal.classList.contains('active')) {
            closeImageModal();
        }
    }
});
