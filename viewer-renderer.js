const { ipcRenderer } = require('electron');
const path = require('path');

let currentSettings = {};
let currentFolderPath = null;

// Load settings and folders on startup
loadViewer();

async function loadViewer() {
    currentSettings = await ipcRenderer.invoke('get-settings');
    await loadFolders();
}

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
            <div class="folder-send-icon" title="Send to Etcim">📤</div>
            <div class="folder-delete-icon" title="Delete folder">×</div>
            ${thumbnailHtml}
            <div class="folder-card-content">
                <h3>${escapeHtml(folder.name)}</h3>
                <div class="folder-info">${folder.imageCount} image(s)</div>
                <div class="folder-date">${formatDate(folder.created)}</div>
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
    const { images, description } = await ipcRenderer.invoke('get-folder-images', folder.path);

    // Switch views
    document.getElementById('foldersView').style.display = 'none';
    document.getElementById('imagesView').classList.add('active');

    // Update header
    document.getElementById('folderName').textContent = folder.name;

    // Show description if exists
    const descBox = document.getElementById('descriptionBox');
    if (description) {
        descBox.textContent = description;
        descBox.style.display = 'block';
    } else {
        descBox.style.display = 'none';
    }

    // Display images
    displayImages(images);
}

function displayImages(images) {
    const grid = document.getElementById('imagesGrid');
    grid.innerHTML = '';

    if (images.length === 0) {
        grid.innerHTML = '<p style="color: #666; text-align: center; grid-column: 1 / -1;">No images in this folder.</p>';
        return;
    }

    images.forEach(image => {
        const card = document.createElement('div');
        card.className = 'image-card';
        card.innerHTML = `
            <img src="file:///${image.path.replace(/\\/g, '/')}" alt="${image.name}">
            <div class="image-info">
                <div class="image-name" title="${image.name}">${escapeHtml(image.name)}</div>
            </div>
        `;
        card.addEventListener('click', () => openImageModal(image.path));
        grid.appendChild(card);
    });
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
        const etcimPath = await ipcRenderer.invoke('save-etcim-json', etcimData);
        alert(`Folder "${folder.name}" sent to Etcim:\n${etcimPath}`);
    } catch (error) {
        alert('Error sending to Etcim: ' + error.message);
    }
}

// Event listeners
document.getElementById('refreshFolders').addEventListener('click', loadFolders);
document.getElementById('backToFolders').addEventListener('click', backToFolders);
document.getElementById('deleteFolder').addEventListener('click', deleteCurrentFolder);
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
