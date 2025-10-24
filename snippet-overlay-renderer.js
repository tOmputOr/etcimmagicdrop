const { ipcRenderer } = require('electron');

const canvas = document.getElementById('canvas');
const ctx = canvas.getContext('2d');

let startX = 0;
let startY = 0;
let isDrawing = false;
let displayInfo = null;
let overlayReady = false;

// Get display info and setup canvas
ipcRenderer.invoke('get-display-info').then(info => {
    displayInfo = info;

    // Set canvas size to full virtual screen (all monitors)
    canvas.width = info.totalWidth;
    canvas.height = info.totalHeight;

    console.log('Display info:', info);

    // Setup event listeners AFTER display info is ready
    setupEventListeners();

    // Mark overlay as ready after a short delay to prevent accidental clicks
    setTimeout(() => {
        overlayReady = true;
        console.log('Overlay ready for capture');
    }, 200);
});

function setupEventListeners() {
    // Mouse down - start selection
    canvas.addEventListener('mousedown', (e) => {
        if (!overlayReady) {
            console.log('Overlay not ready yet, ignoring click');
            return;
        }
        isDrawing = true;
        startX = e.screenX - displayInfo.bounds.x;
        startY = e.screenY - displayInfo.bounds.y;
        console.log('Start selection:', startX, startY);
    });

    // Mouse move - draw selection rectangle
    canvas.addEventListener('mousemove', (e) => {
        if (!isDrawing) return;

        // Clear canvas
        ctx.clearRect(0, 0, canvas.width, canvas.height);

        // Draw semi-transparent overlay
        ctx.fillStyle = 'rgba(0, 0, 0, 0.3)';
        ctx.fillRect(0, 0, canvas.width, canvas.height);

        // Calculate current position
        const currentX = e.screenX - displayInfo.bounds.x;
        const currentY = e.screenY - displayInfo.bounds.y;
        const width = currentX - startX;
        const height = currentY - startY;

        // Clear selected area (make it transparent)
        ctx.clearRect(startX, startY, width, height);

        // Draw selection border
        ctx.strokeStyle = '#DC143C';
        ctx.lineWidth = 2;
        ctx.strokeRect(startX, startY, width, height);

        // Draw dimensions label
        const labelX = e.clientX;
        const labelY = e.clientY;
        ctx.fillStyle = 'rgba(0, 0, 0, 0.8)';
        ctx.fillRect(labelX + 10, labelY + 10, 100, 30);
        ctx.fillStyle = 'white';
        ctx.font = '12px Arial';
        ctx.fillText(`${Math.abs(width)}×${Math.abs(height)}`, labelX + 15, labelY + 30);
    });

    // Mouse up - capture selection
    canvas.addEventListener('mouseup', async (e) => {
        if (!isDrawing) return;
        isDrawing = false;

        const endX = e.screenX - displayInfo.bounds.x;
        const endY = e.screenY - displayInfo.bounds.y;

        // Calculate bounds (normalize to handle dragging in any direction)
        const bounds = {
            x: Math.min(startX, endX),
            y: Math.min(startY, endY),
            width: Math.abs(endX - startX),
            height: Math.abs(endY - startY)
        };

        console.log('End selection:', endX, endY, 'Bounds:', bounds);

        // Only capture if selection has size
        if (bounds.width > 5 && bounds.height > 5) {
            console.log('Capturing region...');
            await ipcRenderer.invoke('capture-region', bounds);
        } else {
            console.log('Selection too small, closing overlay');
            // Cancel if too small
            ipcRenderer.invoke('close-snippet-overlay');
        }
    });

    // ESC key to cancel
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            ipcRenderer.invoke('close-snippet-overlay');
        }
    });
}
