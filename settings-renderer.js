const { ipcRenderer } = require('electron');

let currentSettings = {};

loadSettings();

async function loadSettings() {
    currentSettings = await ipcRenderer.invoke('get-settings');
    document.getElementById('rootFolder').value = currentSettings.rootFolder;
    document.getElementById('useOpenAI').checked = currentSettings.useOpenAI;

    if (currentSettings.hasCompanyKey) {
        document.getElementById('apiKeyRow').style.display = 'none';
        document.getElementById('useOpenAI').disabled = true;
        document.getElementById('useOpenAI').checked = true;

        const label = document.getElementById('useOpenAI').parentElement;
        if (label && label.tagName === 'LABEL') {
            label.innerHTML = `
                <input type="checkbox" id="useOpenAI" checked disabled>
                Use OpenAI for descriptions <span style="color: #28a745; font-weight: bold;">(Company Key Active)</span>
            `;
        }
    } else {
        document.getElementById('openAIKey').value = currentSettings.openAIKey;
        toggleAPIKeyField();
    }
}

document.getElementById('useOpenAI').addEventListener('change', toggleAPIKeyField);

function toggleAPIKeyField() {
    const apiKeyRow = document.getElementById('apiKeyRow');
    if (!currentSettings.hasCompanyKey) {
        apiKeyRow.style.display = document.getElementById('useOpenAI').checked ? 'block' : 'none';
    }
}

document.getElementById('selectFolder').addEventListener('click', async () => {
    const folder = await ipcRenderer.invoke('select-folder');
    if (folder) {
        document.getElementById('rootFolder').value = folder;
    }
});

document.getElementById('saveSettings').addEventListener('click', async () => {
    const settings = {
        rootFolder: document.getElementById('rootFolder').value,
        useOpenAI: document.getElementById('useOpenAI').checked,
        openAIKey: document.getElementById('openAIKey').value
    };

    await ipcRenderer.invoke('save-settings', settings);
    currentSettings = settings;
    showStatus('Settings saved successfully!', 'success');
});

document.getElementById('closeSettings').addEventListener('click', () => {
    window.close();
});

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
