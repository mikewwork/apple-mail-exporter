const { ipcRenderer } = require('electron');

// DOM elements
const folderSelect = document.getElementById('folderSelect');
const outputFolder = document.getElementById('outputFolder');
const browseBtn = document.getElementById('browseBtn');
const refreshBtn = document.getElementById('refreshBtn');
const exportBtn = document.getElementById('exportBtn');
const status = document.getElementById('status');
const progress = document.getElementById('progress');
const progressFill = document.querySelector('.progress-fill');
const progressText = document.getElementById('progressText');

// Event listeners
refreshBtn.addEventListener('click', loadFolders);
folderSelect.addEventListener('change', validateForm);
outputFolder.addEventListener('click', browseOutputFolder);
browseBtn.addEventListener('click', browseOutputFolder);
exportBtn.addEventListener('click', exportEmails);

// Load folders when page loads
document.addEventListener('DOMContentLoaded', loadFolders);

// Load available folders
async function loadFolders() {
    try {
        updateStatus('Scanning mail folders...');
        refreshBtn.disabled = true;
        
        const folders = await ipcRenderer.invoke('scan-mail-folders');
        
        folderSelect.innerHTML = '<option value="">Choose a folder...</option>';
        
        if (folders.length === 0) {
            updateStatus('No mail folders found. Make sure you have emails in your Apple Mail accounts.');
        } else {
            folders.forEach(folder => {
                const option = document.createElement('option');
                option.value = folder;
                option.textContent = folder;
                folderSelect.appendChild(option);
            });
            
            updateStatus(`Found ${folders.length} mail folders. Select a folder and output location to begin exporting emails.`);
        }
        
    } catch (error) {
        updateStatus(`Error scanning mail folders: ${error.message}`);
        console.error('Error scanning mail folders:', error);
    } finally {
        refreshBtn.disabled = false;
    }
}

// Browse for output folder
async function browseOutputFolder() {
    try {
        const selectedPath = await ipcRenderer.invoke('select-folder');
        if (selectedPath) {
            outputFolder.value = selectedPath;
            validateForm();
        }
    } catch (error) {
        updateStatus(`Error selecting output folder: ${error.message}`);
        console.error('Error selecting output folder:', error);
    }
}

// Validate form and enable/disable export button
function validateForm() {
    const selectedFolder = folderSelect.value;
    const outputPath = outputFolder.value;
    
    const isValid = selectedFolder && outputPath;
    exportBtn.disabled = !isValid;
    
    if (selectedFolder && outputPath) {
        updateStatus(`Ready to export from "${selectedFolder}" to "${outputPath}"`);
    } else if (selectedFolder) {
        updateStatus(`Selected folder: ${selectedFolder}. Please select an output location.`);
    } else if (outputPath) {
        updateStatus(`Output location: ${outputPath}. Please select a mail folder.`);
    } else {
        updateStatus('Select a mail folder and output location to begin exporting emails.');
    }
}

// Export emails from selected folder
async function exportEmails() {
    const selectedFolder = folderSelect.value;
    const outputPath = outputFolder.value;
    
    if (!selectedFolder || !outputPath) return;
    
    try {
        exportBtn.disabled = true;
        refreshBtn.disabled = true;
        browseBtn.disabled = true;
        showProgress();
        
        updateProgress(0, 'Starting export...');
        
        const result = await ipcRenderer.invoke('export-emails', selectedFolder, outputPath);
        
        updateProgress(100, 'Export completed!');
        updateStatus(result.message);
        
        setTimeout(() => {
            hideProgress();
            exportBtn.disabled = false;
            refreshBtn.disabled = false;
            browseBtn.disabled = false;
        }, 2000);
        
    } catch (error) {
        updateStatus(`Error exporting emails: ${error.message}`);
        console.error('Error exporting emails:', error);
        hideProgress();
        exportBtn.disabled = false;
        refreshBtn.disabled = false;
        browseBtn.disabled = false;
    }
}

// Update status message
function updateStatus(message) {
    status.innerHTML = `<p>${message}</p>`;
}

// Show progress bar
function showProgress() {
    progress.classList.remove('hidden');
}

// Hide progress bar
function hideProgress() {
    progress.classList.add('hidden');
}

// Update progress bar
function updateProgress(percentage, text) {
    progressFill.style.width = `${percentage}%`;
    progressText.textContent = text;
}

// Listen for progress updates from main process
ipcRenderer.on('progress-update', (event, percentage, text) => {
    updateProgress(percentage, text);
});