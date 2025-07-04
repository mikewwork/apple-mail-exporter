// DOM Elements
const accountSelect = document.getElementById('accountSelect');
const folderSelect = document.getElementById('folderSelect');
const emailCountInput = document.getElementById('emailCount');
const outputPathInput = document.getElementById('outputPath');
const selectFolderBtn = document.getElementById('selectFolderBtn');
const exportBtn = document.getElementById('exportBtn');
const progressCard = document.getElementById('progressCard');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const statusText = document.getElementById('statusText');
const resultsCard = document.getElementById('resultsCard');
const resultsContent = document.getElementById('resultsContent');
const openFolderBtn = document.getElementById('openFolderBtn');

// State
let mailAccounts = [];
let selectedOutputPath = '';

// Initialize the app
document.addEventListener('DOMContentLoaded', async () => {
    await loadMailAccounts();
    setupEventListeners();
});

async function loadMailAccounts() {
    try {
        statusText.textContent = 'Loading mail accounts...';
        mailAccounts = await window.electronAPI.getMailAccounts();
        
        // Populate account dropdown
        accountSelect.innerHTML = '<option value="">Select an account</option>';
        mailAccounts.forEach(account => {
            const option = document.createElement('option');
            option.value = account.name;
            option.textContent = account.name;
            accountSelect.appendChild(option);
        });
        
        statusText.textContent = 'Mail accounts loaded successfully';
    } catch (error) {
        console.error('Error loading mail accounts:', error);
        statusText.textContent = 'Error loading mail accounts. Please make sure Apple Mail is running.';
        statusText.className = 'status-text error';
    }
}

function setupEventListeners() {
    // Account selection
    accountSelect.addEventListener('change', (e) => {
        const selectedAccount = mailAccounts.find(acc => acc.name === e.target.value);
        populateFolders(selectedAccount);
        updateExportButton();
    });
    
    // Folder selection
    folderSelect.addEventListener('change', updateExportButton);
    
    // Email count
    emailCountInput.addEventListener('input', updateExportButton);
    
    // Output folder selection
    selectFolderBtn.addEventListener('click', selectOutputFolder);
    
    // Export button
    exportBtn.addEventListener('click', startExport);
    
    // Open folder button
    openFolderBtn.addEventListener('click', () => {
        if (selectedOutputPath) {
            window.electronAPI.openFolder(selectedOutputPath);
        }
    });
}

function populateFolders(account) {
    folderSelect.innerHTML = '<option value="">Select a folder</option>';
    folderSelect.disabled = !account;
    
    if (account && account.folders) {
        account.folders.forEach(folder => {
            const option = document.createElement('option');
            option.value = folder.name;
            option.textContent = folder.name;
            folderSelect.appendChild(option);
        });
    }
}

async function selectOutputFolder() {
    try {
        const folderPath = await window.electronAPI.selectOutputFolder();
        console.log('Renderer received output folder:', folderPath);
        if (folderPath) {
            selectedOutputPath = folderPath;
            outputPathInput.value = folderPath;
            updateExportButton();
        }
    } catch (error) {
        console.error('Error selecting output folder:', error);
        showError('Error selecting output folder');
    }
}

function updateExportButton() {
    const hasAccount = accountSelect.value !== '';
    const hasFolder = folderSelect.value !== '';
    const hasOutputPath = selectedOutputPath !== '';
    const hasEmailCount = emailCountInput.value > 0;
    exportBtn.disabled = !(hasAccount && hasFolder && hasOutputPath && hasEmailCount);
}

async function startExport() {
    const accountName = accountSelect.value;
    const folderName = folderSelect.value;
    const emailCount = parseInt(emailCountInput.value);
    if (!accountName || !folderName || !selectedOutputPath || emailCount <= 0) {
        showError('Please fill in all required fields');
        return;
    }
    
    // Show progress UI
    showProgressUI();
    
    // Disable form controls
    setFormControlsEnabled(false);
    
    try {
        const result = await window.electronAPI.exportEmails({
            accountName,
            folderName,
            outputPath: selectedOutputPath,
            emailCount
        });
        
        if (result.success) {
            showResults(result.pdfFiles);
        } else {
            showError('Export failed');
        }
    } catch (error) {
        console.error('Export error:', error);
        showError(`Export failed: ${error.message}`);
    } finally {
        // Re-enable form controls
        setFormControlsEnabled(true);
        hideProgressUI();
    }
}

function showProgressUI() {
    progressCard.style.display = 'block';
    progressFill.style.width = '0%';
    progressText.textContent = '0%';
    statusText.textContent = 'Starting export...';
    
    // Hide results if visible
    resultsCard.style.display = 'none';
}

function hideProgressUI() {
    progressCard.style.display = 'none';
}

function showResults(pdfFiles) {
    resultsCard.style.display = 'block';
    if (pdfFiles.length === 0) {
        resultsContent.innerHTML = `<h3 class="warning">No emails found in the selected date range.</h3>`;
        return;
    }
    const maxToShow = 10;
    const filesToShow = pdfFiles.slice(0, maxToShow);
    const html = `
        <h3 class="success">✅ Export completed successfully!</h3>
        <p>${pdfFiles.length} email(s) have been exported to PDF format.</p>
        <h4>Exported files:</h4>
        <div class="pdf-files-list">
            ${filesToShow.map(file => `
                <div class="pdf-file-item">
                    <span class="pdf-file-name">${file}</span>
                    <button class="mui-btn mui-btn-small" onclick="viewPDF('${file}')">
                        <span class="material-icons">visibility</span> View
                    </button>
                </div>
            `).join('')}
        </div>
        <p><strong>Location:</strong> ${selectedOutputPath}</p>
        ${pdfFiles.length > maxToShow ? `<p style='color:#888;'>Showing first 10 of ${pdfFiles.length} exported files.</p>` : ''}
    `;
    resultsContent.innerHTML = html;
}

function showError(message) {
    resultsCard.style.display = 'block';
    resultsContent.innerHTML = `
        <h3 class="error">❌ Export failed</h3>
        <p>${message}</p>
    `;
}

function setFormControlsEnabled(enabled) {
    accountSelect.disabled = !enabled;
    folderSelect.disabled = !enabled || accountSelect.value === '';
    emailCountInput.disabled = !enabled;
    selectFolderBtn.disabled = !enabled;
    exportBtn.disabled = !enabled;
    
    if (!enabled) {
        exportBtn.querySelector('.btn-text').style.display = 'none';
        exportBtn.querySelector('.btn-loading').style.display = 'flex';
    } else {
        exportBtn.querySelector('.btn-text').style.display = 'inline';
        exportBtn.querySelector('.btn-loading').style.display = 'none';
        updateExportButton();
    }
}

// Progress updates from main process
window.electronAPI.onProgress((event, data) => {
    if (data.type === 'progress') {
        updateProgress(data.percent, data.status);
    }
});

function updateProgress(percent, status) {
    progressFill.style.width = `${percent}%`;
    progressText.textContent = `${percent}%`;
    statusText.textContent = status || `Exporting... ${percent}%`;
}

// PDF viewing function
function viewPDF(filename) {
    const pdfPath = `${selectedOutputPath}/${filename}`;
    window.electronAPI.openPDF(pdfPath);
}

// Clean up listeners when page unloads
window.addEventListener('beforeunload', () => {
    window.electronAPI.removeAllListeners('export-progress');
});