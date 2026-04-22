/**
 * Excel Image Insert Application - Frontend Logic
 * Handles UI interactions and communicates with Python backend via Eel
 */

// ===== Global State =====
let selectedExcelFile = null;
let selectedSheet = null;
let selectedImages = [];

// ===== Initialization =====
document.addEventListener('DOMContentLoaded', function() {
    console.log('Application initialized');
    addLog('Application started', 'info');
    
    // Update auto-size checkbox behavior
    document.getElementById('autoSizeCheckbox').addEventListener('change', function() {
        const container = document.getElementById('sizeInputsContainer');
        if (this.checked) {
            container.classList.add('hidden');
        } else {
            container.classList.remove('hidden');
        }
    });
});

// ===== Logging Utilities =====
function addLog(message, level = 'info') {
    const timestamp = new Date().toLocaleTimeString('en-US', {
        hour12: false,
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
    });
    
    const logOutput = document.getElementById('logOutput');
    const logLine = document.createElement('p');
    logLine.className = `log-line log-${level}`;
    logLine.textContent = `[${timestamp}] [${level.toUpperCase()}] ${message}`;
    logOutput.appendChild(logLine);
    
    // Auto-scroll to bottom
    logOutput.parentElement.scrollTop = logOutput.parentElement.scrollHeight;
    
    console.log(`[${level.toUpperCase()}] ${message}`);
}

// ===== Excel File Operations =====
function selectExcelFile() {
    addLog('Opening file selection dialog for Excel file...', 'info');
    
    eel.select_excel_file()(function(result) {
        if (result.success) {
            selectedExcelFile = result.file_path;
            document.getElementById('excelFilePath').value = result.file_path;
            document.getElementById('sheetsContainer').classList.remove('hidden');
            
            // Populate sheets
            const sheetSelect = document.getElementById('sheetSelect');
            sheetSelect.innerHTML = '<option value="">Choose a sheet...</option>';
            
            result.sheets.forEach(sheet => {
                const option = document.createElement('option');
                option.value = sheet;
                option.textContent = sheet;
                sheetSelect.appendChild(option);
            });
            
            addLog(`Excel file opened: ${result.file_path}`, 'success');
            addLog(`Available sheets: ${result.sheets.join(', ')}`, 'info');
        } else {
            addLog(`Failed to open Excel file: ${result.message}`, 'error');
        }
    });
}

function getSelectedSheet() {
    const sheet = document.getElementById('sheetSelect').value;
    if (!sheet) {
        addLog('Please select a sheet first', 'warning');
        return null;
    }
    return sheet;
}

// ===== Range Selection =====
function selectRangeFromExcel() {
    if (!selectedExcelFile) {
        addLog('Please select an Excel file first', 'warning');
        return;
    }
    
    addLog('Click on cells in Excel to select the range, then this function will capture it...', 'info');
    
    eel.select_range_from_excel()(function(result) {
        if (result.success) {
            const cellPos = result.selected_range;
            document.getElementById('cellPosition').value = cellPos;
            addLog(`Cell range selected from Excel: ${cellPos}`, 'success');
        } else {
            addLog(`Failed to select range: ${result.message}`, 'warning');
        }
    });
}

// ===== Image Selection =====
function selectImages() {
    if (!getSelectedSheet()) {
        return;
    }
    
    addLog('Opening file selection dialog for images...', 'info');
    
    eel.select_images()(function(result) {
        if (result.success) {
            const sheet = getSelectedSheet();
            
            result.file_paths.forEach(filePath => {
                eel.add_image(
                    filePath,
                    sheet,
                    document.getElementById('cellPosition').value || 'A1'
                )(function(addResult) {
                    if (addResult.success) {
                        addLog(`Image added: ${getFilename(filePath)}`, 'success');
                        updateImagesList();
                    } else {
                        addLog(`Failed to add image: ${addResult.message}`, 'error');
                    }
                });
            });
        } else {
            addLog(`Image selection cancelled or failed`, 'warning');
        }
    });
}

function removeImage(index) {
    eel.remove_image(index)(function(result) {
        if (result.success) {
            addLog('Image removed', 'info');
            updateImagesList();
        } else {
            addLog(`Failed to remove image: ${result.message}`, 'error');
        }
    });
}

function clearImages() {
    if (confirm('Are you sure you want to clear all images?')) {
        eel.clear_images()(function(result) {
            if (result.success) {
                addLog('All images cleared', 'info');
                updateImagesList();
            } else {
                addLog(`Failed to clear images: ${result.message}`, 'error');
            }
        });
    }
}

function updateImagesList() {
    eel.get_images()(function(result) {
        const container = document.getElementById('imagesList');
        
        if (result.images.length === 0) {
            container.innerHTML = '<p class="empty-message">No images selected</p>';
            return;
        }
        
        container.innerHTML = '';
        
        result.images.forEach((image, index) => {
            const itemDiv = document.createElement('div');
            itemDiv.className = 'image-item';
            
            const infoDiv = document.createElement('div');
            infoDiv.className = 'image-info';
            
            const nameP = document.createElement('p');
            nameP.className = 'image-name';
            nameP.textContent = `📷 ${image.filename}`;
            
            const detailsP = document.createElement('p');
            detailsP.className = 'image-details';
            detailsP.textContent = `Sheet: ${image.sheet_name} | Cell: ${image.cell_position} | Size: ${image.file_size_mb} MB`;
            
            infoDiv.appendChild(nameP);
            infoDiv.appendChild(detailsP);
            
            const removeBtn = document.createElement('button');
            removeBtn.className = 'image-remove-btn';
            removeBtn.textContent = 'Remove';
            removeBtn.onclick = () => removeImage(index);
            
            itemDiv.appendChild(infoDiv);
            itemDiv.appendChild(removeBtn);
            
            container.appendChild(itemDiv);
        });
    });
}

// ===== Image Insertion =====
function insertImages() {
    const sheet = getSelectedSheet();
    if (!sheet) {
        return;
    }
    
    if (!selectedExcelFile) {
        addLog('Please select an Excel file first', 'warning');
        return;
    }
    
    eel.get_images()(function(result) {
        if (result.images.length === 0) {
            addLog('No images to insert', 'warning');
            return;
        }
        
        addLog(`Starting image insertion (${result.images.length} image(s))...`, 'info');
        showProgressBar(true);
        
        // Disable button during insertion
        const insertBtn = document.getElementById('insertBtn');
        insertBtn.disabled = true;
        
        // Update progress
        const totalSteps = result.images.length + 2; // Adding 2 for initialization and saving
        let currentStep = 1;
        
        updateProgress(currentStep, totalSteps, 'Initializing...');
        
        const progressInterval = setInterval(() => {
            currentStep = Math.min(currentStep + 0.5, totalSteps - 1);
            updateProgress(currentStep, totalSteps, `Processing (${Math.floor(currentStep)}/${totalSteps})...`);
        }, 300);
        
        eel.insert_images()(function(insertResult) {
            clearInterval(progressInterval);
            insertBtn.disabled = false;
            
            if (insertResult.success) {
                currentStep = totalSteps;
                updateProgress(currentStep, totalSteps, 'Complete! ✓', 'success');
                addLog(`Successfully inserted all images!`, 'success');
                addLog('Excel file saved', 'success');
                updateImagesList();
                
                // Hide progress after 2 seconds
                setTimeout(() => {
                    showProgressBar(false);
                }, 2000);
            } else {
                updateProgress(totalSteps, totalSteps, 'Failed! ✗', 'error');
                addLog(`Image insertion failed: ${insertResult.message}`, 'error');
                
                // Hide progress after 3 seconds
                setTimeout(() => {
                    showProgressBar(false);
                }, 3000);
            }
        });
    });
}

// ===== Utility Functions =====
function getFilename(filepath) {
    return filepath.split('\\').pop().split('/').pop();
}

// ===== Progress Bar Functions =====
function showProgressBar(show) {
    const container = document.getElementById('progressContainer');
    if (show) {
        container.classList.remove('hidden');
    } else {
        container.classList.add('hidden');
    }
}

function updateProgress(current, total, text, state = null) {
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    
    const percentage = (current / total) * 100;
    progressFill.style.width = percentage + '%';
    progressText.textContent = text;
    
    // Update state classes
    progressFill.classList.remove('success', 'error');
    if (state === 'success') {
        progressFill.classList.add('success');
    } else if (state === 'error') {
        progressFill.classList.add('error');
    }
}

// ===== Floating Panel Functions =====
function toggleFloatingPanel() {
    const content = document.getElementById('floatingPanelContent');
    const header = document.querySelector('.floating-panel-header');
    const btn = document.querySelector('.btn-close');
    
    if (content.style.display === 'none') {
        // Maximize
        content.style.display = 'block';
        btn.textContent = '−';
    } else {
        // Minimize
        content.style.display = 'none';
        btn.textContent = '+';
    }
}

// Make floating panel draggable
const floatingPanel = document.querySelector('.floating-panel');
const header = document.querySelector('.floating-panel-header');

let isDragging = false;
let offsetX = 0;
let offsetY = 0;

header.addEventListener('mousedown', (e) => {
    if (e.target.classList.contains('btn-close')) return;
    
    isDragging = true;
    offsetX = e.clientX - floatingPanel.getBoundingClientRect().left;
    offsetY = e.clientY - floatingPanel.getBoundingClientRect().top;
    header.style.cursor = 'grabbing';
});

document.addEventListener('mousemove', (e) => {
    if (!isDragging) return;
    
    floatingPanel.style.right = 'auto';
    floatingPanel.style.bottom = 'auto';
    floatingPanel.style.left = (e.clientX - offsetX) + 'px';
    floatingPanel.style.top = (e.clientY - offsetY) + 'px';
});

document.addEventListener('mouseup', () => {
    if (isDragging) {
        isDragging = false;
        header.style.cursor = 'grab';
    }
});
