/* ==========================================
   6. FILES MODULE (Multiple File Upload & Management)
   ========================================== */
const FilesModule = (function() {

    const fileInput = document.getElementById('document');
    const fileNameDisplay = document.getElementById('fileNameDisplay');
    const filesContainer = document.getElementById('uploadedFilesContainer');
    const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10 MB limit

    function init() {
        if (!fileInput) return;

        fileInput.addEventListener('change', function() {
            let invalidFiles = [];
            
            for (let i = 0; i < this.files.length; i++) {
                if (this.files[i].size > MAX_FILE_SIZE) {
                    invalidFiles.push(this.files[i].name);
                } else {
                    dt.items.add(this.files[i]);
                }
            }

            if (invalidFiles.length > 0) {
                alert(`The following file(s) exceed the 10MB limit and were not added:\n${invalidFiles.join('\n')}`);
            }

            this.files = dt.files;
            renderFilesList();
        });
    }

    // Render the list of selected files
    function renderFilesList() {
        if (!filesContainer || !fileNameDisplay) return;
        filesContainer.innerHTML = '';
        
        if (dt.files.length > 0) {
            fileNameDisplay.style.display = 'none';
            
            for (let i = 0; i < dt.files.length; i++) {
                const file = dt.files[i];
                const fileItem = document.createElement('div');
                fileItem.className = 'file-item';
                
                let iconClass = 'fa-file-lines';
                if (file.type.includes('pdf')) iconClass = 'fa-file-pdf';
                if (file.type.includes('word') || file.name.endsWith('.docx')) iconClass = 'fa-file-word';
                
                fileItem.innerHTML = `
                    <div class="file-item-info">
                        <i class="fa-solid ${iconClass}"></i> 
                        <span>${file.name} <small style="color:#888;">(${(file.size / (1024 * 1024)).toFixed(2)} MB)</small></span>
                    </div>
                    <i class="fa-solid fa-xmark file-item-remove" onclick="FilesModule.removeFile(${i}, event)" title="Remove File"></i>
                `;
                
                filesContainer.appendChild(fileItem);
            }
        } else {
            fileNameDisplay.style.display = 'block';
        }
    }

    // Remove a specific file from the list
    function removeFile(index, event) {
        event.stopPropagation();
        dt.items.remove(index);
        if (fileInput) fileInput.files = dt.files;
        renderFilesList();
    }

    // Reset files selection
    function clearFiles() {
        dt = new DataTransfer();
        if (fileInput) fileInput.files = dt.files;
        renderFilesList();
    }

    return {
        init,
        removeFile,
        clearFiles,
        getFiles: () => dt.files
    };
})();