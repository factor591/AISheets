jQuery(document).ready(function($) {
    // Debug initialization
    console.log('AISheets initialized');
    console.log('Config:', aiExcelEditor);
    
    // Cache DOM elements
    const dropZone = $('#excel-dropzone');
    const fileInput = $('#file-upload');
    const instructions = $('#ai-instructions');
    const processBtn = $('#process-btn');
    const resetBtn = $('#reset-btn');
    const preview = $('#file-preview');
    const messagesContainer = $('.messages');
    
    // Initialize file upload triggers
    dropZone.on('click', function(e) {
        // Prevent click if target is a button
        if (!$(e.target).is('button')) {
            fileInput.click();
        }
    });
    
    // File upload handling
    dropZone.on('dragover', function(e) {
        e.preventDefault();
        e.stopPropagation();
        $(this).addClass('dragover');
    });

    dropZone.on('dragleave', function(e) {
        e.preventDefault();
        e.stopPropagation();
        $(this).removeClass('dragover');
    });

    dropZone.on('drop', function(e) {
        e.preventDefault();
        e.stopPropagation();
        $(this).removeClass('dragover');
        
        const files = e.originalEvent.dataTransfer.files;
        handleFiles(files);
    });

    fileInput.on('change', function(e) {
        console.log('File input changed:', this.files);
        handleFiles(this.files);
    });

    // Handle file selection with validation
    function handleFiles(files) {
        console.log('handleFiles called with:', files);
        if (files.length > 0) {
            const file = files[0];
            console.log('File selected:', file);
            
            // Check if we have the config available
            if (!aiExcelEditor || !aiExcelEditor.max_file_size) {
                console.error('AISheets config missing. Check if wp_localize_script is working properly.');
                showMessage('Configuration error. Please contact the administrator.', 'error');
                return;
            }
            
            // Validate file size
            if (file.size > aiExcelEditor.max_file_size) {
                showMessage(`File size (${formatFileSize(file.size)}) exceeds the maximum limit of ${formatFileSize(aiExcelEditor.max_file_size)}`, 'error');
                return;
            }

            // Validate file type
            if (!validateFileType(file)) {
                showMessage(`Invalid file type. Allowed types: ${aiExcelEditor.allowed_types.join(', ')}`, 'error');
                return;
            }

            preview.html(`
                <div class="file-info">
                    <p><strong>File:</strong> ${escapeHtml(file.name)}</p>
                    <p><strong>Size:</strong> ${formatFileSize(file.size)}</p>
                    <p><strong>Type:</strong> ${escapeHtml(file.type || 'unknown')}</p>
                </div>
            `);

            processBtn.prop('disabled', false);
            instructions.focus();
        }
    }

    // Process button click
    processBtn.on('click', function() {
        console.log('Process button clicked');
        const file = fileInput[0].files[0];
        const instructionsText = instructions.val().trim();

        if (!file) {
            showMessage('Please upload a file first.', 'error');
            return;
        }

        if (!instructionsText) {
            showMessage('Please provide instructions for processing.', 'error');
            return;
        }

        processFile(file, instructionsText);
    });

    // Reset button click
    resetBtn.on('click', function() {
        resetWorkspace();
    });

    function processFile(file, instructions) {
        console.log('Processing file:', file.name);
        
        const formData = new FormData();
        formData.append('action', 'process_excel');
        formData.append('file', file);
        formData.append('instructions', instructions);
        formData.append('nonce', aiExcelEditor.nonce);

        processBtn.prop('disabled', true).addClass('loading');
        showMessage('Processing your file...', 'info');

        // Add a progress indicator for long operations
        let processTime = 0;
        const progressTimer = setInterval(() => {
            processTime += 1;
            if (processTime % 3 === 0) {
                showMessage(`Processing your file... (${processTime}s)`, 'info');
            }
        }, 1000);

        $.ajax({
            url: aiExcelEditor.ajax_url,
            type: 'POST',
            data: formData,
            processData: false,
            contentType: false,
            success: function(response) {
                clearInterval(progressTimer);
                console.log('AJAX response:', response);
                
                try {
                    if (response.success) {
                        // Handle success
                        const downloadUrl = response.data.file_url;
                        const link = document.createElement('a');
                        link.href = downloadUrl;
                        link.download = 'processed_' + file.name;
                        document.body.appendChild(link);
                        link.click();
                        document.body.removeChild(link);
                        
                        showMessage('File processed and downloaded successfully!', 'success');
                        
                        if (response.data.preview) {
                            preview.html(response.data.preview);
                        }
                    } else {
                        // Handle error with detailed information
                        const errorDetails = response.data;
                        let errorMessage = `Error: ${errorDetails.message}\n`;
                        if (errorDetails.code) {
                            errorMessage += `Code: ${errorDetails.code}\n`;
                        }
                        if (errorDetails.details) {
                            errorMessage += `\nDetails:\n`;
                            if (typeof errorDetails.details === 'object') {
                                Object.entries(errorDetails.details).forEach(([key, value]) => {
                                    errorMessage += `${key}: ${JSON.stringify(value, null, 2)}\n`;
                                });
                            } else {
                                errorMessage += errorDetails.details;
                            }
                        }
                        showMessage(errorMessage, 'error');
                        console.error('Processing Error:', errorDetails);
                    }
                } catch (e) {
                    showMessage('Error parsing server response: ' + e.message, 'error');
                    console.error('Parse error:', e, response);
                }
            },
            error: function(xhr, status, error) {
                clearInterval(progressTimer);
                let errorMessage = 'Server error occurred.\n';
                try {
                    const response = JSON.parse(xhr.responseText);
                    if (response.data && response.data.message) {
                        errorMessage += `Message: ${response.data.message}\n`;
                    }
                    if (response.data && response.data.details) {
                        errorMessage += `Details: ${JSON.stringify(response.data.details, null, 2)}`;
                    }
                    console.error('Server Error Details:', response);
                } catch (e) {
                    errorMessage += `Status: ${status}\nError: ${error}`;
                    console.error('Ajax error:', {xhr, status, error});
                }
                showMessage(errorMessage, 'error');
            },
            complete: function() {
                processBtn.removeClass('loading').prop('disabled', false);
            }
        });
    }

    // Helper functions
    function validateFileType(file) {
        const validTypes = [
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'text/csv',
            'application/csv',
            'text/plain'
        ];

        const fileName = file.name.toLowerCase();
        const fileExt = fileName.split('.').pop();

        return validTypes.includes(file.type) || 
               aiExcelEditor.allowed_types.includes(fileExt);
    }

    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    function showMessage(message, type = 'info') {
        console.log('Showing message:', type, message);
        const messageDiv = $('<div>')
            .addClass(`message message-${type}`);
        
        // Handle multiline messages
        if (message.includes('\n')) {
            const lines = message.split('\n').filter(line => line.trim());
            lines.forEach(line => {
                if (line.trim()) {
                    messageDiv.append($('<p>').text(line));
                }
            });
        } else {
            messageDiv.text(message);
        }
        
        messagesContainer.html(messageDiv);
        
        if (type !== 'error') {
            setTimeout(() => {
                messageDiv.fadeOut(() => messageDiv.remove());
            }, 5000);
        }

        // Scroll to message if it's not visible
        if (!isElementInViewport(messageDiv[0])) {
            $('html, body').animate({
                scrollTop: messageDiv.offset().top - 20
            }, 500);
        }
    }

    function resetWorkspace() {
        fileInput.val('');
        instructions.val('');
        preview.html('');
        processBtn.prop('disabled', true);
        messagesContainer.empty();
        dropZone.removeClass('dragover');
    }

    function escapeHtml(unsafe) {
        return unsafe
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    function isElementInViewport(el) {
        const rect = el.getBoundingClientRect();
        return (
            rect.top >= 0 &&
            rect.left >= 0 &&
            rect.bottom <= (window.innerHeight || document.documentElement.clientHeight) &&
            rect.right <= (window.innerWidth || document.documentElement.clientWidth)
        );
    }

    // Add example instruction functionality if present
    $('.instruction-example').on('click', function() {
        const exampleText = $(this).text();
        instructions.val(exampleText);
    });

    // Initial setup
    processBtn.prop('disabled', true);
});