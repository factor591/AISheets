// Dual implementation with both jQuery and vanilla JS
(function() {
    // Wait for DOM to be ready using both methods
    function initialize() {
      console.log('AISheets: Initializing...');
      
      // Check for jQuery
      if (typeof jQuery === 'undefined') {
        console.error('AISheets: jQuery not loaded! Using vanilla JS only.');
        vanillaImplementation();
      } else {
        console.log('AISheets: jQuery loaded. Version:', jQuery.fn.jquery);
        jqueryImplementation();
        vanillaImplementation(); // Run both for redundancy
      }
    }
    
    // jQuery implementation
    function jqueryImplementation() {
      jQuery(function($) {
        console.log('AISheets: jQuery ready');
        
        // Cache DOM elements
        const dropZone = $('#excel-dropzone');
        const fileInput = $('#file-upload');
        const instructions = $('#ai-instructions');
        const processBtn = $('#process-btn');
        const resetBtn = $('#reset-btn');
        const preview = $('#file-preview');
        const messagesContainer = $('.messages');
        
        // Log element existence
        console.log('AISheets: Elements found via jQuery:', {
          dropZone: dropZone.length > 0,
          fileInput: fileInput.length > 0,
          instructions: instructions.length > 0,
          processBtn: processBtn.length > 0,
          resetBtn: resetBtn.length > 0
        });
        
        if (dropZone.length === 0 || fileInput.length === 0) {
          showMessage('Critical elements missing. Check console for details.', 'error');
          return;
        }
        
        // Initialize file upload triggers
        dropZone.on('click', function(e) {
          console.log('AISheets: Dropzone clicked (jQuery)');
          if (!$(e.target).is('button')) {
            console.log('AISheets: Triggering file input click');
            fileInput.trigger('click');
          }
        });
        
        dropZone.on('dragover', function(e) {
          e.preventDefault();
          e.stopPropagation();
          $(this).addClass('dragover');
          console.log('AISheets: Dragover event');
        });
        
        dropZone.on('dragleave', function(e) {
          e.preventDefault();
          e.stopPropagation();
          $(this).removeClass('dragover');
          console.log('AISheets: Dragleave event');
        });
        
        dropZone.on('drop', function(e) {
          e.preventDefault();
          e.stopPropagation();
          console.log('AISheets: Drop event detected');
          $(this).removeClass('dragover');
          
          const files = e.originalEvent.dataTransfer.files;
          console.log('AISheets: Files dropped:', files.length);
          handleFiles(files);
        });
        
        fileInput.on('change', function() {
          console.log('AISheets: File input changed, files:', this.files);
          handleFiles(this.files);
        });
        
        // Process button click
        processBtn.on('click', function() {
          console.log('AISheets: Process button clicked');
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
        
        // Helper functions
        function handleFiles(files) {
          console.log('AISheets: handleFiles called with:', files.length, 'files');
          if (files.length > 0) {
            const file = files[0];
            console.log('AISheets: File selected:', file.name, file.size, file.type);
            
            // Check if we have the config available
            if (typeof aiExcelEditor === 'undefined' || !aiExcelEditor.max_file_size) {
              console.error('AISheets: Config missing. Using default values.');
              showMessage('Warning: Configuration incomplete. Using default values.', 'info');
              // Use default values
              var maxSize = 5 * 1024 * 1024;
              var allowedTypes = ['xlsx', 'xls', 'csv'];
            } else {
              var maxSize = aiExcelEditor.max_file_size;
              var allowedTypes = aiExcelEditor.allowed_types;
            }
            
            // Validate file size
            if (file.size > maxSize) {
              showMessage(`File size (${formatFileSize(file.size)}) exceeds the maximum limit of ${formatFileSize(maxSize)}`, 'error');
              return;
            }
            
            // Validate file type
            const fileExt = file.name.split('.').pop().toLowerCase();
            if (!allowedTypes.includes(fileExt)) {
              showMessage(`Invalid file type. Allowed types: ${allowedTypes.join(', ')}`, 'error');
              return;
            }
            
            // Update preview
            preview.html(`
              <div class="file-info">
                <p><strong>File:</strong> ${escapeHtml(file.name)}</p>
                <p><strong>Size:</strong> ${formatFileSize(file.size)}</p>
                <p><strong>Type:</strong> ${escapeHtml(file.type || 'unknown')}</p>
              </div>
            `);
            
            processBtn.prop('disabled', false);
            showMessage(`File "${file.name}" ready for processing.`, 'info');
            instructions.focus();
          }
        }
        
        function processFile(file, instructions) {
          console.log('AISheets: Processing file:', file.name);
          
          // Create FormData
          const formData = new FormData();
          formData.append('action', 'process_excel');
          formData.append('file', file);
          formData.append('instructions', instructions);
          
          if (typeof aiExcelEditor !== 'undefined' && aiExcelEditor.nonce) {
            formData.append('nonce', aiExcelEditor.nonce);
          }
          
          // AJAX URL
          const ajaxUrl = typeof aiExcelEditor !== 'undefined' && aiExcelEditor.ajax_url 
            ? aiExcelEditor.ajax_url 
            : '/wp-admin/admin-ajax.php';
          
          processBtn.prop('disabled', true).addClass('loading');
          showMessage('Processing your file...', 'info');
          
          $.ajax({
            url: ajaxUrl,
            type: 'POST',
            data: formData,
            processData: false,
            contentType: false,
            success: function(response) {
              console.log('AISheets: AJAX success response:', response);
              
              try {
                if (response.success) {
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
                  let errorMessage = 'Processing error';
                  if (response.data && response.data.message) {
                    errorMessage = response.data.message;
                  }
                  showMessage(errorMessage, 'error');
                  console.error('AISheets: Processing Error:', response.data);
                }
              } catch (e) {
                showMessage('Error parsing server response', 'error');
                console.error('AISheets: Parse error:', e, response);
              }
            },
            error: function(xhr, status, error) {
              console.error('AISheets: AJAX error:', status, error);
              showMessage('Server error occurred. Please try again.', 'error');
            },
            complete: function() {
              processBtn.removeClass('loading').prop('disabled', false);
            }
          });
        }
        
        function resetWorkspace() {
          fileInput.val('');
          instructions.val('');
          preview.html('');
          processBtn.prop('disabled', true);
          messagesContainer.empty();
          dropZone.removeClass('dragover');
          showMessage('Workspace reset.', 'info');
        }
        
        function showMessage(message, type = 'info') {
          console.log('AISheets: Showing message:', type, message);
          const messageDiv = $('<div>').addClass(`message message-${type}`).text(message);
          messagesContainer.html(messageDiv);
          
          if (type !== 'error') {
            setTimeout(() => {
              messageDiv.fadeOut(() => messageDiv.remove());
            }, 5000);
          }
        }
        
        function formatFileSize(bytes) {
          if (bytes === 0) return '0 Bytes';
          const k = 1024;
          const sizes = ['Bytes', 'KB', 'MB', 'GB'];
          const i = Math.floor(Math.log(bytes) / Math.log(k));
          return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        }
        
        function escapeHtml(unsafe) {
          return unsafe
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
        }
        
        // Add example instruction functionality
        $('.instruction-example').on('click', function() {
          const exampleText = $(this).text();
          instructions.val(exampleText);
        });
        
        // Initial setup
        processBtn.prop('disabled', true);
        showMessage('AISheets ready. Upload your Excel or CSV file.', 'info');
      });
    }
    
    // Vanilla JS implementation (backup)
    function vanillaImplementation() {
      console.log('AISheets: Vanilla JS implementation running');
      
      // Check if DOM is already loaded
      if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', setupVanillaJS);
      } else {
        setupVanillaJS();
      }
      
      function setupVanillaJS() {
        console.log('AISheets: Setting up vanilla JS implementation');
        
        // Get DOM elements
        const dropZone = document.getElementById('excel-dropzone');
        const fileInput = document.getElementById('file-upload');
        const instructions = document.getElementById('ai-instructions');
        const processBtn = document.getElementById('process-btn');
        const resetBtn = document.getElementById('reset-btn');
        const preview = document.getElementById('file-preview');
        const messagesContainer = document.querySelector('.messages');
        
        // Log element existence
        console.log('AISheets: Elements found via vanilla JS:', {
          dropZone: !!dropZone,
          fileInput: !!fileInput,
          instructions: !!instructions,
          processBtn: !!processBtn,
          resetBtn: !!resetBtn
        });
        
        // Skip if critical elements missing
        if (!dropZone || !fileInput) {
          console.error('AISheets: Critical elements missing');
          return;
        }
        
        // Add event listeners
        dropZone.addEventListener('click', function(e) {
          console.log('AISheets: Dropzone clicked (vanilla)');
          if (e.target.tagName !== 'BUTTON') {
            console.log('AISheets: Triggering file input click (vanilla)');
            fileInput.click();
          }
        });
        
        dropZone.addEventListener('dragover', function(e) {
          e.preventDefault();
          e.stopPropagation();
          this.classList.add('dragover');
        });
        
        dropZone.addEventListener('dragleave', function(e) {
          e.preventDefault();
          e.stopPropagation();
          this.classList.remove('dragover');
        });
        
        dropZone.addEventListener('drop', function(e) {
          e.preventDefault();
          e.stopPropagation();
          console.log('AISheets: Drop event detected (vanilla)');
          this.classList.remove('dragover');
          
          if (e.dataTransfer && e.dataTransfer.files.length > 0) {
            console.log('AISheets: Files dropped (vanilla):', e.dataTransfer.files.length);
            handleFilesVanilla(e.dataTransfer.files);
          }
        });
        
        fileInput.addEventListener('change', function() {
          console.log('AISheets: File input changed (vanilla), files:', this.files);
          handleFilesVanilla(this.files);
        });
        
        // Helper functions for vanilla JS
        function handleFilesVanilla(files) {
          console.log('AISheets: handleFilesVanilla with', files.length, 'files');
          if (files.length > 0) {
            const file = files[0];
            console.log('AISheets: File selected (vanilla):', file.name, file.size, file.type);
            
            // Preview file
            if (preview) {
              preview.innerHTML = `
                <div class="file-info">
                  <p><strong>File:</strong> ${file.name}</p>
                  <p><strong>Size:</strong> ${formatFileSizeVanilla(file.size)}</p>
                  <p><strong>Type:</strong> ${file.type || 'unknown'}</p>
                </div>
              `;
            }
            
            // Enable process button
            if (processBtn) {
              processBtn.disabled = false;
            }
            
            // Show message
            showMessageVanilla(`File "${file.name}" ready for processing.`, 'info');
            
            // Focus instructions
            if (instructions) {
              instructions.focus();
            }
          }
        }
        
        function formatFileSizeVanilla(bytes) {
          if (bytes === 0) return '0 Bytes';
          const k = 1024;
          const sizes = ['Bytes', 'KB', 'MB', 'GB'];
          const i = Math.floor(Math.log(bytes) / Math.log(k));
          return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        }
        
        function showMessageVanilla(message, type = 'info') {
          console.log('AISheets: Showing message (vanilla):', type, message);
          if (messagesContainer) {
            const messageDiv = document.createElement('div');
            messageDiv.className = `message message-${type}`;
            messageDiv.textContent = message;
            
            messagesContainer.innerHTML = '';
            messagesContainer.appendChild(messageDiv);
            
            if (type !== 'error') {
              setTimeout(() => {
                messageDiv.style.opacity = '0';
                setTimeout(() => {
                  if (messagesContainer.contains(messageDiv)) {
                    messagesContainer.removeChild(messageDiv);
                  }
                }, 500);
              }, 5000);
            }
          }
        }
      }
    }
    
    // Start initialization
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', initialize);
    } else {
      initialize();
    }
  })();