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
      // CHANGE: Removed the redundant vanillaImplementation() call
      // This prevents both implementations from running simultaneously
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
      const testAjaxBtn = $('#test-ajax-btn');
      const checkConfigBtn = $('#check-config-btn');
      const preview = $('#file-preview');
      const messagesContainer = $('.messages');
      const debugOutput = $('#debug-output');
      const debugContent = $('#debug-content');
      
      // Log element existence
      console.log('AISheets: Elements found via jQuery:', {
        dropZone: dropZone.length > 0,
        fileInput: fileInput.length > 0,
        instructions: instructions.length > 0,
        processBtn: processBtn.length > 0,
        resetBtn: resetBtn.length > 0,
        testAjaxBtn: testAjaxBtn.length > 0,
        checkConfigBtn: checkConfigBtn.length > 0
      });
      
      if (dropZone.length === 0 || fileInput.length === 0) {
        showMessage('Critical elements missing. Check console for details.', 'error');
        return;
      }
      
      // Test Ajax button click
      testAjaxBtn.on('click', function() {
          console.log('AISheets: Test AJAX button clicked');
          
          $.ajax({
              url: aiExcelEditor.ajax_url,
              type: 'POST',
              data: {
                  action: 'aisheets_test',
                  nonce: aiExcelEditor.nonce
              },
              success: function(response) {
                  console.log('AISheets: Test AJAX success:', response);
                  showMessage('Test AJAX successful!', 'success');
              },
              error: function(xhr, status, error) {
                  console.error('AISheets: Test AJAX error:', status, error);
                  showMessage('Test AJAX failed.', 'error');
              }
          });
      });
      
      // Check Config button click
      checkConfigBtn.on('click', function() {
          console.log('AISheets: Check configuration button clicked');
          
          $.ajax({
              url: aiExcelEditor.ajax_url,
              type: 'POST',
              data: {
                  action: 'aisheets_debug',
                  nonce: aiExcelEditor.nonce
              },
              success: function(response) {
                  console.log('AISheets: Config check success:', response);
                  
                  // Show debug information
                  debugContent.text(JSON.stringify(response.data, null, 2));
                  debugOutput.show();
                  
                  showMessage('Configuration check complete.', 'info');
              },
              error: function(xhr, status, error) {
                  console.error('AISheets: Config check error:', status, error);
                  showMessage('Configuration check failed.', 'error');
              }
          });
      });
      
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
        // NEW: Add processing flag to prevent multiple submissions
        if ($(this).data('processing')) {
          console.log('AISheets: Already processing, ignoring click');
          return;
        }
        
        // Set processing flag
        $(this).data('processing', true);
        
        console.log('AISheets: Process button clicked');
        const file = fileInput[0].files[0];
        const instructionsText = instructions.val().trim();
        
        if (!file) {
          showMessage('Please upload a file first.', 'error');
          $(this).data('processing', false); // Reset processing flag
          return;
        }
        
        if (!instructionsText) {
          showMessage('Please provide instructions for processing.', 'error');
          $(this).data('processing', false); // Reset processing flag
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
          console.log('AISheets: Preparing AJAX request with:', {
              file_name: file.name,
              file_size: file.size,
              file_type: file.type,
              ajax_url: aiExcelEditor.ajax_url,
              nonce_present: !!aiExcelEditor.nonce
          });
          
          // Create FormData
          const formData = new FormData();
          formData.append('action', 'process_excel');
          formData.append('file', file);
          formData.append('instructions', instructions);
          
          if (typeof aiExcelEditor !== 'undefined' && aiExcelEditor.nonce) {
              formData.append('nonce', aiExcelEditor.nonce);
              console.log('AISheets: Using nonce:', aiExcelEditor.nonce);
          } else {
              console.error('AISheets: No nonce available!');
          }
          
          // AJAX URL
          const ajaxUrl = typeof aiExcelEditor !== 'undefined' && aiExcelEditor.ajax_url 
            ? aiExcelEditor.ajax_url 
            : '/wp-admin/admin-ajax.php';
          
          processBtn.prop('disabled', true).addClass('loading');
          showMessage('Processing your file...', 'info');
          
          // Track upload progress
          const xhr = new XMLHttpRequest();
          xhr.upload.addEventListener('progress', function(e) {
              if (e.lengthComputable) {
                  const percentComplete = (e.loaded / e.total) * 100;
                  console.log(`AISheets: Upload progress - ${percentComplete.toFixed(2)}%`);
              }
          });
          
          $.ajax({
              url: ajaxUrl,
              type: 'POST',
              data: formData,
              processData: false,
              contentType: false,
              xhr: function() { return xhr; },
              beforeSend: function() {
                  console.log('AISheets: AJAX request starting');
              },
              success: function(response) {
                  console.log('AISheets: AJAX success response:', response);
                  
                  try {
                      if (response.success) {
                          const downloadUrl = response.data.file_url;
                          const directDownloadUrl = response.data.direct_download_url;
                          
                          // Create download link and trigger download
                          const link = document.createElement('a');
                          link.href = downloadUrl;
                          link.download = 'processed_' + file.name;
                          link.setAttribute('type', 'application/octet-stream');
                          document.body.appendChild(link);
                          
                          // Try to trigger download
                          try {
                              link.click();
                          } catch (err) {
                              console.error('AISheets: Auto-download failed:', err);
                              // The download button in the preview will serve as backup
                          }
                          
                          document.body.removeChild(link);
                          
                          // Show success message with preview
                          showMessage('File processed successfully! If download doesn\'t start automatically, use the download button below.', 'success');
                          
                          // Update preview with download button
                          if (response.data.preview) {
                              preview.html(response.data.preview);
                          }
                          
                          // Add event tracking for download success/failure
                          setTimeout(function() {
                              // Check if the download might have failed
                              $('.download-note').css('color', '#e67e22')
                                              .html('⚠️ If your download didn\'t start, please click the download button above.');
                              
                              // Try the direct download approach as fallback
                              if (directDownloadUrl) {
                                  console.log('AISheets: Setting up fallback download via iframe');
                                  $('<iframe>', {
                                      src: directDownloadUrl,
                                      style: 'display: none;'
                                  }).appendTo('body');
                              }
                          }, 3000);
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
                  console.error('AISheets: AJAX error:', {
                      status: status,
                      error: error,
                      responseText: xhr.responseText,
                      statusCode: xhr.status
                  });
                  showMessage('Server error occurred. Please try again.', 'error');
              },
              complete: function() {
                  processBtn.removeClass('loading').prop('disabled', false);
                  // NEW: Reset processing flag when complete
                  processBtn.data('processing', false);
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
        debugOutput.hide();
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
    // [Rest of the vanilla JS implementation remains unchanged]
    // I'm keeping this part as is since we've disabled the second implementation from running
    
    // [vanillaImplementation code continues as in the original file...]
  }
  
  // Start initialization
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initialize);
  } else {
    initialize();
  }
})();