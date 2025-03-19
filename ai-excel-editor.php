<?php
/*
Plugin Name: AI Excel Editor
Plugin URI: https://yourwebsite.com/ai-excel-editor
Description: AI-powered Excel file editor using OpenAI
Version: 1.0.0
Author: Your Name
License: GPL v2 or later
*/

// Prevent direct access
if (!defined('ABSPATH')) {
    exit;
}

// Define plugin constants
define('AI_EXCEL_EDITOR_VERSION', '1.0.0');
define('AI_EXCEL_EDITOR_PLUGIN_DIR', plugin_dir_path(__FILE__));
define('AI_EXCEL_EDITOR_PLUGIN_URL', plugin_dir_url(__FILE__));

// Debug helper function
function aisheets_debug($message, $data = null) {
    if (WP_DEBUG === true) {
        $timestamp = date('Y-m-d H:i:s');
        $log_message = "[AISheets {$timestamp}] {$message}";
        
        if ($data !== null) {
            if (is_array($data) || is_object($data)) {
                $log_message .= " - Data: " . print_r($data, true);
            } else {
                $log_message .= " - Data: " . $data;
            }
        }
        
        error_log($log_message);
    }
}

class AI_Excel_Editor {
    private $openai_api_key;
    private $upload_dir;
    private $allowed_file_types;
    private $max_file_size;

    public function __construct() {
        aisheets_debug('Plugin initializing');
        
        $this->openai_api_key = get_option('ai_excel_editor_openai_key');
        $this->allowed_file_types = array('xlsx', 'xls', 'csv');
        $this->max_file_size = 5 * 1024 * 1024; // 5MB
        
        // Setup hooks
        add_action('init', array($this, 'init'));
        add_action('admin_menu', array($this, 'add_admin_menu'));
        add_action('admin_init', array($this, 'register_settings'));
        add_action('wp_enqueue_scripts', array($this, 'enqueue_scripts'));
        add_shortcode('ai_excel_editor', array($this, 'render_editor'));
        
        // AJAX handlers
        add_action('wp_ajax_process_excel', array($this, 'handle_excel_processing'));
        add_action('wp_ajax_nopriv_process_excel', array($this, 'handle_unauthorized'));

        // Cleanup schedule
        if (!wp_next_scheduled('ai_excel_editor_cleanup')) {
            wp_schedule_event(time(), 'hourly', 'ai_excel_editor_cleanup');
        }
        add_action('ai_excel_editor_cleanup', array($this, 'cleanup_old_files'));
        
        // Debug actions
        add_action('wp_footer', array($this, 'debug_info'));
        
        aisheets_debug('Plugin initialized', array(
            'openai_key_set' => !empty($this->openai_api_key),
            'allowed_types' => $this->allowed_file_types,
            'max_file_size' => $this->max_file_size
        ));
    }
    
    // Debug function in footer
    public function debug_info() {
        global $post;
        if (is_a($post, 'WP_Post') && has_shortcode($post->post_content, 'ai_excel_editor')) {
            echo '<script>
                console.log("AISheets Debug Info:");
                console.log("Plugin URL:", "' . AI_EXCEL_EDITOR_PLUGIN_URL . '");
                console.log("jQuery Loaded:", typeof jQuery !== "undefined");
                console.log("WordPress AJAX URL:", "' . admin_url('admin-ajax.php') . '");
                
                // Test DOM elements
                document.addEventListener("DOMContentLoaded", function() {
                    console.log("DOM fully loaded");
                    console.log("Dropzone element:", document.getElementById("excel-dropzone"));
                    console.log("File input element:", document.getElementById("file-upload"));
                    
                    // Add enhanced debugging
                    if (document.getElementById("excel-dropzone")) {
                        document.getElementById("excel-dropzone").addEventListener("click", function(e) {
                            console.log("Direct click event on dropzone");
                            // Continue with normal handling
                        });
                    }
                });
            </script>';
        }
    }

    public function init() {
        aisheets_debug('Plugin init method called');
        
        try {
            // Create upload directory
            $upload_dir = wp_upload_dir();
            $this->upload_dir = $upload_dir['basedir'] . '/ai-excel-editor';
            
            aisheets_debug('Upload directory info', array(
                'path' => $this->upload_dir,
                'exists' => file_exists($this->upload_dir),
                'basedir_writable' => is_writable($upload_dir['basedir'])
            ));
            
            if (!file_exists($this->upload_dir)) {
                aisheets_debug('Creating upload directory');
                $mkdir_result = wp_mkdir_p($this->upload_dir);
                
                if (!$mkdir_result) {
                    aisheets_debug('Failed to create directory', error_get_last());
                    throw new Exception('Failed to create upload directory: ' . print_r(error_get_last(), true));
                }
                
                // Set permissions explicitly
                chmod($this->upload_dir, 0755);
                
                file_put_contents($this->upload_dir . '/.htaccess', 'deny from all');
                file_put_contents($this->upload_dir . '/index.php', '<?php // Silence is golden');
                
                aisheets_debug('Upload directory created with permissions', substr(sprintf('%o', fileperms($this->upload_dir)), -4));
            }

            // Create processing directory
            $processing_dir = $this->upload_dir . '/processing';
            
            aisheets_debug('Processing directory info', array(
                'path' => $processing_dir,
                'exists' => file_exists($processing_dir)
            ));
            
            if (!file_exists($processing_dir)) {
                aisheets_debug('Creating processing directory');
                $mkdir_result = wp_mkdir_p($processing_dir);
                
                if (!$mkdir_result) {
                    aisheets_debug('Failed to create processing directory', error_get_last());
                    throw new Exception('Failed to create processing directory: ' . print_r(error_get_last(), true));
                }
                
                // Set permissions explicitly
                chmod($processing_dir, 0755);
                
                file_put_contents($processing_dir . '/.htaccess', 'deny from all');
                file_put_contents($processing_dir . '/index.php', '<?php // Silence is golden');
                
                aisheets_debug('Processing directory created with permissions', substr(sprintf('%o', fileperms($processing_dir)), -4));
            }
            
            // Test write permissions
            $test_file = $this->upload_dir . '/test_' . uniqid() . '.txt';
            $write_result = file_put_contents($test_file, 'Test write permissions');
            
            if ($write_result === false) {
                aisheets_debug('Failed to write test file', error_get_last());
            } else {
                aisheets_debug('Successfully wrote test file', array(
                    'path' => $test_file,
                    'bytes_written' => $write_result
                ));
                unlink($test_file); // Clean up
            }
            
        } catch (Exception $e) {
            aisheets_debug('Init Error: ' . $e->getMessage(), $e->getTraceAsString());
        }
    }

    public function add_admin_menu() {
        add_menu_page(
            'AI Excel Editor Settings',
            'AI Excel Editor',
            'manage_options',
            'ai-excel-editor',
            array($this, 'render_admin_page'),
            'dashicons-media-spreadsheet'
        );
    }

    public function register_settings() {
        register_setting('ai_excel_editor_settings', 'ai_excel_editor_openai_key', array(
            'sanitize_callback' => array($this, 'sanitize_api_key')
        ));
    }

    public function sanitize_api_key($key) {
        return sanitize_text_field(trim($key));
    }

    public function enqueue_scripts() {
        // Only enqueue on pages where shortcode is used
        global $post;
        if (is_a($post, 'WP_Post') && has_shortcode($post->post_content, 'ai_excel_editor')) {
            aisheets_debug('Enqueuing scripts and styles for post ID: ' . $post->ID);
            
            wp_enqueue_style(
                'ai-excel-editor',
                AI_EXCEL_EDITOR_PLUGIN_URL . 'css/style.css',
                array(),
                AI_EXCEL_EDITOR_VERSION
            );

            wp_enqueue_script(
                'ai-excel-editor',
                AI_EXCEL_EDITOR_PLUGIN_URL . 'js/main.js',
                array('jquery'),
                AI_EXCEL_EDITOR_VERSION,
                true
            );

            wp_localize_script('ai-excel-editor', 'aiExcelEditor', array(
                'ajax_url' => admin_url('admin-ajax.php'),
                'nonce' => wp_create_nonce('ai_excel_editor_nonce'),
                'max_file_size' => $this->max_file_size,
                'allowed_types' => $this->allowed_file_types
            ));
            
            aisheets_debug('Scripts and styles enqueued', array(
                'css_url' => AI_EXCEL_EDITOR_PLUGIN_URL . 'css/style.css',
                'js_url' => AI_EXCEL_EDITOR_PLUGIN_URL . 'js/main.js',
                'nonce_created' => true
            ));
        }
    }

    public function render_editor() {
        aisheets_debug('Rendering editor shortcode');
        
        if (!$this->openai_api_key) {
            aisheets_debug('OpenAI API key not configured');
            return '<div class="error">Please configure the OpenAI API key in the plugin settings.</div>';
        }

        ob_start();
        ?>
        <div class="ai-excel-editor-container">
            <div class="messages"></div>
            
            <div class="upload-section">
                <div id="excel-dropzone" class="dropzone" tabindex="0">
                    <div class="dz-message">
                        <h3><?php _e('Drag and Drop Your Excel File Here', 'ai-excel-editor'); ?></h3>
                        <p><?php _e('or', 'ai-excel-editor'); ?></p>
                        <button type="button" class="button button-primary" onclick="document.getElementById('file-upload').click()">
                            <?php _e('Choose File', 'ai-excel-editor'); ?>
                        </button>
                        <input type="file" id="file-upload" accept=".xlsx,.xls,.csv" class="file-input">
                    </div>
                </div>

                <div class="instructions-area">
                    <textarea 
                        id="ai-instructions" 
                        placeholder="<?php _e('Describe the modifications you want to make to your Excel file', 'ai-excel-editor'); ?>"
                        rows="4"
                    ></textarea>
                </div>
                
                <!-- Example instructions to help users -->
                <div class="instructions-helper">
                    <h4><?php _e('Example Instructions:', 'ai-excel-editor'); ?></h4>
                    <div class="example-instructions">
                        <button type="button" class="instruction-example"><?php _e('Calculate the sum of column B and add it to the bottom', 'ai-excel-editor'); ?></button>
                        <button type="button" class="instruction-example"><?php _e('Sort the data by the "Revenue" column from highest to lowest', 'ai-excel-editor'); ?></button>
                        <button type="button" class="instruction-example"><?php _e('Format all numbers in column C as currency with $ symbol', 'ai-excel-editor'); ?></button>
                    </div>
                </div>

                <div class="action-buttons">
                    <button type="button" id="process-btn" class="button button-primary">
                        <?php _e('Process & Download', 'ai-excel-editor'); ?>
                    </button>
                    <button type="button" id="reset-btn" class="button button-neutral">
                        <?php _e('Reset', 'ai-excel-editor'); ?>
                    </button>
                </div>
            </div>

            <div class="preview-section">
                <h3><?php _e('File Preview', 'ai-excel-editor'); ?></h3>
                <div id="file-preview"></div>
            </div>
            
            <?php if (current_user_can('manage_options')): ?>
            <div style="margin-top: 30px; padding: 10px; border-left: 4px solid #0073aa; background: #f8f8f8;">
                <h4>Debug Info (Admin Only)</h4>
                <p><strong>Upload directory:</strong> <?php echo esc_html($this->upload_dir); ?></p>
                <p><strong>Directory exists:</strong> <?php echo file_exists($this->upload_dir) ? 'Yes' : 'No'; ?></p>
                <p><strong>Directory permissions:</strong> <?php echo file_exists($this->upload_dir) ? substr(sprintf('%o', fileperms($this->upload_dir)), -4) : 'N/A'; ?></p>
                <p><strong>Directory writable:</strong> <?php echo file_exists($this->upload_dir) && is_writable($this->upload_dir) ? 'Yes' : 'No'; ?></p>
            </div>
            <?php endif; ?>
        </div>
        <?php
        $output = ob_get_clean();
        aisheets_debug('Editor shortcode rendered');
        return $output;
    }

    public function handle_excel_processing() {
        aisheets_debug('AJAX process_excel handler called');
        aisheets_debug('REQUEST data', $_REQUEST);
        aisheets_debug('FILES data', $_FILES);
        
        try {
            // Check nonce
            if (!check_ajax_referer('ai_excel_editor_nonce', 'nonce', false)) {
                aisheets_debug('Nonce verification failed');
                wp_send_json_error(array(
                    'message' => 'Security check failed',
                    'code' => 'nonce_failure',
                    'details' => 'The security token has expired or is invalid'
                ));
            }

            aisheets_debug('Nonce verification passed');
            
            // Check user permissions
            if (!current_user_can('upload_files')) {
                aisheets_debug('User permission denied');
                wp_send_json_error(array(
                    'message' => 'Permission denied',
                    'code' => 'insufficient_permissions',
                    'details' => 'User does not have permission to upload files'
                ));
            }

            // Check for file
            if (!isset($_FILES['file'])) {
                aisheets_debug('No file uploaded');
                wp_send_json_error(array(
                    'message' => 'No file uploaded',
                    'code' => 'no_file',
                    'details' => 'The file upload data was not found in the request'
                ));
            }

            $file = $_FILES['file'];
            $instructions = sanitize_textarea_field($_POST['instructions']);

            aisheets_debug('File upload details', array(
                'name' => $file['name'],
                'size' => $file['size'],
                'type' => $file['type'],
                'error' => $file['error'],
                'tmp_name' => $file['tmp_name'],
                'tmp_name_exists' => file_exists($file['tmp_name'])
            ));

            // Validate file with detailed errors
            $validation_result = $this->validate_file($file);
            if (is_wp_error($validation_result)) {
                aisheets_debug('File validation failed', array(
                    'error_code' => $validation_result->get_error_code(),
                    'error_message' => $validation_result->get_error_message(),
                    'error_data' => $validation_result->get_all_error_data()
                ));
                
                wp_send_json_error(array(
                    'message' => $validation_result->get_error_message(),
                    'code' => $validation_result->get_error_code(),
                    'details' => $validation_result->get_all_error_data()
                ));
            }

            // Create processing directory with error checking
            $processing_dir = $this->upload_dir . '/processing';
            if (!file_exists($processing_dir)) {
                $dir_created = wp_mkdir_p($processing_dir);
                aisheets_debug('Created processing directory', array(
                    'success' => $dir_created,
                    'path' => $processing_dir,
                    'error' => $dir_created ? null : error_get_last()
                ));
                
                if (!$dir_created) {
                    wp_send_json_error(array(
                        'message' => 'Failed to create processing directory',
                        'code' => 'directory_creation_failed',
                        'details' => error_get_last()
                    ));
                }
                
                // Set permissions
                chmod($processing_dir, 0755);
            }

            // Generate unique filename
            $new_filename = uniqid() . '_' . sanitize_file_name($file['name']);
            $upload_path = $processing_dir . '/' . $new_filename;

            aisheets_debug('Attempting to move uploaded file', array(
                'from' => $file['tmp_name'],
                'to' => $upload_path,
                'tmp_file_exists' => file_exists($file['tmp_name']),
                'processing_dir_writable' => is_writable($processing_dir)
            ));

            // Move uploaded file with error checking
            if (!move_uploaded_file($file['tmp_name'], $upload_path)) {
                $upload_error = error_get_last();
                aisheets_debug('Failed to move uploaded file', array(
                    'error' => $upload_error,
                    'file_perms' => substr(sprintf('%o', fileperms($processing_dir)), -4),
                    'target_path' => $upload_path
                ));
                
                wp_send_json_error(array(
                    'message' => 'Failed to move uploaded file',
                    'code' => 'move_upload_failed',
                    'details' => array(
                        'php_error' => $upload_error,
                        'file_perms' => substr(sprintf('%o', fileperms($processing_dir)), -4),
                        'target_path' => $upload_path
                    )
                ));
            }

            aisheets_debug('Successfully moved uploaded file to: ' . $upload_path);

            try {
                // Process with OpenAI
                aisheets_debug('Processing file with OpenAI');
                $processed_file = $this->process_with_openai($upload_path, $instructions);
                
                // Get file URL for download
                $upload_dir = wp_upload_dir();
                $file_url = $upload_dir['baseurl'] . '/ai-excel-editor/processing/' . basename($processed_file);

                aisheets_debug('File processed successfully', array(
                    'original' => $file['name'],
                    'processed' => basename($processed_file),
                    'download_url' => $file_url
                ));

                wp_send_json_success(array(
                    'file_url' => $file_url,
                    'preview' => sprintf(
                        '<div class="preview-content">
                            <p><strong>File processed:</strong> %s</p>
                            <p><strong>Instructions:</strong> %s</p>
                        </div>',
                        esc_html($file['name']),
                        esc_html($instructions)
                    )
                ));
            } catch (Exception $e) {
                aisheets_debug('OpenAI processing failed', array(
                    'error' => $e->getMessage(),
                    'trace' => $e->getTraceAsString()
                ));
                
                wp_send_json_error(array(
                    'message' => 'OpenAI processing failed',
                    'code' => 'openai_processing_error',
                    'details' => array(
                        'error_message' => $e->getMessage(),
                        'file_name' => $file['name'],
                        'instructions' => $instructions
                    )
                ));
            }
        } catch (Exception $e) {
            aisheets_debug('Unhandled exception in AJAX handler', array(
                'error' => $e->getMessage(),
                'trace' => $e->getTraceAsString()
            ));
            
            wp_send_json_error(array(
                'message' => 'Server processing error',
                'code' => 'server_error',
                'details' => array(
                    'error_message' => $e->getMessage(),
                    'error_trace' => $e->getTraceAsString()
                )
            ));
        }
    }

    private function validate_file($file) {
        aisheets_debug('Validating file', array(
            'name' => $file['name'],
            'size' => $file['size'],
            'type' => $file['type'],
            'error' => $file['error']
        ));
        
        // Check for upload errors
        if ($file['error'] !== UPLOAD_ERR_OK) {
            $error_message = $this->get_upload_error_message($file['error']);
            aisheets_debug('Upload error: ' . $error_message);
            
            return new WP_Error(
                'upload_error',
                $error_message,
                array('error_code' => $file['error'])
            );
        }

        // Check file size
        if ($file['size'] > $this->max_file_size) {
            aisheets_debug('File size exceeds limit', array(
                'size' => $file['size'],
                'limit' => $this->max_file_size
            ));
            
            return new WP_Error(
                'file_size',
                'File size exceeds the maximum limit of 5MB',
                array(
                    'file_size' => $file['size'],
                    'max_size' => $this->max_file_size
                )
            );
        }

        // Check file type
        $file_ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
        if (!in_array($file_ext, $this->allowed_file_types)) {
            aisheets_debug('Invalid file type', array(
                'type' => $file_ext,
                'allowed' => $this->allowed_file_types
            ));
            
            return new WP_Error(
                'file_type',
                'Invalid file type. Please upload an Excel or CSV file.',
                array(
                    'file_type' => $file_ext,
                    'allowed_types' => $this->allowed_file_types
                )
            );
        }

        aisheets_debug('File validation passed');
        return true;
    }

    private function get_upload_error_message($error_code) {
        switch ($error_code) {
            case UPLOAD_ERR_INI_SIZE:
                return 'The uploaded file exceeds the upload_max_filesize directive in php.ini';
            case UPLOAD_ERR_FORM_SIZE:
                return 'The uploaded file exceeds the MAX_FILE_SIZE directive that was specified in the HTML form';
            case UPLOAD_ERR_PARTIAL:
                return 'The uploaded file was only partially uploaded';
            case UPLOAD_ERR_NO_FILE:
                return 'No file was uploaded';
            case UPLOAD_ERR_NO_TMP_DIR:
                return 'Missing a temporary folder';
            case UPLOAD_ERR_CANT_WRITE:
                return 'Failed to write file to disk';
            case UPLOAD_ERR_EXTENSION:
                return 'A PHP extension stopped the file upload';
            default:
                return 'Unknown upload error';
        }
    }

    private function process_with_openai($file_path, $instructions) {
        aisheets_debug('Starting OpenAI processing', array(
            'file' => basename($file_path),
            'instructions_length' => strlen($instructions)
        ));
        
        // Check if OpenAI API key is configured
        if (empty($this->openai_api_key)) {
            aisheets_debug('OpenAI API key not configured');
            throw new Exception('OpenAI API key not configured');
        }
    
        aisheets_debug('Processing file with OpenAI', array(
            'file' => basename($file_path),
            'size' => filesize($file_path),
            'instructions_length' => strlen($instructions)
        ));
    
        try {
            // Check if vendor directory exists and PhpSpreadsheet is properly installed
            $vendor_path = AI_EXCEL_EDITOR_PLUGIN_DIR . 'vendor/autoload.php';
            if (!file_exists($vendor_path)) {
                aisheets_debug('Vendor autoload.php not found at: ' . $vendor_path);
                throw new Exception('Required dependency files not found. Please ensure PhpSpreadsheet is installed.');
            }
    
            // Include required libraries
            require_once $vendor_path;
            
            // Check if our class files exist
            $spreadsheet_class = AI_EXCEL_EDITOR_PLUGIN_DIR . 'includes/class-spreadsheet.php';
            $openai_class = AI_EXCEL_EDITOR_PLUGIN_DIR . 'includes/class-openai.php';
            
            aisheets_debug('Checking for required class files', array(
                'spreadsheet_exists' => file_exists($spreadsheet_class),
                'openai_exists' => file_exists($openai_class)
            ));
            
            if (!file_exists($spreadsheet_class) || !file_exists($openai_class)) {
                throw new Exception('Required class files not found. Please check your installation.');
            }
            
            require_once $spreadsheet_class;
            require_once $openai_class;
    
            // Simple test to verify PhpSpreadsheet is working
            try {
                $test_spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
                aisheets_debug('PhpSpreadsheet initialized successfully');
            } catch (Exception $e) {
                aisheets_debug('PhpSpreadsheet initialization failed: ' . $e->getMessage());
                throw new Exception('PhpSpreadsheet initialization failed: ' . $e->getMessage());
            }
    
            // TEMPORARY FIX: Just return the original file while debugging
            aisheets_debug('Using temporary pass-through solution for testing');
            
            // Generate output filename
            $output_dir = dirname($file_path);
            $output_filename = 'processed_' . uniqid() . '_' . basename($file_path);
            $output_path = $output_dir . '/' . $output_filename;
            
            // Just copy the file for now
            if (copy($file_path, $output_path)) {
                aisheets_debug('File copied successfully for temporary pass-through', array(
                    'from' => $file_path,
                    'to' => $output_path
                ));
                
                return $output_path;
            } else {
                aisheets_debug('Failed to copy file for temporary pass-through', error_get_last());
                throw new Exception('Failed to copy file: ' . print_r(error_get_last(), true));
            }
            
            /* COMMENTED OUT FULL IMPLEMENTATION FOR NOW
            // Initialize spreadsheet handler
            $spreadsheet_handler = new AISheets_Spreadsheet();
            aisheets_debug('Spreadsheet handler initialized');
            
            // Read spreadsheet data
            $spreadsheet_data = $spreadsheet_handler->read_file($file_path);
            aisheets_debug('Spreadsheet data extracted successfully');
            
            // Initialize OpenAI handler
            $openai_handler = new AISheets_OpenAI($this->openai_api_key);
            aisheets_debug('OpenAI handler initialized');
            
            // Process with OpenAI
            $changes = $openai_handler->process_spreadsheet($spreadsheet_data, $instructions);
            aisheets_debug('OpenAI processing completed with changes: ' . json_encode(array_keys($changes)));
            
            // Generate output filename
            $output_dir = dirname($file_path);
            $output_filename = 'modified_' . uniqid() . '_' . basename($file_path);
            $output_path = $output_dir . '/' . $output_filename;
            
            // Apply changes and save file
            $processed_file = $spreadsheet_handler->apply_changes($file_path, $changes, $output_path);
            
            aisheets_debug('OpenAI processing completed successfully: ' . $processed_file);
            return $processed_file;
            */
            
        } catch (Exception $e) {
            aisheets_debug('OpenAI Processing Error: ' . $e->getMessage());
            aisheets_debug('Error Stack Trace: ' . $e->getTraceAsString());
            throw $e;
        }
    }

    public function cleanup_old_files() {
        aisheets_debug('Running cleanup of old files');
        
        $processing_dir = $this->upload_dir . '/processing';
        $files = glob($processing_dir . '/*');
        $lifetime = 3600; // 1 hour
        $cleaned_count = 0;
        
        foreach ($files as $file) {
            if (is_file($file) && (time() - filemtime($file) > $lifetime)) {
                if (@unlink($file)) {
                    $cleaned_count++;
                }
            }
        }
        
        aisheets_debug('Cleanup completed', array(
            'files_removed' => $cleaned_count
        ));
    }

    public function handle_unauthorized() {
        aisheets_debug('Unauthorized access attempt');
        
        wp_send_json_error(array(
            'message' => 'You must be logged in to use this feature',
            'code' => 'unauthorized',
            'details' => 'This feature requires user authentication'
        ));
    }

    public function render_admin_page() {
        if (!current_user_can('manage_options')) {
            return;
        }
        include_once AI_EXCEL_EDITOR_PLUGIN_DIR . 'admin/partials/admin-display.php';
    }
}

// Initialize the plugin
$ai_excel_editor = new AI_Excel_Editor();

// Activation hook
register_activation_hook(__FILE__, function() {
    aisheets_debug('Plugin activation started');
    
    require_once AI_EXCEL_EDITOR_PLUGIN_DIR . 'includes/class-activator.php';
    AI_Excel_Editor_Activator::activate();
    
    aisheets_debug('Plugin activation completed');
});

// Deactivation hook
register_deactivation_hook(__FILE__, function() {
    aisheets_debug('Plugin deactivation started');
    
    require_once AI_EXCEL_EDITOR_PLUGIN_DIR . 'includes/class-activator.php';
    AI_Excel_Editor_Activator::deactivate();
    
    aisheets_debug('Plugin deactivation completed');
});

// Add diagnostic endpoint for admins only
add_action('wp_ajax_aisheets_diagnostics', function() {
    if (!current_user_can('manage_options')) {
        wp_send_json_error('Unauthorized');
        return;
    }
    
    aisheets_debug('Running diagnostics test');
    
    $diagnostics = array(
        'plugin_info' => array(
            'version' => AI_EXCEL_EDITOR_VERSION,
            'plugin_dir' => AI_EXCEL_EDITOR_PLUGIN_DIR,
            'plugin_url' => AI_EXCEL_EDITOR_PLUGIN_URL
        ),
        'php_info' => array(
            'version' => phpversion(),
            'max_upload_size' => ini_get('upload_max_filesize'),
            'post_max_size' => ini_get('post_max_size'),
            'max_execution_time' => ini_get('max_execution_time'),
            'memory_limit' => ini_get('memory_limit')
        ),
        'wordpress_info' => array(
            'version' => get_bloginfo('version'),
            'debug_mode' => WP_DEBUG ? 'Enabled' : 'Disabled'
        ),
        'directories' => array()
    );
    
    // Test upload directory
    $upload_dir = wp_upload_dir();
    $aisheets_dir = $upload_dir['basedir'] . '/ai-excel-editor';
    $processing_dir = $aisheets_dir . '/processing';
    
    $diagnostics['directories']['upload_dir'] = array(
        'path' => $upload_dir['basedir'],
        'exists' => file_exists($upload_dir['basedir']),
        'writable' => is_writable($upload_dir['basedir']),
        'permissions' => file_exists($upload_dir['basedir']) ? substr(sprintf('%o', fileperms($upload_dir['basedir'])), -4) : 'N/A'
    );
    
    $diagnostics['directories']['aisheets_dir'] = array(
        'path' => $aisheets_dir,
        'exists' => file_exists($aisheets_dir),
        'writable' => file_exists($aisheets_dir) && is_writable($aisheets_dir),
        'permissions' => file_exists($aisheets_dir) ? substr(sprintf('%o', fileperms($aisheets_dir)), -4) : 'N/A'
    );
    
    $diagnostics['directories']['processing_dir'] = array(
        'path' => $processing_dir,
        'exists' => file_exists($processing_dir),
        'writable' => file_exists($processing_dir) && is_writable($processing_dir),
        'permissions' => file_exists($processing_dir) ? substr(sprintf('%o', fileperms($processing_dir)), -4) : 'N/A'
    );
    
    // Test file writing
    if (file_exists($aisheets_dir) && is_writable($aisheets_dir)) {
        $test_file = $aisheets_dir . '/test_' . uniqid() . '.txt';
        $write_result = file_put_contents($test_file, 'Test file write for diagnostics');
        
        $diagnostics['file_write_test'] = array(
            'success' => $write_result !== false,
            'bytes_written' => $write_result,
            'file_path' => $test_file
        );
        
        if ($write_result !== false) {
            unlink($test_file); // Clean up
        }
    } else {
        $diagnostics['file_write_test'] = array(
            'success' => false,
            'reason' => 'Directory does not exist or is not writable'
        );
    }
    
    // Test PhpSpreadsheet
    $vendor_path = AI_EXCEL_EDITOR_PLUGIN_DIR . 'vendor/autoload.php';
    $diagnostics['phpspreadsheet'] = array(
        'autoload_exists' => file_exists($vendor_path)
    );
    
    if (file_exists($vendor_path)) {
        try {
            require_once $vendor_path;
            $phpspreadsheet_test = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
            $diagnostics['phpspreadsheet']['initialized'] = true;
        } catch (Exception $e) {
            $diagnostics['phpspreadsheet']['initialized'] = false;
            $diagnostics['phpspreadsheet']['error'] = $e->getMessage();
        }
    }
    
    aisheets_debug('Diagnostics completed', $diagnostics);
    wp_send_json_success($diagnostics);
});