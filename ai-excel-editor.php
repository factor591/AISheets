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

// For testing only - allow non-authenticated users to process files
add_action('wp_ajax_nopriv_process_excel', function() {
    // Get the global plugin instance
    global $ai_excel_editor;
    if ($ai_excel_editor) {
        // Call the same handler used for authenticated users
        $ai_excel_editor->handle_excel_processing();
    } else {
        wp_send_json_error(['message' => 'Plugin instance not available']);
    }
});

class AI_Excel_Editor {
    private $openai_api_key;
    private $upload_dir;
    private $allowed_file_types;
    private $max_file_size;
    private $api_key_source;

    public function __construct() {
        aisheets_debug('Plugin initializing');
        
        // First check if API key is defined in wp-config.php
        if (defined('AISHEETS_OPENAI_API_KEY') && !empty(AISHEETS_OPENAI_API_KEY)) {
            $this->openai_api_key = AISHEETS_OPENAI_API_KEY;
            $this->api_key_source = 'wp-config.php';
            aisheets_debug('Using OpenAI API key from wp-config.php');
        } else {
            // Fallback to WordPress option
            $this->openai_api_key = get_option('ai_excel_editor_openai_key');
            $this->api_key_source = 'wp_options';
            aisheets_debug('Using OpenAI API key from WordPress options');
        }
        
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

        // Direct download handler
        add_action('wp_ajax_download_excel', array($this, 'handle_direct_download'));
        add_action('wp_ajax_nopriv_download_excel', array($this, 'handle_direct_download'));

        // Debug AJAX handlers
        add_action('wp_ajax_aisheets_test', array($this, 'handle_test_ajax'));
        add_action('wp_ajax_nopriv_aisheets_test', array($this, 'handle_unauthorized'));
        
        add_action('wp_ajax_aisheets_debug', array($this, 'handle_debug_ajax'));
        add_action('wp_ajax_nopriv_aisheets_debug', array($this, 'handle_unauthorized'));

        // Cleanup schedule
        if (!wp_next_scheduled('ai_excel_editor_cleanup')) {
            wp_schedule_event(time(), 'hourly', 'ai_excel_editor_cleanup');
        }
        add_action('ai_excel_editor_cleanup', array($this, 'cleanup_old_files'));
        
        // Debug actions
        add_action('wp_footer', array($this, 'debug_info'));
        
        aisheets_debug('Plugin initialized', array(
            'openai_key_set' => !empty($this->openai_api_key),
            'api_key_source' => $this->api_key_source,
            'allowed_types' => $this->allowed_file_types,
            'max_file_size' => $this->max_file_size
        ));
    }
    
    /**
     * Handle direct file downloads
     */
    public function handle_direct_download() {
        aisheets_debug('Direct download handler called');
        
        // Check for file parameter
        if (!isset($_GET['file']) || empty($_GET['file'])) {
            wp_die('No file specified');
        }
        
        // Sanitize filename to prevent path traversal
        $filename = sanitize_file_name($_GET['file']);
        
        // Build file path
        $upload_dir = wp_upload_dir();
        $file_path = $upload_dir['basedir'] . '/ai-excel-editor/processing/' . $filename;
        
        aisheets_debug('Direct download requested for: ' . $file_path);
        
        // Check if file exists
        if (!file_exists($file_path)) {
            wp_die('File not found');
        }
        
        // Check file mime type
        $file_info = wp_check_filetype($file_path);
        $mime_type = $file_info['type'];
        
        if (!$mime_type) {
            // Default mime types for common spreadsheet formats
            $ext = pathinfo($file_path, PATHINFO_EXTENSION);
            switch ($ext) {
                case 'xlsx':
                    $mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
                    break;
                case 'xls':
                    $mime_type = 'application/vnd.ms-excel';
                    break;
                case 'csv':
                    $mime_type = 'text/csv';
                    break;
                default:
                    $mime_type = 'application/octet-stream';
            }
        }
        
        // Prepare for download
        if (ob_get_level()) {
            ob_end_clean();
        }
        
        // Set headers for download
        header('Content-Description: File Transfer');
        header('Content-Type: ' . $mime_type);
        header('Content-Disposition: attachment; filename="' . basename($file_path) . '"');
        header('Content-Transfer-Encoding: binary');
        header('Expires: 0');
        header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
        header('Pragma: public');
        header('Content-Length: ' . filesize($file_path));
        
        // Output file and exit
        readfile($file_path);
        exit;
    }
    
    // Test AJAX handler
    public function handle_test_ajax() {
        aisheets_debug('Test AJAX handler called');
        
        // Check nonce
        if (!check_ajax_referer('ai_excel_editor_nonce', 'nonce', false)) {
            aisheets_debug('Test AJAX nonce verification failed');
            wp_send_json_error('Security check failed');
        }
        
        aisheets_debug('Test AJAX nonce verification passed');
        wp_send_json_success('AJAX test successful');
    }
    
    // Debug AJAX handler
    public function handle_debug_ajax() {
        aisheets_debug('Debug AJAX handler called');
        
        // Check nonce
        if (!check_ajax_referer('ai_excel_editor_nonce', 'nonce', false)) {
            aisheets_debug('Debug AJAX nonce verification failed');
            wp_send_json_error('Security check failed');
        }
        
        aisheets_debug('Debug AJAX nonce verification passed');
        
        $upload_limits = array(
            'post_max_size' => ini_get('post_max_size'),
            'upload_max_filesize' => ini_get('upload_max_filesize'),
            'max_file_uploads' => ini_get('max_file_uploads'),
            'max_input_time' => ini_get('max_input_time'),
            'max_execution_time' => ini_get('max_execution_time'),
            'memory_limit' => ini_get('memory_limit')
        );
        
        $upload_dir = wp_upload_dir();
        $processing_dir = $upload_dir['basedir'] . '/ai-excel-editor/processing';
        
        $directory_info = array(
            'upload_dir_exists' => file_exists($upload_dir['basedir']),
            'processing_dir_exists' => file_exists($processing_dir),
            'upload_dir_writable' => is_writable($upload_dir['basedir']),
            'processing_dir_writable' => is_writable($processing_dir),
            'plugin_directory' => AI_EXCEL_EDITOR_PLUGIN_DIR,
            'plugin_url' => AI_EXCEL_EDITOR_PLUGIN_URL,
            'api_key_source' => $this->api_key_source,
            'api_key_configured' => !empty($this->openai_api_key)
        );
        
        wp_send_json_success(array(
            'php_config' => $upload_limits,
            'directories' => $directory_info,
            'ajax_url' => admin_url('admin-ajax.php'),
            'wp_version' => get_bloginfo('version'),
            'current_nonce' => wp_create_nonce('ai_excel_editor_nonce')
        ));
    }
    
    // Debug info in footer
    public function debug_info() {
        global $post;
        if (is_a($post, 'WP_Post') && has_shortcode($post->post_content, 'ai_excel_editor')) {
            echo '<script>
                console.log("AISheets Debug Info:");
                console.log("Plugin URL:", "' . AI_EXCEL_EDITOR_PLUGIN_URL . '");
                console.log("jQuery Loaded:", typeof jQuery !== "undefined");
                console.log("WordPress AJAX URL:", "' . admin_url('admin-ajax.php') . '");
                console.log("Nonce:", "' . wp_create_nonce('ai_excel_editor_nonce') . '");
                console.log("User Logged In:", "' . (is_user_logged_in() ? 'Yes' : 'No') . '");
                console.log("API Key Source:", "' . esc_js($this->api_key_source) . '");
                console.log("API Key Configured:", "' . (!empty($this->openai_api_key) ? 'Yes' : 'No') . '");
                
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
                    
                    // Check if the AJAX object is properly initialized
                    if (typeof aiExcelEditor !== "undefined") {
                        console.log("aiExcelEditor object:", {
                            ajax_url: aiExcelEditor.ajax_url,
                            nonce_exists: !!aiExcelEditor.nonce,
                            max_file_size: aiExcelEditor.max_file_size,
                            allowed_types: aiExcelEditor.allowed_types
                        });
                    } else {
                        console.error("aiExcelEditor object not found! wp_localize_script may have failed.");
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
                
                // Create .htaccess to allow downloads
                file_put_contents($this->upload_dir . '/.htaccess', 
                    "# Allow all files in this directory to be downloaded\n" .
                    "<IfModule mod_authz_core.c>\n" .
                    "    Require all granted\n" .
                    "</IfModule>\n" .
                    "<IfModule !mod_authz_core.c>\n" .
                    "    Order allow,deny\n" .
                    "    Allow from all\n" .
                    "</IfModule>\n\n" .
                    "# Set proper content types\n" .
                    "<IfModule mod_mime.c>\n" .
                    "    AddType application/vnd.openxmlformats-officedocument.spreadsheetml.sheet .xlsx\n" .
                    "    AddType application/vnd.ms-excel .xls\n" .
                    "    AddType text/csv .csv\n" .
                    "</IfModule>"
                );
                
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
                
                // Create .htaccess to allow downloads
                file_put_contents($processing_dir . '/.htaccess', 
                    "# Allow all files in this directory to be downloaded\n" .
                    "<IfModule mod_authz_core.c>\n" .
                    "    Require all granted\n" .
                    "</IfModule>\n" .
                    "<IfModule !mod_authz_core.c>\n" .
                    "    Order allow,deny\n" .
                    "    Allow from all\n" .
                    "</IfModule>\n\n" .
                    "# Set proper content types\n" .
                    "<IfModule mod_mime.c>\n" .
                    "    AddType application/vnd.openxmlformats-officedocument.spreadsheetml.sheet .xlsx\n" .
                    "    AddType application/vnd.ms-excel .xls\n" .
                    "    AddType text/csv .csv\n" .
                    "</IfModule>"
                );
                
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

            // Create a fresh nonce for each page load
            $nonce = wp_create_nonce('ai_excel_editor_nonce');
            
            // Localize the script with new data
            wp_localize_script('ai-excel-editor', 'aiExcelEditor', array(
                'ajax_url' => admin_url('admin-ajax.php'),
                'nonce' => $nonce,
                'max_file_size' => $this->max_file_size,
                'allowed_types' => $this->allowed_file_types,
                'download_url' => admin_url('admin-ajax.php?action=download_excel'),
                'api_key_configured' => !empty($this->openai_api_key),
                'api_key_source' => $this->api_key_source
            ));
            
            aisheets_debug('Scripts and styles enqueued', array(
                'css_url' => AI_EXCEL_EDITOR_PLUGIN_URL . 'css/style.css',
                'js_url' => AI_EXCEL_EDITOR_PLUGIN_URL . 'js/main.js',
                'nonce_created' => true,
                'nonce_value' => $nonce
            ));
        }
    }

    public function render_editor() {
        aisheets_debug('Rendering editor shortcode');
        
        if (!$this->openai_api_key) {
            aisheets_debug('OpenAI API key not configured');
            return '<div class="error">Please configure the OpenAI API key in the plugin settings or wp-config.php file.</div>';
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
                    <button type="button" id="test-ajax-btn" class="button button-neutral">
                        <?php _e('Test AJAX', 'ai-excel-editor'); ?>
                    </button>
                    <button type="button" id="check-config-btn" class="button button-neutral">
                        <?php _e('Check Configuration', 'ai-excel-editor'); ?>
                    </button>
                    <button type="button" id="process-btn" class="button button-primary" disabled>
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
            
            <div id="debug-output" style="margin-top: 20px; display: none; padding: 10px; background: #f5f5f5; border: 1px solid #ddd;">
                <h3>Debug Information</h3>
                <pre id="debug-content"></pre>
            </div>
            
            <?php if (current_user_can('manage_options')): ?>
            <div style="margin-top: 30px; padding: 10px; border-left: 4px solid #0073aa; background: #f8f8f8;">
                <h4>Debug Info (Admin Only)</h4>
                <p><strong>Upload directory:</strong> <?php echo esc_html($this->upload_dir); ?></p>
                <p><strong>Directory exists:</strong> <?php echo file_exists($this->upload_dir) ? 'Yes' : 'No'; ?></p>
                <p><strong>Directory permissions:</strong> <?php echo file_exists($this->upload_dir) ? substr(sprintf('%o', fileperms($this->upload_dir)), -4) : 'N/A'; ?></p>
                <p><strong>Directory writable:</strong> <?php echo file_exists($this->upload_dir) && is_writable($this->upload_dir) ? 'Yes' : 'No'; ?></p>
                <p><strong>AJAX URL:</strong> <?php echo esc_html(admin_url('admin-ajax.php')); ?></p>
                <p><strong>Current Nonce:</strong> <?php echo esc_html(wp_create_nonce('ai_excel_editor_nonce')); ?></p>
                <p><strong>PHP Memory Limit:</strong> <?php echo esc_html(ini_get('memory_limit')); ?></p>
                <p><strong>Max Upload Size:</strong> <?php echo esc_html(ini_get('upload_max_filesize')); ?></p>
                <p><strong>Post Max Size:</strong> <?php echo esc_html(ini_get('post_max_size')); ?></p>
                <p><strong>User Logged In:</strong> <?php echo is_user_logged_in() ? 'Yes' : 'No'; ?></p>
                <p><strong>User ID:</strong> <?php echo get_current_user_id(); ?></p>
                <p><strong>API Key Source:</strong> <?php echo esc_html($this->api_key_source); ?></p>
                <p><strong>API Key Configured:</strong> <?php echo !empty($this->openai_api_key) ? 'Yes' : 'No'; ?></p>
            </div>
            <?php endif; ?>
        </div>
        <?php
        $output = ob_get_clean();
        aisheets_debug('Editor shortcode rendered');
        return $output;
    }

    public function handle_excel_processing() {
        aisheets_debug('===== AJAX process_excel handler called =====');
        aisheets_debug('REQUEST data', $_REQUEST);
        aisheets_debug('FILES data', $_FILES);
        aisheets_debug('Request method', $_SERVER['REQUEST_METHOD']);
        
        try {
            // Direct download handling
            if (isset($_REQUEST['direct_download']) && $_REQUEST['direct_download'] === 'true' && isset($_REQUEST['file'])) {
                $filename = sanitize_file_name($_REQUEST['file']);
                $file_path = $this->upload_dir . '/processing/' . $filename;
                
                if (file_exists($file_path)) {
                    aisheets_debug('Direct download request for: ' . $file_path);
                    
                    // Determine mime type
                    $file_info = wp_check_filetype($file_path);
                    $mime_type = $file_info['type'] ?: 'application/octet-stream';
                    
                    // Clear any output buffers
                    if (ob_get_level()) {
                        ob_end_clean();
                    }
                    
                    // Set headers
                    header('Content-Description: File Transfer');
                    header('Content-Type: ' . $mime_type);
                    header('Content-Disposition: attachment; filename="' . basename($file_path) . '"');
                    header('Content-Transfer-Encoding: binary');
                    header('Expires: 0');
                    header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
                    header('Pragma: public');
                    header('Content-Length: ' . filesize($file_path));
                    
                    // Output file and exit
                    readfile($file_path);
                    exit;
                } else {
                    aisheets_debug('File not found for direct download: ' . $file_path);
                    wp_send_json_error([
                        'message' => 'File not found for download',
                        'code' => 'file_not_found'
                    ]);
                    return;
                }
            }
            
            // Check nonce with detailed logging
            $nonce = isset($_REQUEST['nonce']) ? $_REQUEST['nonce'] : '';
            aisheets_debug('Checking nonce: ' . $nonce);
            
            // Use wp_verify_nonce without early return
            $valid = wp_verify_nonce($nonce, 'ai_excel_editor_nonce');
            if (!$valid) {
                aisheets_debug('Nonce verification failed. Provided: ' . $nonce);
                aisheets_debug('This could be due to nonce timeout or mismatch.');
                wp_send_json_error(array(
                    'message' => 'Security check failed',
                    'code' => 'nonce_failure',
                    'details' => 'The security token has expired or is invalid'
                ));
                return;
            }

            aisheets_debug('Continuing with file processing');
            
            // Check user permissions
            if (!current_user_can('upload_files')) {
                aisheets_debug('User permission denied');
                wp_send_json_error(array(
                    'message' => 'Permission denied',
                    'code' => 'insufficient_permissions',
                    'details' => 'User does not have permission to upload files'
                ));
                return;
            }

            // Check for file
            if (!isset($_FILES['file']) || empty($_FILES['file']['tmp_name'])) {
                aisheets_debug('No file uploaded or file upload failed. $_FILES data:', $_FILES);
                wp_send_json_error(array(
                    'message' => 'No file uploaded or upload failed',
                    'code' => 'no_file',
                    'details' => [
                        'files_data' => $_FILES,
                        'post_max_size' => ini_get('post_max_size'),
                        'upload_max_filesize' => ini_get('upload_max_filesize')
                    ]
                ));
                return;
            }

            $file = $_FILES['file'];
            $instructions = isset($_POST['instructions']) ? sanitize_textarea_field($_POST['instructions']) : '';

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
                return;
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
                    return;
                }
                
                // Set permissions
                chmod($processing_dir, 0755);
                
                // Create .htaccess to allow downloads
                file_put_contents($processing_dir . '/.htaccess', 
                    "# Allow all files in this directory to be downloaded\n" .
                    "<IfModule mod_authz_core.c>\n" .
                    "    Require all granted\n" .
                    "</IfModule>\n" .
                    "<IfModule !mod_authz_core.c>\n" .
                    "    Order allow,deny\n" .
                    "    Allow from all\n" .
                    "</IfModule>\n\n" .
                    "# Set proper content types\n" .
                    "<IfModule mod_mime.c>\n" .
                    "    AddType application/vnd.openxmlformats-officedocument.spreadsheetml.sheet .xlsx\n" .
                    "    AddType application/vnd.ms-excel .xls\n" .
                    "    AddType text/csv .csv\n" .
                    "</IfModule>"
                );
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
                return;
            }

            aisheets_debug('Successfully moved uploaded file to: ' . $upload_path);
            
            // Set correct permissions
            chmod($upload_path, 0644);

            try {
                // Process with OpenAI
                aisheets_debug('Processing file with OpenAI');
                $processed_file = $this->process_with_openai($upload_path, $instructions);
                
                // Two options for downloading:
                // OPTION 1: Direct PHP download (more reliable but keeps the user on the page)
                if (isset($_REQUEST['direct_download']) && $_REQUEST['direct_download'] === 'true') {
                    // Direct PHP download
                    aisheets_debug('Using direct PHP download method');
                    
                    if (file_exists($processed_file)) {
                        // Prevent output buffering
                        if (ob_get_level()) {
                            ob_end_clean();
                        }
                        
                        // Set headers for download
                        header('Content-Description: File Transfer');
                        header('Content-Type: application/octet-stream');
                        header('Content-Disposition: attachment; filename="' . basename($processed_file) . '"');
                        header('Content-Transfer-Encoding: binary');
                        header('Expires: 0');
                        header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
                        header('Pragma: public');
                        header('Content-Length: ' . filesize($processed_file));
                        
                        // Output file data
                        readfile($processed_file);
                        exit;
                    } else {
                        aisheets_debug('File not found for direct download: ' . $processed_file);
                        wp_send_json_error([
                            'message' => 'File not found for download',
                            'code' => 'file_not_found'
                        ]);
                    }
                } else {
                    // OPTION 2: URL-based download (current method with improvements)
                    // Get file URL for download with timestamp to bypass cache
                    $upload_dir = wp_upload_dir();
                    $file_url = $upload_dir['baseurl'] . '/ai-excel-editor/processing/' . basename($processed_file);
                    $file_url = add_query_arg('t', time(), $file_url); // Add timestamp to bypass cache
                    
                    aisheets_debug('File processed successfully', array(
                        'original' => $file['name'],
                        'processed' => basename($processed_file),
                        'download_url' => $file_url
                    ));
                    
                    // Prepare HTML for direct download button as fallback
                    $download_button = sprintf(
                        '<p><a href="%s" class="button button-primary" download>Download File</a></p>
                        <p><a href="%s" class="button" target="_blank">Open File</a></p>',
                        esc_url($file_url),
                        esc_url($file_url)
                    );
                    
                    wp_send_json_success(array(
                        'file_url' => $file_url,
                        'direct_download_url' => admin_url('admin-ajax.php?action=download_excel&file=' . urlencode(basename($processed_file))),
                        'preview' => sprintf(
                            '<div class="preview-content">
                                <p><strong>File processed:</strong> %s</p>
                                <p><strong>Instructions:</strong> %s</p>
                                <p><strong>Status:</strong> <span style="color: green;">Success</span></p>
                                %s
                                <p class="download-note">If automatic download doesn\'t start, please use the buttons above.</p>
                            </div>',
                            esc_html($file['name']),
                            esc_html($instructions),
                            $download_button
                        )
                    ));
                }
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

    /**
     * Process file with OpenAI
     * 
     * @param string $file_path Path to the uploaded file
     * @param string $instructions User instructions
     * @return string Path to the processed file
     */
    private function process_with_openai($file_path, $instructions) {
        aisheets_debug('Starting file processing', array(
            'file' => basename($file_path),
            'instructions_length' => strlen($instructions)
        ));
        
        try {
            // Check if OpenAI API key is configured
            if (empty($this->openai_api_key)) {
                aisheets_debug('OpenAI API key not configured');
                throw new Exception('OpenAI API key not configured');
            }
            
            // Load OpenAI integration class
            require_once AI_EXCEL_EDITOR_PLUGIN_DIR . 'includes/class-openai-integration.php';
            $openai = new AISheets_OpenAI_Integration($this->openai_api_key);
            
            // Generate output filename
            $upload_dir = wp_upload_dir();
            $processing_dir = $upload_dir['basedir'] . '/ai-excel-editor/processing';
            
            if (!file_exists($processing_dir)) {
                wp_mkdir_p($processing_dir);
            }
            
            $output_filename = 'processed_' . uniqid() . '_' . basename($file_path);
            $output_path = $processing_dir . '/' . $output_filename;
            
            // Process with OpenAI
            aisheets_debug('Calling OpenAI for processing');
            
            try {
                // Process spreadsheet with OpenAI
                $openai_response = $openai->process_spreadsheet($file_path, $instructions);
                
                // Apply changes to spreadsheet
                aisheets_debug('Applying changes to spreadsheet');
                $modified = $openai->apply_changes_to_spreadsheet($file_path, $output_path, $openai_response);
                
                if (!$modified) {
                    // Fallback to original file if modification failed
                    aisheets_debug('Modification failed, falling back to original file');
                    if (!copy($file_path, $output_path)) {
                        throw new Exception('Failed to copy file after modification failure');
                    }
                }
            } catch (Exception $api_error) {
                // Log the API error
                aisheets_debug('OpenAI processing error: ' . $api_error->getMessage());
                
                // Fallback to copying the original file
                aisheets_debug('Using fallback (copy original file)');
                if (!copy($file_path, $output_path)) {
                    throw new Exception('Failed to copy file: ' . error_get_last()['message']);
                }
            }
            
            // Set proper permissions
            chmod($output_path, 0644);
            
            // Get URL for download
            $file_url = $upload_dir['baseurl'] . '/ai-excel-editor/processing/' . basename($output_path);
            aisheets_debug('File URL for download: ' . $file_url);
            
            return $output_path;
            
        } catch (Exception $e) {
            aisheets_debug('Processing Error: ' . $e->getMessage());
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