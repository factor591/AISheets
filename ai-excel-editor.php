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

class AI_Excel_Editor {
    private $openai_api_key;
    private $upload_dir;
    private $allowed_file_types;
    private $max_file_size;

    public function __construct() {
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
    }
    
    // Debug function - can be removed in production
    public function debug_info() {
        global $post;
        if (is_a($post, 'WP_Post') && has_shortcode($post->post_content, 'ai_excel_editor')) {
            echo '<script>
                console.log("AISheets Debug Info:");
                console.log("Plugin URL:", "' . AI_EXCEL_EDITOR_PLUGIN_URL . '");
                console.log("jQuery Loaded:", typeof jQuery !== "undefined");
            </script>';
        }
    }

    public function init() {
        try {
            // Create upload directory
            $upload_dir = wp_upload_dir();
            $this->upload_dir = $upload_dir['basedir'] . '/ai-excel-editor';
            
            if (!file_exists($this->upload_dir)) {
                if (!wp_mkdir_p($this->upload_dir)) {
                    throw new Exception('Failed to create upload directory');
                }
                file_put_contents($this->upload_dir . '/.htaccess', 'deny from all');
                file_put_contents($this->upload_dir . '/index.php', '<?php // Silence is golden');
            }

            // Create processing directory
            $processing_dir = $this->upload_dir . '/processing';
            if (!file_exists($processing_dir)) {
                if (!wp_mkdir_p($processing_dir)) {
                    throw new Exception('Failed to create processing directory');
                }
                file_put_contents($processing_dir . '/.htaccess', 'deny from all');
                file_put_contents($processing_dir . '/index.php', '<?php // Silence is golden');
            }
            
            // Log upload directories for debugging
            error_log('AISheets upload directory: ' . $this->upload_dir);
            error_log('AISheets processing directory: ' . $processing_dir);
            // Check directory permissions
            error_log('Upload directory permissions: ' . substr(sprintf('%o', fileperms($this->upload_dir)), -4));
            error_log('Processing directory permissions: ' . substr(sprintf('%o', fileperms($processing_dir)), -4));
            
        } catch (Exception $e) {
            error_log('AI Excel Editor Init Error: ' . $e->getMessage());
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
            
            // Log script loading
            error_log('AISheets scripts enqueued for post ID: ' . $post->ID);
        }
    }

    public function render_editor() {
        if (!$this->openai_api_key) {
            return '<div class="error">Please configure the OpenAI API key in the plugin settings.</div>';
        }

        ob_start();
        ?>
        <div class="ai-excel-editor-container">
            <div class="messages"></div>
            
            <div class="upload-section">
                <div id="excel-dropzone" class="dropzone">
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
        </div>
        <?php
        return ob_get_clean();
    }

    public function handle_excel_processing() {
        try {
            // Check nonce
            if (!check_ajax_referer('ai_excel_editor_nonce', 'nonce', false)) {
                error_log('AISheets: Nonce verification failed');
                wp_send_json_error(array(
                    'message' => 'Security check failed',
                    'code' => 'nonce_failure',
                    'details' => 'The security token has expired or is invalid'
                ));
            }

            // Log the request data for debugging
            error_log('AISheets AJAX request received');
            
            // Check user permissions
            if (!current_user_can('upload_files')) {
                wp_send_json_error(array(
                    'message' => 'Permission denied',
                    'code' => 'insufficient_permissions',
                    'details' => 'User does not have permission to upload files'
                ));
            }

            // Check for file
            if (!isset($_FILES['file'])) {
                error_log('AISheets: No file uploaded');
                wp_send_json_error(array(
                    'message' => 'No file uploaded',
                    'code' => 'no_file',
                    'details' => 'The file upload data was not found in the request'
                ));
            }

            $file = $_FILES['file'];
            $instructions = sanitize_textarea_field($_POST['instructions']);

            // Log uploaded file details
            error_log(sprintf(
                'AISheets file upload: Name=%s, Size=%d, Type=%s, Error=%d',
                $file['name'],
                $file['size'],
                $file['type'],
                $file['error']
            ));

            // Validate file with detailed errors
            $validation_result = $this->validate_file($file);
            if (is_wp_error($validation_result)) {
                wp_send_json_error(array(
                    'message' => $validation_result->get_error_message(),
                    'code' => $validation_result->get_error_code(),
                    'details' => $validation_result->get_all_error_data()
                ));
            }

            // Create processing directory with error checking
            $processing_dir = $this->upload_dir . '/processing';
            if (!file_exists($processing_dir) && !wp_mkdir_p($processing_dir)) {
                wp_send_json_error(array(
                    'message' => 'Failed to create processing directory',
                    'code' => 'directory_creation_failed',
                    'details' => error_get_last()
                ));
            }

            // Generate unique filename
            $new_filename = uniqid() . '_' . sanitize_file_name($file['name']);
            $upload_path = $processing_dir . '/' . $new_filename;

            // Move uploaded file with error checking
            if (!move_uploaded_file($file['tmp_name'], $upload_path)) {
                $upload_error = error_get_last();
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

            try {
                // Process with OpenAI
                $processed_file = $this->process_with_openai($upload_path, $instructions);
                
                // Get file URL for download
                $upload_dir = wp_upload_dir();
                $file_url = $upload_dir['baseurl'] . '/ai-excel-editor/processing/' . basename($processed_file);

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
            error_log('AI Excel Editor Error: ' . $e->getMessage());
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
        // Check for upload errors
        if ($file['error'] !== UPLOAD_ERR_OK) {
            return new WP_Error(
                'upload_error',
                $this->get_upload_error_message($file['error']),
                array('error_code' => $file['error'])
            );
        }

        // Check file size
        if ($file['size'] > $this->max_file_size) {
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
            return new WP_Error(
                'file_type',
                'Invalid file type. Please upload an Excel or CSV file.',
                array(
                    'file_type' => $file_ext,
                    'allowed_types' => $this->allowed_file_types
                )
            );
        }

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
        // Check if OpenAI API key is configured
        if (empty($this->openai_api_key)) {
            error_log('OpenAI Error: API key not configured');
            throw new Exception('OpenAI API key not configured');
        }
    
        error_log(sprintf(
            'OpenAI Processing Details: File: %s, Size: %s, Instructions: %s',
            basename($file_path),
            size_format(filesize($file_path)),
            $instructions
        ));
    
        try {
            // Check if vendor directory exists and PhpSpreadsheet is properly installed
            $vendor_path = AI_EXCEL_EDITOR_PLUGIN_DIR . 'vendor/autoload.php';
            if (!file_exists($vendor_path)) {
                error_log('Vendor autoload.php not found at: ' . $vendor_path);
                throw new Exception('Required dependency files not found. Please ensure PhpSpreadsheet is installed.');
            }
    
            // Include required libraries
            require_once $vendor_path;
            require_once AI_EXCEL_EDITOR_PLUGIN_DIR . 'includes/class-spreadsheet.php';
            require_once AI_EXCEL_EDITOR_PLUGIN_DIR . 'includes/class-openai.php';
    
            // Simple test to verify PhpSpreadsheet is working
            try {
                $test_spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
                error_log('PhpSpreadsheet initialized successfully');
            } catch (Exception $e) {
                error_log('PhpSpreadsheet initialization failed: ' . $e->getMessage());
                throw new Exception('PhpSpreadsheet initialization failed: ' . $e->getMessage());
            }
    
            // Initialize spreadsheet handler
            $spreadsheet_handler = new AISheets_Spreadsheet();
            error_log('Spreadsheet handler initialized');
            
            // Read spreadsheet data
            $spreadsheet_data = $spreadsheet_handler->read_file($file_path);
            error_log('Spreadsheet data extracted successfully');
            
            // Initialize OpenAI handler
            $openai_handler = new AISheets_OpenAI($this->openai_api_key);
            error_log('OpenAI handler initialized');
            
            // Process with OpenAI
            $changes = $openai_handler->process_spreadsheet($spreadsheet_data, $instructions);
            error_log('OpenAI processing completed with changes: ' . json_encode(array_keys($changes)));
            
            // Generate output filename
            $output_dir = dirname($file_path);
            $output_filename = 'modified_' . uniqid() . '_' . basename($file_path);
            $output_path = $output_dir . '/' . $output_filename;
            
            // Apply changes and save file
            $processed_file = $spreadsheet_handler->apply_changes($file_path, $changes, $output_path);
            
            error_log('OpenAI processing completed successfully: ' . $processed_file);
            return $processed_file;
        } catch (Exception $e) {
            error_log('OpenAI Processing Error: ' . $e->getMessage());
            error_log('Error Stack Trace: ' . $e->getTraceAsString());
            throw $e;
        }
    }

    public function cleanup_old_files() {
        $processing_dir = $this->upload_dir . '/processing';
        $files = glob($processing_dir . '/*');
        $lifetime = 3600; // 1 hour
        
        foreach ($files as $file) {
            if (is_file($file) && (time() - filemtime($file) > $lifetime)) {
                @unlink($file);
            }
        }
    }

    public function handle_unauthorized() {
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
    require_once AI_EXCEL_EDITOR_PLUGIN_DIR . 'includes/class-activator.php';
    AI_Excel_Editor_Activator::activate();
});

// Deactivation hook
register_deactivation_hook(__FILE__, function() {
    require_once AI_EXCEL_EDITOR_PLUGIN_DIR . 'includes/class-activator.php';
    AI_Excel_Editor_Activator::deactivate();
});