<?php
/**
 * Plugin Activation and Setup
 *
 * Handles activation, deactivation, and uninstallation processes for the AI Excel Editor plugin.
 */

// If this file is called directly, abort.
if (!defined('ABSPATH')) {
    exit;
}

class AI_Excel_Editor_Activator {
    // Minimum requirements
    private static $min_php_version = '7.4';
    private static $min_wp_version = '5.6';
    private static $required_extensions = array(
        'curl',
        'json',
        'fileinfo',
        'zip'
    );

    /**
     * Main activation function
     */
    public static function activate() {
        // Check system requirements
        self::check_requirements();

        // Create necessary database tables
        self::create_tables();

        // Set up default options
        self::setup_options();

        // Create required directories
        self::create_directories();

        // Set up roles and capabilities
        self::setup_roles();

        // Schedule cron jobs
        self::setup_cron_jobs();

        // Set activation flag
        update_option('ai_excel_editor_activated', true);
        update_option('ai_excel_editor_version', AI_EXCEL_EDITOR_VERSION);

        // Flush rewrite rules
        flush_rewrite_rules();
        
        // Log activation
        error_log('AI Excel Editor activated: Version ' . AI_EXCEL_EDITOR_VERSION);
    }

    /**
     * Check system requirements
     */
    private static function check_requirements() {
        global $wp_version;
        $errors = array();

        // Check PHP version
        if (version_compare(PHP_VERSION, self::$min_php_version, '<')) {
            $errors[] = sprintf('PHP version %s or higher is required. Your version: %s', 
                self::$min_php_version, PHP_VERSION);
        }

        // Check WordPress version
        if (version_compare($wp_version, self::$min_wp_version, '<')) {
            $errors[] = sprintf('WordPress version %s or higher is required. Your version: %s', 
                self::$min_wp_version, $wp_version);
        }

        // Check required PHP extensions
        foreach (self::$required_extensions as $ext) {
            if (!extension_loaded($ext)) {
                $errors[] = sprintf('PHP extension %s is required but missing.', $ext);
            }
        }
        
        // Check if PhpSpreadsheet can be loaded
        if (file_exists(AI_EXCEL_EDITOR_PLUGIN_DIR . 'vendor/autoload.php')) {
            try {
                require_once AI_EXCEL_EDITOR_PLUGIN_DIR . 'vendor/autoload.php';
                if (!class_exists('PhpOffice\\PhpSpreadsheet\\Spreadsheet')) {
                    $errors[] = 'PhpSpreadsheet library not found. Make sure composer dependencies are installed.';
                }
            } catch (Exception $e) {
                $errors[] = 'Error loading dependencies: ' . $e->getMessage();
            }
        } else {
            $errors[] = 'Vendor autoload.php not found. Please run composer install or upload the vendor directory.';
        }

        // Check if errors exist
        if (!empty($errors)) {
            deactivate_plugins(plugin_basename(AI_EXCEL_EDITOR_PLUGIN_DIR . 'ai-excel-editor.php'));
            wp_die(implode('<br>', $errors), 'Plugin Activation Error', array(
                'back_link' => true
            ));
        }
    }

    /**
     * Create necessary database tables
     */
    private static function create_tables() {
        global $wpdb;
        require_once(ABSPATH . 'wp-admin/includes/upgrade.php');

        $charset_collate = $wpdb->get_charset_collate();

        // Table for processing history
        $table_name = $wpdb->prefix . 'ai_excel_processing_history';
        $sql = "CREATE TABLE IF NOT EXISTS $table_name (
            id bigint(20) NOT NULL AUTO_INCREMENT,
            user_id bigint(20) NOT NULL,
            file_name varchar(255) NOT NULL,
            original_file varchar(255) NOT NULL,
            processed_file varchar(255) NOT NULL,
            instructions text NOT NULL,
            status varchar(50) NOT NULL,
            created_at datetime NOT NULL,
            completed_at datetime DEFAULT NULL,
            PRIMARY KEY  (id),
            KEY user_id (user_id),
            KEY status (status)
        ) $charset_collate;";
        
        dbDelta($sql);

        // Table for user quotas
        $table_name = $wpdb->prefix . 'ai_excel_user_quotas';
        $sql = "CREATE TABLE IF NOT EXISTS $table_name (
            user_id bigint(20) NOT NULL,
            monthly_quota int(11) NOT NULL DEFAULT 0,
            used_quota int(11) NOT NULL DEFAULT 0,
            last_reset datetime DEFAULT NULL,
            PRIMARY KEY  (user_id)
        ) $charset_collate;";
        
        dbDelta($sql);
        
        // Table for processing metrics (optional)
        $table_name = $wpdb->prefix . 'ai_excel_processing_metrics';
        $sql = "CREATE TABLE IF NOT EXISTS $table_name (
            id bigint(20) NOT NULL AUTO_INCREMENT,
            processed_at datetime NOT NULL,
            user_id bigint(20) NOT NULL,
            metrics text NOT NULL,
            PRIMARY KEY  (id),
            KEY user_id (user_id)
        ) $charset_collate;";
        
        dbDelta($sql);
        
        // Log table creation
        error_log('AI Excel Editor tables created or updated');
    }

    /**
     * Set up default options
     */
    private static function setup_options() {
        // General settings
        add_option('ai_excel_editor_openai_key', '');
        add_option('ai_excel_editor_max_file_size', 5242880); // 5MB
        add_option('ai_excel_editor_allowed_file_types', array('xlsx', 'xls', 'csv'));
        
        // Quota settings
        add_option('ai_excel_editor_free_quota', 3);
        add_option('ai_excel_editor_pro_quota', 100);
        
        // Email settings
        add_option('ai_excel_editor_notification_email', get_option('admin_email'));
        add_option('ai_excel_editor_email_templates', array(
            'processing_complete' => "Your file {filename} has been processed successfully.",
            'processing_failed' => "There was an error processing your file {filename}."
        ));
        
        // Security settings
        add_option('ai_excel_editor_max_attempts', 3);
        add_option('ai_excel_editor_lockout_duration', 1800); // 30 minutes
        
        // Log options creation
        error_log('AI Excel Editor default options set up');
    }

    /**
     * Create required directories
     */
    private static function create_directories() {
        // Get WordPress upload directory
        $upload_dir = wp_upload_dir();
        $base_dir = $upload_dir['basedir'] . '/ai-excel-editor';
        
        // Create directories
        $directories = array(
            $base_dir,
            $base_dir . '/temp',
            $base_dir . '/processing',
            $base_dir . '/original'
        );

        $created_dirs = [];
        foreach ($directories as $dir) {
            if (!file_exists($dir)) {
                if (wp_mkdir_p($dir)) {
                    $created_dirs[] = $dir;
                }
            }

            // Create or update .htaccess to deny direct access
            $htaccess = $dir . '/.htaccess';
            file_put_contents(
                $htaccess,
                "<IfModule mod_authz_core.c>\n" .
                "    Require all denied\n" .
                "</IfModule>\n" .
                "<IfModule !mod_authz_core.c>\n" .
                "    Order deny,allow\n" .
                "    Deny from all\n" .
                "</IfModule>\n"
            );

            // Create index.php for security
            $index = $dir . '/index.php';
            if (!file_exists($index)) {
                file_put_contents($index, 
                    "<?php\n// Silence is golden.");
            }
        }
        
        // Log directory creation
        if (!empty($created_dirs)) {
            error_log('AI Excel Editor directories created: ' . implode(', ', $created_dirs));
        }
    }

    /**
     * Set up roles and capabilities
     */
    private static function setup_roles() {
        // Add capabilities to administrator
        $admin = get_role('administrator');
        if ($admin) {
            $admin->add_cap('manage_excel_editor');
            $admin->add_cap('process_excel_files');
            $admin->add_cap('view_processing_history');
        }
        
        // Create custom role for Excel Editor users
        remove_role('excel_editor_user'); // Remove first to update if exists
        add_role('excel_editor_user', 'Excel Editor User', array(
            'read' => true,
            'process_excel_files' => true,
            'view_processing_history' => true
        ));
        
        // Log roles setup
        error_log('AI Excel Editor roles and capabilities set up');
    }

    /**
     * Set up cron jobs
     */
    private static function setup_cron_jobs() {
        // Schedule daily cleanup
        if (!wp_next_scheduled('ai_excel_editor_daily_cleanup')) {
            wp_schedule_event(time(), 'daily', 'ai_excel_editor_daily_cleanup');
        }

        // Schedule monthly quota reset
        if (!wp_next_scheduled('ai_excel_editor_monthly_quota_reset')) {
            wp_schedule_event(time(), 'monthly', 'ai_excel_editor_monthly_quota_reset');
        }
        
        // Schedule hourly file cleanup
        if (!wp_next_scheduled('ai_excel_editor_cleanup')) {
            wp_schedule_event(time(), 'hourly', 'ai_excel_editor_cleanup');
        }
        
        // Log cron setup
        error_log('AI Excel Editor cron jobs scheduled');
    }

    /**
     * Plugin deactivation
     */
    public static function deactivate() {
        // Clear scheduled hooks
        wp_clear_scheduled_hook('ai_excel_editor_daily_cleanup');
        wp_clear_scheduled_hook('ai_excel_editor_monthly_quota_reset');
        wp_clear_scheduled_hook('ai_excel_editor_cleanup');
        
        // Flush rewrite rules
        flush_rewrite_rules();
        
        // Log deactivation
        error_log('AI Excel Editor deactivated');
    }

    /**
     * Plugin uninstall
     */
    public static function uninstall() {
        global $wpdb;

        // Only run if explicitly uninstalling
        if (!defined('WP_UNINSTALL_PLUGIN')) {
            return;
        }

        // Remove database tables
        $wpdb->query("DROP TABLE IF EXISTS {$wpdb->prefix}ai_excel_processing_history");
        $wpdb->query("DROP TABLE IF EXISTS {$wpdb->prefix}ai_excel_user_quotas");
        $wpdb->query("DROP TABLE IF EXISTS {$wpdb->prefix}ai_excel_processing_metrics");

        // Remove options
        delete_option('ai_excel_editor_activated');
        delete_option('ai_excel_editor_version');
        delete_option('ai_excel_editor_openai_key');
        delete_option('ai_excel_editor_max_file_size');
        delete_option('ai_excel_editor_allowed_file_types');
        delete_option('ai_excel_editor_free_quota');
        delete_option('ai_excel_editor_pro_quota');
        delete_option('ai_excel_editor_notification_email');
        delete_option('ai_excel_editor_email_templates');
        delete_option('ai_excel_editor_max_attempts');
        delete_option('ai_excel_editor_lockout_duration');

        // Remove custom roles
        remove_role('excel_editor_user');
        
        // Remove capabilities from admin
        $admin = get_role('administrator');
        if ($admin) {
            $admin->remove_cap('manage_excel_editor');
            $admin->remove_cap('process_excel_files');
            $admin->remove_cap('view_processing_history');
        }

        // Clean up upload directory
        $upload_dir = wp_upload_dir();
        $base_dir = $upload_dir['basedir'] . '/ai-excel-editor';
        self::recursive_rmdir($base_dir);
        
        // Log uninstall
        error_log('AI Excel Editor uninstalled - all data removed');
    }

    /**
     * Helper function to recursively remove directories
     */
    private static function recursive_rmdir($dir) {
        if (is_dir($dir)) {
            $objects = scandir($dir);
            foreach ($objects as $object) {
                if ($object != "." && $object != "..") {
                    if (is_dir($dir . "/" . $object)) {
                        self::recursive_rmdir($dir . "/" . $object);
                    } else {
                        unlink($dir . "/" . $object);
                    }
                }
            }
            rmdir($dir);
            return true;
        }
        return false;
    }
}