<?php
/**
 * Plugin Activation and Setup
 */

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
        update_option('ai_excel_editor_version', '1.0.0');

        // Flush rewrite rules
        flush_rewrite_rules();
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

        // Check if errors exist
        if (!empty($errors)) {
            deactivate_plugins(plugin_basename(__FILE__));
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
            $base_dir . '/processed',
            $base_dir . '/original'
        );

        foreach ($directories as $dir) {
            if (!file_exists($dir)) {
                wp_mkdir_p($dir);
            }

            // Create .htaccess for security
            $htaccess = $dir . '/.htaccess';
            if (!file_exists($htaccess)) {
                file_put_contents($htaccess, 
                    "Order Deny,Allow\nDeny from all\n");
            }

            // Create index.php for security
            $index = $dir . '/index.php';
            if (!file_exists($index)) {
                file_put_contents($index, 
                    "<?php\n// Silence is golden.");
            }
        }
    }

    /**
     * Set up roles and capabilities
     */
    private static function setup_roles() {
        // Add capabilities to administrator
        $admin = get_role('administrator');
        $admin->add_cap('manage_excel_editor');
        $admin->add_cap('process_excel_files');
        $admin->add_cap('view_processing_history');
        
        // Create custom role for Excel Editor users
        add_role('excel_editor_user', 'Excel Editor User', array(
            'read' => true,
            'process_excel_files' => true,
            'view_processing_history' => true
        ));
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
    }

    /**
     * Plugin deactivation
     */
    public static function deactivate() {
        // Clear scheduled hooks
        wp_clear_scheduled_hook('ai_excel_editor_daily_cleanup');
        wp_clear_scheduled_hook('ai_excel_editor_monthly_quota_reset');
        
        // Flush rewrite rules
        flush_rewrite_rules();
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

        // Clean up upload directory
        $upload_dir = wp_upload_dir();
        $base_dir = $upload_dir['basedir'] . '/ai-excel-editor';
        self::recursive_rmdir($base_dir);
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
        }
    }
}
