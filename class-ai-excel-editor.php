<?php
class AI_Excel_Editor {
    private $openai_api_key;
    private $plugin_name;
    private $version;

    public function __construct() {
        $this->plugin_name = 'ai-excel-editor';
        $this->version = AI_EXCEL_EDITOR_VERSION;
        $this->openai_api_key = get_option('ai_excel_editor_openai_key');
    }

    public function run() {
        $this->set_locale();
        $this->define_admin_hooks();
        $this->define_public_hooks();
    }

    private function set_locale() {
        add_action('plugins_loaded', array($this, 'load_plugin_textdomain'));
    }

    public function load_plugin_textdomain() {
        load_plugin_textdomain(
            'ai-excel-editor',
            false,
            dirname(dirname(plugin_basename(__FILE__))) . '/languages/'
        );
    }

    private function define_admin_hooks() {
        // Add menu items
        add_action('admin_menu', array($this, 'add_plugin_admin_menu'));
        // Register settings
        add_action('admin_init', array($this, 'register_settings'));
        // Add settings link
        add_filter('plugin_action_links_' . plugin_basename(__FILE__), array($this, 'add_settings_link'));
    }

    private function define_public_hooks() {
        // Enqueue scripts and styles
        add_action('wp_enqueue_scripts', array($this, 'enqueue_styles'));
        add_action('wp_enqueue_scripts', array($this, 'enqueue_scripts'));
        // Add shortcode
        add_shortcode('ai_excel_editor', array($this, 'render_editor'));
        // AJAX handlers
        add_action('wp_ajax_process_excel', array($this, 'handle_excel_processing'));
        add_action('wp_ajax_nopriv_process_excel', array($this, 'handle_unauthorized'));
    }

    public function add_plugin_admin_menu() {
        add_menu_page(
            'AI Excel Editor Settings',
            'AI Excel Editor',
            'manage_options',
            'ai-excel-editor',
            array($this, 'display_plugin_admin_page'),
            'dashicons-media-spreadsheet'
        );
    }

    public function register_settings() {
        register_setting('ai_excel_editor_settings', 'ai_excel_editor_openai_key');
    }

    public function display_plugin_admin_page() {
        include_once AI_EXCEL_EDITOR_PLUGIN_DIR . 'admin/partials/admin-display.php';
    }

    public function enqueue_styles() {
        wp_enqueue_style(
            $this->plugin_name,
            AI_EXCEL_EDITOR_PLUGIN_URL . 'assets/css/style.css',
            array(),
            $this->version
        );
    }

    public function enqueue_scripts() {
        wp_enqueue_script(
            $this->plugin_name,
            AI_EXCEL_EDITOR_PLUGIN_URL . 'assets/js/ai-excel-editor.js',
            array('jquery'),
            $this->version,
            true
        );

        wp_localize_script($this->plugin_name, 'aiExcelEditor', array(
            'ajax_url' => admin_url('admin-ajax.php'),
            'nonce' => wp_create_nonce('ai_excel_editor_nonce')
        ));
    }
}
