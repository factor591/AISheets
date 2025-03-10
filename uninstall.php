<?php
// If uninstall not called from WordPress, exit
if (!defined('WP_UNINSTALL_PLUGIN')) {
    exit;
}

// Call uninstall method
require_once plugin_dir_path(__FILE__) . 'includes/class-activator.php';
AI_Excel_Editor_Activator::uninstall();
