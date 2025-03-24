<?php
// If this file is called directly, abort.
if (!defined('ABSPATH')) {
    exit;
}
?>

<div class="wrap">
    <h1><?php echo esc_html(get_admin_page_title()); ?></h1>
    
    <div class="notice notice-info">
        <p>
            <strong>Security Note:</strong> For production use, it's recommended to set your OpenAI API key in your wp-config.php file rather than storing it in the database.
            Add this line to your wp-config.php file:
            <code>define('AISHEETS_OPENAI_API_KEY', 'your_api_key_here');</code>
        </p>
    </div>

    <?php
    // Check if API key is defined in wp-config.php
    $api_key_configured_in_wp_config = defined('AISHEETS_OPENAI_API_KEY') && !empty(AISHEETS_OPENAI_API_KEY);
    
    if ($api_key_configured_in_wp_config) :
    ?>
        <div class="notice notice-success">
            <p>
                <strong>API Key Status:</strong> Your OpenAI API key is configured in wp-config.php. This is the recommended and most secure method.
                The API key field below is disabled as the wp-config.php setting takes precedence.
            </p>
        </div>
    <?php endif; ?>

    <form method="post" action="options.php">
        <?php
        settings_fields('ai_excel_editor_settings');
        do_settings_sections('ai_excel_editor_settings');
        ?>
        
        <table class="form-table">
            <tr valign="top">
                <th scope="row">OpenAI API Key</th>
                <td>
                    <?php if ($api_key_configured_in_wp_config) : ?>
                        <input type="text" 
                               value="API key is configured in wp-config.php" 
                               class="regular-text"
                               disabled
                        />
                        <p class="description">
                            The API key is currently set in wp-config.php. To change it, edit your wp-config.php file.
                        </p>
                    <?php else : ?>
                        <input type="text" 
                               name="ai_excel_editor_openai_key" 
                               value="<?php echo esc_attr(get_option('ai_excel_editor_openai_key')); ?>" 
                               class="regular-text"
                        />
                        <p class="description">
                            Enter your OpenAI API key. Get one from 
                            <a href="https://platform.openai.com/account/api-keys" target="_blank">
                                OpenAI's website
                            </a>
                        </p>
                    <?php endif; ?>
                </td>
            </tr>
        </table>
        
        <h2>Diagnostics</h2>
        <div id="aisheets-diagnostics" style="margin-top: 20px;">
            <button type="button" id="run-diagnostics" class="button button-secondary">Run Diagnostics</button>
            <div id="diagnostics-results" style="margin-top: 15px; padding: 10px; background: #f8f8f8; display: none; border-left: 4px solid #007cba;"></div>
        </div>
        
        <script>
        jQuery(document).ready(function($) {
            $('#run-diagnostics').on('click', function() {
                var $button = $(this);
                var $results = $('#diagnostics-results');
                
                $button.prop('disabled', true).text('Running...');
                $results.html('<p>Running diagnostics, please wait...</p>').show();
                
                $.ajax({
                    url: ajaxurl,
                    type: 'POST',
                    data: {
                        action: 'aisheets_debug',
                        nonce: '<?php echo wp_create_nonce('ai_excel_editor_nonce'); ?>'
                    },
                    success: function(response) {
                        if (response.success) {
                            var html = '<h3>Diagnostics Results</h3>';
                            
                            // Add API Key Status
                            html += '<h4>API Key Status</h4>';
                            html += '<ul>';
                            html += '<li><strong>API Key Source:</strong> ' + 
                                    (<?php echo $api_key_configured_in_wp_config ? 'true' : 'false'; ?> ? 
                                     'wp-config.php' : 'WordPress Database') + '</li>';
                            html += '<li><strong>API Key Configured:</strong> ' + 
                                    (<?php echo !empty(get_option('ai_excel_editor_openai_key')) || $api_key_configured_in_wp_config ? 'true' : 'false'; ?> ? 
                                     'Yes' : 'No') + '</li>';
                            html += '</ul>';
                            
                            // PHP Configuration
                            html += '<h4>PHP Configuration</h4>';
                            html += '<ul>';
                            for (var key in response.data.php_config) {
                                html += '<li><strong>' + key + ':</strong> ' + response.data.php_config[key] + '</li>';
                            }
                            html += '</ul>';
                            
                            // Directories
                            html += '<h4>Directories</h4>';
                            html += '<ul>';
                            for (var key in response.data.directories) {
                                var value = response.data.directories[key];
                                html += '<li><strong>' + key + ':</strong> ' + (value === true ? '✅ Yes' : (value === false ? '❌ No' : value)) + '</li>';
                            }
                            html += '</ul>';
                            
                            // WordPress Info
                            html += '<h4>WordPress Information</h4>';
                            html += '<ul>';
                            html += '<li><strong>WordPress Version:</strong> ' + response.data.wp_version + '</li>';
                            html += '<li><strong>AJAX URL:</strong> ' + response.data.ajax_url + '</li>';
                            html += '</ul>';
                            
                            $results.html(html);
                        } else {
                            $results.html('<p>Error running diagnostics: ' + (response.data ? response.data.message : 'Unknown error') + '</p>');
                        }
                    },
                    error: function(xhr, status, error) {
                        $results.html('<p>AJAX error: ' + error + '</p>');
                    },
                    complete: function() {
                        $button.prop('disabled', false).text('Run Diagnostics');
                    }
                });
            });
        });
        </script>
        
        <?php 
        // Only show submit button if API key is not configured in wp-config.php
        if (!$api_key_configured_in_wp_config) {
            submit_button('Save Settings');
        }
        ?>
    </form>
</div>