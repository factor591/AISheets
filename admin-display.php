<?php
// If this file is called directly, abort.
if (!defined('ABSPATH')) {
    exit;
}
?>

<div class="wrap">
    <h1><?php echo esc_html(get_admin_page_title()); ?></h1>

    <form method="post" action="options.php">
        <?php
        settings_fields('ai_excel_editor_settings');
        do_settings_sections('ai_excel_editor_settings');
        ?>
        
        <table class="form-table">
            <tr valign="top">
                <th scope="row">OpenAI API Key</th>
                <td>
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
                </td>
            </tr>
        </table>
        
        <?php submit_button(); ?>
    </form>
</div>
