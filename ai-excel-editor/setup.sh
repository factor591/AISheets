# Create main plugin directory
mkdir -p wp-content/uploads/ai-excel-editor/processing

# Create .htaccess files
echo '# Deny direct access
<IfModule mod_rewrite.c>
RewriteEngine On
RewriteRule ^.*$ - [F,L]
</IfModule>
# Protect directory
Options -Indexes' > wp-content/uploads/ai-excel-editor/.htaccess

# Copy the same .htaccess to processing directory
cp wp-content/uploads/ai-excel-editor/.htaccess wp-content/uploads/ai-excel-editor/processing/.htaccess

# Create index.php files
echo '<?php // Silence is golden' > wp-content/uploads/ai-excel-editor/index.php
echo '<?php // Silence is golden' > wp-content/uploads/ai-excel-editor/processing/index.php