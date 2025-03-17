<?php
// Test script for AISheets WordPress plugin
// Access this file directly in your browser to run tests

// Set content type for readable output
header('Content-Type: text/plain');

echo "AISheets Plugin Test Script\n";
echo "==========================\n\n";

echo "PHP Version: " . phpversion() . "\n";
echo "Current Directory: " . __DIR__ . "\n\n";

// Test for required PHP extensions
echo "Checking required PHP extensions...\n";
$required_extensions = ['curl', 'fileinfo', 'json', 'zip'];
$missing_extensions = [];

foreach ($required_extensions as $ext) {
    if (!extension_loaded($ext)) {
        $missing_extensions[] = $ext;
        echo "❌ {$ext} extension is missing\n";
    } else {
        echo "✓ {$ext} extension is loaded\n";
    }
}

if (!empty($missing_extensions)) {
    echo "\nWARNING: Some required extensions are missing. Please install them before continuing.\n\n";
} else {
    echo "\nAll required extensions are loaded.\n\n";
}

// Test for vendor directory
echo "Checking for vendor directory...\n";
$vendor_path = __DIR__ . '/vendor';
$autoload_path = $vendor_path . '/autoload.php';

if (!is_dir($vendor_path)) {
    echo "❌ Vendor directory not found at: {$vendor_path}\n";
    echo "   Please run 'composer install' or upload the vendor directory.\n\n";
} else {
    echo "✓ Vendor directory found at: {$vendor_path}\n";
    
    if (!file_exists($autoload_path)) {
        echo "❌ Autoload.php not found at: {$autoload_path}\n";
        echo "   Your Composer installation may be incomplete.\n\n";
    } else {
        echo "✓ Autoload.php found at: {$autoload_path}\n\n";
        
        // Test PhpSpreadsheet
        echo "Testing PhpSpreadsheet library...\n";
        try {
            require_once $autoload_path;
            
            if (!class_exists('PhpOffice\\PhpSpreadsheet\\Spreadsheet')) {
                echo "❌ PhpSpreadsheet class not found\n";
                echo "   Please make sure PhpSpreadsheet is installed correctly via Composer.\n\n";
            } else {
                echo "✓ PhpSpreadsheet class found\n";
                
                // Try creating a simple spreadsheet
                $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
                $sheet = $spreadsheet->getActiveSheet();
                $sheet->setCellValue('A1', 'Hello World!');
                echo "✓ Successfully created a test spreadsheet\n\n";
            }
        } catch (Exception $e) {
            echo "❌ Error loading PhpSpreadsheet: " . $e->getMessage() . "\n\n";
        }
    }
}

// Test for includes directory
echo "Checking for includes directory...\n";
$includes_path = __DIR__ . '/includes';

if (!is_dir($includes_path)) {
    echo "❌ Includes directory not found at: {$includes_path}\n";
} else {
    echo "✓ Includes directory found at: {$includes_path}\n";
    
    // Check for required class files
    $required_files = [
        'class-activator.php',
        'class-spreadsheet.php',
        'class-openai.php'
    ];
    
    foreach ($required_files as $file) {
        $file_path = $includes_path . '/' . $file;
        if (!file_exists($file_path)) {
            echo "❌ Required file missing: {$file}\n";
        } else {
            echo "✓ Found required file: {$file}\n";
        }
    }
    echo "\n";
}

// Check upload directory
echo "Checking upload directory...\n";

// Try to determine the WordPress upload directory
$wp_upload_dir = null;
if (function_exists('wp_upload_dir')) {
    $wp_upload_dir = wp_upload_dir();
} else {
    // If we're running standalone, try to guess based on common paths
    $possible_wp_dirs = [
        dirname(__DIR__, 2), // Two levels up from plugin
        dirname(__DIR__, 3), // Three levels up from plugin
    ];
    
    foreach ($possible_wp_dirs as $wp_dir) {
        if (file_exists($wp_dir . '/wp-config.php')) {
            // Found WordPress installation
            if (is_dir($wp_dir . '/wp-content/uploads')) {
                $wp_upload_dir = [
                    'basedir' => $wp_dir . '/wp-content/uploads'
                ];
                break;
            }
        }
    }
}

if (!$wp_upload_dir) {
    echo "❌ Could not determine WordPress upload directory\n";
    echo "   This test script may not be running within WordPress environment.\n\n";
} else {
    $upload_path = $wp_upload_dir['basedir'] . '/ai-excel-editor';
    
    if (!is_dir($upload_path)) {
        echo "❌ AISheets upload directory not found at: {$upload_path}\n";
        echo "   The directory will be created when the plugin initializes.\n\n";
    } else {
        echo "✓ AISheets upload directory found at: {$upload_path}\n";
        echo "   Directory permissions: " . substr(sprintf('%o', fileperms($upload_path)), -4) . "\n";
        
        // Check if we can write to the directory
        $test_file = $upload_path . '/test_' . uniqid() . '.txt';
        $write_test = false;
        
        try {
            $write_test = file_put_contents($test_file, 'Test file for AISheets plugin.');
            if ($write_test !== false) {
                echo "✓ Successfully wrote test file to upload directory\n";
                unlink($test_file); // Clean up
            } else {
                echo "❌ Could not write to upload directory\n";
            }
        } catch (Exception $e) {
            echo "❌ Error writing to upload directory: " . $e->getMessage() . "\n";
        }
    }
}

echo "\nTest completed.\n";
echo "==========================\n";
echo "If any issues were found, please address them before using the plugin.\n";
?>t