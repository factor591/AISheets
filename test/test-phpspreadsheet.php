// Create a simple test script in your plugin directory
require_once 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function test_phpspreadsheet() {
    try {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'Hello World');
        
        $writer = new Xlsx($spreadsheet);
        $test_file = plugin_dir_path(__FILE__) . 'test_output.xlsx';
        $writer->save($test_file);
        
        return file_exists($test_file) ? 'Success: PhpSpreadsheet is working' : 'Error: Could not create file';
    } catch (Exception $e) {
        return 'Error: ' . $e->getMessage();
    }
}

// Add to a test admin page or temporary function