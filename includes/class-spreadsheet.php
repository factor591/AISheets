<?php
/**
 * Spreadsheet handling class for AI Excel Editor
 *
 * Provides functionality for reading, modifying, and saving Excel and CSV files
 * using the PhpSpreadsheet library.
 */

// If this file is called directly, abort.
if (!defined('ABSPATH')) {
    exit;
}

class AISheets_Spreadsheet {
    /**
     * Read an Excel or CSV file and extract its data
     *
     * @param string $file_path Path to the file
     * @return array Structured data from the spreadsheet
     * @throws Exception If the file cannot be read
     */
    public function read_file($file_path) {
        if (!file_exists($file_path)) {
            throw new Exception("File not found: $file_path");
        }

        aisheets_debug('Reading spreadsheet file: ' . basename($file_path));

        try {
            // Get file extension
            $file_extension = strtolower(pathinfo($file_path, PATHINFO_EXTENSION));
            
            // Check if vendor directory exists and PhpSpreadsheet is properly installed
            $vendor_path = AI_EXCEL_EDITOR_PLUGIN_DIR . 'vendor/autoload.php';
            if (!file_exists($vendor_path)) {
                aisheets_debug('Vendor autoload.php not found at: ' . $vendor_path);
                throw new Exception('Required dependency files not found. Please ensure PhpSpreadsheet is installed.');
            }
    
            // Include required libraries
            require_once $vendor_path;
            
            // Load the appropriate reader
            $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($file_path);
            
            // For CSV files, set specific settings
            if ($file_extension === 'csv') {
                $reader->setDelimiter(',');
                $reader->setEnclosure('"');
                $reader->setSheetIndex(0);
            }
            
            // Read the spreadsheet
            $spreadsheet = $reader->load($file_path);
            
            // Extract data from all worksheets
            $data = [
                'file_type' => $file_extension,
                'file_name' => basename($file_path),
                'worksheets' => []
            ];
            
            foreach ($spreadsheet->getWorksheetIterator() as $worksheet) {
                $worksheet_name = $worksheet->getTitle();
                $worksheet_data = $this->extract_worksheet_data($worksheet);
                
                $data['worksheets'][$worksheet_name] = $worksheet_data;
            }
            
            aisheets_debug('Successfully read spreadsheet with ' . count($data['worksheets']) . ' worksheets');
            return $data;
            
        } catch (Exception $e) {
            aisheets_debug('Error reading spreadsheet file: ' . $e->getMessage());
            throw new Exception('Failed to read spreadsheet file: ' . $e->getMessage());
        }
    }
    
    /**
     * Extract data from a worksheet
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @return array Structured worksheet data
     */
    private function extract_worksheet_data($worksheet) {
        // Get worksheet dimensions
        $highest_row = $worksheet->getHighestRow();
        $highest_column = $worksheet->getHighestColumn();
        $highest_column_index = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highest_column);
        
        // Handle empty worksheets
        if ($highest_row <= 1 && $highest_column_index <= 1) {
            return [
                'headers' => [],
                'data' => [],
                'total_rows' => 0,
                'total_columns' => 0
            ];
        }
        
        // Extract headers (first row)
        $headers = [];
        for ($col = 1; $col <= $highest_column_index; $col++) {
            $cell_value = $worksheet->getCellByColumnAndRow($col, 1)->getValue();
            $headers[] = $cell_value;
        }
        
        // Extract data rows
        $data = [];
        for ($row = 2; $row <= $highest_row; $row++) {
            $row_data = [];
            for ($col = 1; $col <= $highest_column_index; $col++) {
                $cell = $worksheet->getCellByColumnAndRow($col, $row);
                $value = $cell->getValue();
                
                // Handle formulas
                if ($cell->getDataType() === \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA) {
                    $formula = $cell->getValue();
                    $calculated = $cell->getCalculatedValue();
                    $row_data[$headers[$col - 1]] = [
                        'type' => 'formula',
                        'formula' => $formula,
                        'value' => $calculated
                    ];
                } else {
                    $row_data[$headers[$col - 1]] = $value;
                }
            }
            $data[] = $row_data;
        }
        
        return [
            'headers' => $headers,
            'data' => $data,
            'total_rows' => $highest_row - 1, // Excluding header row
            'total_columns' => $highest_column_index
        ];
    }
    
    /**
     * Apply changes to a spreadsheet and save it
     *
     * @param string $original_file_path Path to the original file
     * @param array $changes Changes to apply
     * @param string $output_file_path Path to save the modified file
     * @return string Path to the saved file
     * @throws Exception If changes cannot be applied
     */
    public function apply_changes($original_file_path, $changes, $output_file_path) {
        aisheets_debug('Applying changes to spreadsheet');
        
        try {
            // Load the original spreadsheet
            $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($original_file_path);
            $spreadsheet = $reader->load($original_file_path);
            
            // Apply each change
            if (empty($changes)) {
                // If no changes provided, just copy the file
                aisheets_debug('No changes to apply, copying file directly');
                
                // In this temporary implementation, just copy the file
                if (!copy($original_file_path, $output_file_path)) {
                    throw new Exception('Failed to copy file for temporary implementation');
                }
                
                aisheets_debug('File copied successfully as temporary implementation');
                return $output_file_path;
            }
            
            foreach ($changes as $worksheet_name => $worksheet_changes) {
                // Get the worksheet (create it if it doesn't exist)
                aisheets_debug('Processing changes for worksheet: ' . $worksheet_name);
                
                $worksheet = null;
                if ($spreadsheet->sheetNameExists($worksheet_name)) {
                    $worksheet = $spreadsheet->getSheetByName($worksheet_name);
                } else {
                    $worksheet = $spreadsheet->createSheet();
                    $worksheet->setTitle($worksheet_name);
                }
                
                // Apply specific changes
                if (isset($worksheet_changes['cell_changes'])) {
                    $this->apply_cell_changes($worksheet, $worksheet_changes['cell_changes']);
                }
                
                if (isset($worksheet_changes['column_changes'])) {
                    $this->apply_column_changes($worksheet, $worksheet_changes['column_changes']);
                }
                
                if (isset($worksheet_changes['row_changes'])) {
                    $this->apply_row_changes($worksheet, $worksheet_changes['row_changes']);
                }
                
                if (isset($worksheet_changes['format_changes'])) {
                    $this->apply_format_changes($worksheet, $worksheet_changes['format_changes']);
                }
                
                if (isset($worksheet_changes['sort_changes'])) {
                    $this->apply_sort_changes($worksheet, $worksheet_changes['sort_changes']);
                }
            }
            
            // Determine the output format based on the file extension
            $extension = strtolower(pathinfo($output_file_path, PATHINFO_EXTENSION));
            
            // Create the appropriate writer
            if ($extension === 'csv') {
                $writer = new \PhpOffice\PhpSpreadsheet\Writer\Csv($spreadsheet);
                $writer->setDelimiter(',');
                $writer->setEnclosure('"');
                $writer->setLineEnding("\r\n");
                $writer->setSheetIndex(0);
            } else if ($extension === 'xls') {
                $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xls($spreadsheet);
            } else {
                // Default to xlsx
                $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
            }
            
            // Save the file
            $writer->save($output_file_path);
            
            aisheets_debug('Changes applied and file saved to: ' . $output_file_path);
            return $output_file_path;
            
        } catch (Exception $e) {
            aisheets_debug('Error applying changes to spreadsheet: ' . $e->getMessage());
            throw new Exception('Failed to apply changes to spreadsheet: ' . $e->getMessage());
        }
    }
    
    /**
     * Apply changes to individual cells
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param array $cell_changes Array of cell changes
     */
    private function apply_cell_changes($worksheet, $cell_changes) {
        foreach ($cell_changes as $cell_ref => $change) {
            if (isset($change['value'])) {
                $worksheet->setCellValue($cell_ref, $change['value']);
            } else if (isset($change['formula'])) {
                $worksheet->setCellValue($cell_ref, $change['formula']);
            }
        }
    }
    
    /**
     * Apply changes to columns
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param array $column_changes Array of column changes
     */
    private function apply_column_changes($worksheet, $column_changes) {
        foreach ($column_changes as $column_id => $change) {
            // Get column letter if numeric index provided
            $column_letter = is_numeric($column_id) 
                ? \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($column_id) 
                : $column_id;
            
            // Handle 'add' operation
            if (isset($change['operation']) && $change['operation'] === 'add') {
                // Get highest row in worksheet to know how far to fill the column
                $highest_row = $worksheet->getHighestRow();
                
                // Add column header
                if (isset($change['header'])) {
                    $worksheet->setCellValue($column_letter . '1', $change['header']);
                }
                
                // Add column values or formula
                if (isset($change['values'])) {
                    $row_num = 2; // Start after header row
                    foreach ($change['values'] as $value) {
                        if ($row_num <= $highest_row) {
                            $worksheet->setCellValue($column_letter . $row_num, $value);
                        }
                        $row_num++;
                    }
                } else if (isset($change['formula'])) {
                    // Apply formula to all rows
                    for ($row = 2; $row <= $highest_row; $row++) {
                        // Replace any row placeholders in the formula (e.g., {row})
                        $row_formula = str_replace('{row}', $row, $change['formula']);
                        $worksheet->setCellValue($column_letter . $row, $row_formula);
                    }
                }
            }
            
            // Handle 'delete' operation
            else if (isset($change['operation']) && $change['operation'] === 'delete') {
                $worksheet->removeColumn($column_letter);
            }
        }
    }
    
    /**
     * Apply changes to rows
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param array $row_changes Array of row changes
     */
    private function apply_row_changes($worksheet, $row_changes) {
        foreach ($row_changes as $row_id => $change) {
            // Handle 'add' operation
            if (isset($change['operation']) && $change['operation'] === 'add') {
                // Insert a new row before
                $worksheet->insertNewRowBefore($row_id, 1);
                
                // Add row values
                if (isset($change['values'])) {
                    $col_num = 1;
                    foreach ($change['values'] as $value) {
                        $worksheet->setCellValueByColumnAndRow($col_num, $row_id, $value);
                        $col_num++;
                    }
                }
            }
            
            // Handle 'delete' operation
            else if (isset($change['operation']) && $change['operation'] === 'delete') {
                $worksheet->removeRow($row_id);
            }
        }
    }
    
    /**
     * Apply formatting changes
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param array $format_changes Array of formatting changes
     */
    private function apply_format_changes($worksheet, $format_changes) {
        foreach ($format_changes as $range => $change) {
            // Get the style for the range
            $style = $worksheet->getStyle($range);
            
            // Number format
            if (isset($change['number_format'])) {
                $style->getNumberFormat()->setFormatCode($change['number_format']);
            }
            
            // Font changes
            if (isset($change['font'])) {
                $font = $style->getFont();
                
                if (isset($change['font']['bold'])) {
                    $font->setBold($change['font']['bold']);
                }
                
                if (isset($change['font']['italic'])) {
                    $font->setItalic($change['font']['italic']);
                }
                
                if (isset($change['font']['underline'])) {
                    $font->setUnderline($change['font']['underline']);
                }
                
                if (isset($change['font']['size'])) {
                    $font->setSize($change['font']['size']);
                }
                
                if (isset($change['font']['color'])) {
                    $font->getColor()->setRGB($change['font']['color']);
                }
            }
            
            // Fill/background color
            if (isset($change['fill'])) {
                $fill = $style->getFill();
                
                if (isset($change['fill']['type'])) {
                    $fill->setFillType($change['fill']['type']);
                }
                
                if (isset($change['fill']['color'])) {
                    $fill->getStartColor()->setRGB($change['fill']['color']);
                }
            }
            
            // Borders
            if (isset($change['borders'])) {
                $borders = $style->getBorders();
                
                foreach ($change['borders'] as $position => $border) {
                    $border_obj = null;
                    
                    switch ($position) {
                        case 'outline':
                            $border_obj = $borders->getAllBorders();
                            break;
                        case 'inside':
                            $border_obj = $borders->getInsideBorders();
                            break;
                        case 'horizontal':
                            $border_obj = $borders->getHorizontal();
                            break;
                        case 'vertical':
                            $border_obj = $borders->getVertical();
                            break;
                        case 'top':
                            $border_obj = $borders->getTop();
                            break;
                        case 'right':
                            $border_obj = $borders->getRight();
                            break;
                        case 'bottom':
                            $border_obj = $borders->getBottom();
                            break;
                        case 'left':
                            $border_obj = $borders->getLeft();
                            break;
                    }
                    
                    if ($border_obj) {
                        if (isset($border['style'])) {
                            $border_obj->setBorderStyle($border['style']);
                        }
                        
                        if (isset($border['color'])) {
                            $border_obj->getColor()->setRGB($border['color']);
                        }
                    }
                }
            }
            
            // Alignment
            if (isset($change['alignment'])) {
                $alignment = $style->getAlignment();
                
                if (isset($change['alignment']['horizontal'])) {
                    $alignment->setHorizontal($change['alignment']['horizontal']);
                }
                
                if (isset($change['alignment']['vertical'])) {
                    $alignment->setVertical($change['alignment']['vertical']);
                }
                
                if (isset($change['alignment']['wrap_text'])) {
                    $alignment->setWrapText($change['alignment']['wrap_text']);
                }
            }
        }
    }
    
    /**
     * Apply sorting changes
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param array $sort_changes Array of sorting changes
     */
    private function apply_sort_changes($worksheet, $sort_changes) {
        if (!empty($sort_changes)) {
            // Get the range to sort
            $range = $sort_changes['range'] ?? null;
            
            if (!$range) {
                // If no range specified, use the whole data area (excluding headers)
                $highest_column = $worksheet->getHighestColumn();
                $highest_row = $worksheet->getHighestRow();
                $range = 'A2:' . $highest_column . $highest_row;
            }
            
            // Get the sort column
            $sort_column = $sort_changes['column'] ?? null;
            
            if ($sort_column) {
                // Convert column name or number to the appropriate format
                if (is_numeric($sort_column)) {
                    $sort_column = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($sort_column);
                }
                
                // Get the sort direction
                $sort_direction = $sort_changes['direction'] ?? 'ASC';
                $sort_direction = strtoupper($sort_direction);
                
                // Sort the range
                $worksheet->getStyle($range)->getAlignment()->setWrapText(true);
                
                // Extract data for sorting
                list($start_col, $start_row, $end_col, $end_row) = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::rangeBoundaries($range);
                
                $data_array = [];
                for ($row = $start_row; $row <= $end_row; $row++) {
                    $row_data = [];
                    for ($col = $start_col; $col <= $end_col; $col++) {
                        $cell_coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
                        $row_data[] = $worksheet->getCell($cell_coord)->getValue();
                    }
                    $data_array[] = $row_data;
                }
                
                // Find the sort column index
                $sort_col_index = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($sort_column) - $start_col;
                
                // Sort the data
                array_multisort(
                    array_column($data_array, $sort_col_index),
                    $sort_direction === 'DESC' ? SORT_DESC : SORT_ASC,
                    $data_array
                );
                
                // Write back the sorted data
                foreach ($data_array as $row_idx => $row_data) {
                    foreach ($row_data as $col_idx => $cell_value) {
                        $cell_coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($start_col + $col_idx) . ($start_row + $row_idx);
                        $worksheet->setCellValue($cell_coord, $cell_value);
                    }
                }
            }
        }
    }
}