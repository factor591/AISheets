<?php
/**
 * OpenAI Integration for AI Excel Editor
 */

// Prevent direct access
if (!defined('ABSPATH')) {
    exit;
}

class AISheets_OpenAI_Integration {
    private $api_key;
    private $api_endpoint = 'https://api.openai.com/v1/chat/completions';
    private $model = 'gpt-4';
    private $temperature = 0.2;
    private $max_tokens = 4000;
    
    /**
     * Constructor
     * 
     * @param string $api_key OpenAI API key
     */
    public function __construct($api_key) {
        $this->api_key = $api_key;
        
        if (empty($this->api_key)) {
            aisheets_debug('OpenAI Integration: API key not provided');
            throw new Exception('OpenAI API key is required but was not provided');
        }
        
        aisheets_debug('OpenAI Integration: Successfully initialized');
    }
    
    /**
     * Process spreadsheet with OpenAI
     * 
     * @param string $file_path Path to the spreadsheet file
     * @param string $instructions User's natural language instructions
     * @return array OpenAI response data
     */
    public function process_spreadsheet($file_path, $instructions) {
        aisheets_debug('Processing spreadsheet with OpenAI');
        
        // 1. Read the spreadsheet
        $spreadsheet_data = $this->read_spreadsheet($file_path);
        
        // 2. Prepare data for OpenAI (limit size to avoid token limits)
        $sample_data = $this->prepare_sample_data($spreadsheet_data);
        
        // 3. Generate API request
        $request_data = $this->prepare_api_request($sample_data, $instructions);
        
        // 4. Call OpenAI API
        $response = $this->call_api($request_data);
        
        return $response;
    }
    
    /**
     * Apply changes from OpenAI response to spreadsheet
     * 
     * @param string $input_path Path to the original spreadsheet
     * @param string $output_path Path to save the modified spreadsheet
     * @param array $response OpenAI API response
     * @return bool True if changes were applied successfully
     */
    public function apply_changes_to_spreadsheet($input_path, $output_path, $response) {
        aisheets_debug('Applying changes to spreadsheet');
        
        try {
            // Extract function call arguments from response
            $function_call = $this->extract_function_call($response);
            
            if (!$function_call) {
                aisheets_debug('No function call in OpenAI response');
                return false;
            }
            
            // Parse the changes to apply
            $changes = $this->parse_changes($function_call);
            
            if (empty($changes)) {
                aisheets_debug('No changes to apply from OpenAI response');
                return false;
            }
            
            // Apply changes to spreadsheet
            return $this->modify_spreadsheet($input_path, $output_path, $changes);
            
        } catch (Exception $e) {
            aisheets_debug('Error applying changes: ' . $e->getMessage());
            return false;
        }
    }
    
    /**
     * Read spreadsheet file and extract data
     * 
     * @param string $file_path Path to the spreadsheet file
     * @return array Extracted spreadsheet data
     */
    private function read_spreadsheet($file_path) {
        aisheets_debug('Reading spreadsheet: ' . $file_path);
        
        // Load the PhpSpreadsheet library
        require_once AI_EXCEL_EDITOR_PLUGIN_DIR . 'vendor/autoload.php';
        
        $file_ext = strtolower(pathinfo($file_path, PATHINFO_EXTENSION));
        $spreadsheet_data = [];
        
        try {
            if ($file_ext === 'csv') {
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
                // Configure CSV reader options
                $reader->setDelimiter(',');
                $reader->setEnclosure('"');
                $reader->setSheetIndex(0);
            } else {
                $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($file_path);
            }
            
            // Set read data only (ignore formatting)
            $reader->setReadDataOnly(true);
            
            $spreadsheet = $reader->load($file_path);
            
            // Extract data from all worksheets
            $worksheets = [];
            foreach ($spreadsheet->getWorksheetIterator() as $worksheet) {
                $sheet_name = $worksheet->getTitle();
                $sheet_data = [];
                
                // Get worksheet dimensions
                $highest_row = $worksheet->getHighestRow();
                $highest_column = $worksheet->getHighestColumn();
                $highest_column_index = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highest_column);
                
                // Get headers (first row)
                $headers = [];
                for ($col = 1; $col <= $highest_column_index; $col++) {
                    $cell_value = $worksheet->getCellByColumnAndRow($col, 1)->getValue();
                    $headers[] = $cell_value;
                }
                
                // Get data rows
                $rows = [];
                for ($row = 2; $row <= $highest_row; $row++) {
                    $row_data = [];
                    for ($col = 1; $col <= $highest_column_index; $col++) {
                        $cell_value = $worksheet->getCellByColumnAndRow($col, $row)->getValue();
                        $row_data[$headers[$col-1]] = $cell_value;
                    }
                    $rows[] = $row_data;
                }
                
                $sheet_data = [
                    'headers' => $headers,
                    'rows' => $rows,
                    'total_rows' => $highest_row - 1, // Excluding header row
                    'total_columns' => $highest_column_index
                ];
                
                $worksheets[$sheet_name] = $sheet_data;
            }
            
            $spreadsheet_data = [
                'file_type' => $file_ext,
                'worksheets' => $worksheets
            ];
            
            aisheets_debug('Successfully read spreadsheet', [
                'worksheets' => count($worksheets),
                'first_worksheet_rows' => isset(reset($worksheets)['total_rows']) ? reset($worksheets)['total_rows'] : 0
            ]);
            
            return $spreadsheet_data;
            
        } catch (\Exception $e) {
            aisheets_debug('Error reading spreadsheet: ' . $e->getMessage());
            throw new \Exception('Failed to read spreadsheet: ' . $e->getMessage());
        }
    }
    
    /**
     * Prepare a sample of spreadsheet data to avoid token limits
     * 
     * @param array $spreadsheet_data Full spreadsheet data
     * @return array Sample data for OpenAI API
     */
    private function prepare_sample_data($spreadsheet_data) {
        aisheets_debug('Preparing sample data for OpenAI');
        
        $sample = [
            'file_type' => $spreadsheet_data['file_type'],
            'worksheets' => []
        ];
        
        foreach ($spreadsheet_data['worksheets'] as $sheet_name => $sheet_data) {
            $headers = $sheet_data['headers'];
            $total_rows = count($sheet_data['rows']);
            
            // Take at most 15 rows as a sample
            $max_sample_rows = 5;
            $sample_rows = array_slice($sheet_data['rows'], 0, $max_sample_rows);
            
            // If we have more than sample size rows, include some from the end too
            if ($total_rows > $max_sample_rows && $total_rows > $max_sample_rows * 2) {
                $end_sample = array_slice($sheet_data['rows'], -5, 5);
                $sample_rows = array_merge(
                    array_slice($sample_rows, 0, 10),
                    $end_sample
                );
            }
            
            $sample['worksheets'][$sheet_name] = [
                'headers' => $headers,
                'rows' => $sample_rows,
                'total_rows' => $total_rows,
                'total_columns' => $sheet_data['total_columns'],
                'is_sample' => $total_rows > count($sample_rows)
            ];
        }
        
        aisheets_debug('Sample data prepared');
        return $sample;
    }
    
/**
 * Prepare API request for OpenAI
 * 
 * @param array $sample_data Spreadsheet sample data
 * @param string $instructions User's instructions
 * @return array OpenAI API request data
 */
private function prepare_api_request($sample_data, $instructions) {
    aisheets_debug('Preparing OpenAI API request');
    
    return [
        'model' => $this->model,
        'messages' => [
            [
                'role' => 'system',
                'content' => $this->get_system_prompt()
            ],
            [
                'role' => 'user',
                'content' => "Here is the spreadsheet data:\n" . json_encode($sample_data) . "\n\nInstructions: " . $instructions
            ]
        ],
        'functions' => [
            [
                'name' => 'update_spreadsheet',
                'description' => 'Apply changes to the spreadsheet based on the user instructions',
                'parameters' => [
                    'type' => 'object',
                    'properties' => [
                        'changes' => [
                            'type' => 'array',
                            'description' => 'List of changes to apply to the spreadsheet',
                            'items' => [
                                'type' => 'object',
                                'properties' => [
                                    'worksheet' => [
                                        'type' => 'string',
                                        'description' => 'Name of the worksheet to modify'
                                    ],
                                    'type' => [
                                        'type' => 'string',
                                        'enum' => ['value', 'formula', 'format', 'sort', 'add_column', 'add_row', 'delete_row', 'delete_column'],
                                        'description' => 'Type of change to apply'
                                    ],
                                    'target' => [
                                        'type' => 'object',
                                        'description' => 'Target of the change',
                                        'properties' => [
                                            'type' => [
                                                'type' => 'string',
                                                'enum' => ['cell', 'range', 'column', 'row'],
                                                'description' => 'Type of target'
                                            ],
                                            'reference' => [
                                                'type' => 'string',
                                                'description' => 'Cell reference (e.g., A1), range (e.g., A1:C10), column name/letter, or row number'
                                            ]
                                        ]
                                    ],
                                    'value' => [
                                        'type' => 'string',
                                        'description' => 'Plain text value to set (not JSON objects)'
                                    ],
                                    'parameters' => [
                                        'type' => 'object',
                                        'description' => 'Additional parameters for the change',
                                        'properties' => [
                                            'values_by_column' => [
                                                'type' => 'object',
                                                'description' => 'For adding rows, provide individual values per column',
                                                'properties' => [
                                                    'Name' => ['type' => 'string'],
                                                    'Email' => ['type' => 'string'],
                                                    'Phone' => ['type' => 'string']
                                                ]
                                            ]
                                        ]
                                    ]
                                ],
                                'required' => ['worksheet', 'type', 'target']
                            ]
                        ],
                        'explanation' => [
                            'type' => 'string',
                            'description' => 'Explanation of the changes made to the spreadsheet'
                        ]
                    ],
                    'required' => ['changes', 'explanation']
                ]
            ]
        ],
        'function_call' => ['name' => 'update_spreadsheet'],
        'temperature' => $this->temperature,
        'max_tokens' => $this->max_tokens
    ];
}
    
    /**
     * Call OpenAI API
     * 
     * @param array $request_data Request data for OpenAI API
     * @return array API response
     */
    private function call_api($request_data) {
        aisheets_debug('Calling OpenAI API');
        
        if (empty($this->api_key)) {
            throw new Exception('API key is required to call the OpenAI API');
        }
        
        $headers = [
            'Content-Type: application/json',
            'Authorization: Bearer ' . $this->api_key
        ];
        
        $ch = curl_init($this->api_endpoint);
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        curl_setopt($ch, CURLOPT_POST, true);
        curl_setopt($ch, CURLOPT_POSTFIELDS, json_encode($request_data));
        curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
        curl_setopt($ch, CURLOPT_TIMEOUT, 60); // 60-second timeout
        
        $response = curl_exec($ch);
        $http_code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
        
        if (curl_errno($ch)) {
            $error = curl_error($ch);
            curl_close($ch);
            aisheets_debug('API request failed: ' . $error);
            throw new Exception('API request failed: ' . $error);
        }
        
        curl_close($ch);
        
        if ($http_code !== 200) {
            $error_details = json_decode($response, true);
            $error_message = isset($error_details['error']['message']) 
                ? $error_details['error']['message']
                : 'API returned error ' . $http_code;
            
            aisheets_debug('API returned error ' . $http_code . ': ' . $error_message);
            throw new Exception($error_message);
        }
        
        $decoded_response = json_decode($response, true);
        
        if (json_last_error() !== JSON_ERROR_NONE) {
            aisheets_debug('Failed to parse API response');
            throw new Exception('Failed to parse API response');
        }
        
        aisheets_debug('OpenAI API response received');
        return $decoded_response;
    }
    
    /**
     * Extract function call from OpenAI response
     * 
     * @param array $response OpenAI API response
     * @return array|null Function call arguments or null if not found
     */
    private function extract_function_call($response) {
        if (!isset($response['choices'][0]['message']['function_call'])) {
            return null;
        }
        
        $function_call = $response['choices'][0]['message']['function_call'];
        
        if ($function_call['name'] !== 'update_spreadsheet') {
            return null;
        }
        
        $arguments = json_decode($function_call['arguments'], true);
        
        if (json_last_error() !== JSON_ERROR_NONE) {
            aisheets_debug('Failed to parse function call arguments');
            return null;
        }
        
        return $arguments;
    }
    
    /**
     * Parse changes from function call arguments
     * 
     * @param array $function_call Function call arguments
     * @return array Changes to apply to the spreadsheet
     */
    private function parse_changes($function_call) {
        if (!isset($function_call['changes']) || !is_array($function_call['changes'])) {
            return [];
        }
        
        return $function_call['changes'];
    }
    
    /**
     * Modify spreadsheet based on changes
     * 
     * @param string $input_path Path to the original spreadsheet
     * @param string $output_path Path to save the modified spreadsheet
     * @param array $changes Changes to apply
     * @return bool True if changes were applied successfully
     */
    private function modify_spreadsheet($input_path, $output_path, $changes) {
        aisheets_debug('Modifying spreadsheet with ' . count($changes) . ' changes');
        
        try {
            // Load PhpSpreadsheet
            require_once AI_EXCEL_EDITOR_PLUGIN_DIR . 'vendor/autoload.php';
            
            // Load the spreadsheet
            $file_ext = strtolower(pathinfo($input_path, PATHINFO_EXTENSION));
            
            if ($file_ext === 'csv') {
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
                // Configure CSV reader
                $reader->setDelimiter(',');
                $reader->setEnclosure('"');
                $reader->setSheetIndex(0);
            } else {
                $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($input_path);
            }
            
            $spreadsheet = $reader->load($input_path);
            
            // Apply each change
            foreach ($changes as $change) {
                $this->apply_change($spreadsheet, $change);
            }
            
            // Save the modified spreadsheet
            if ($file_ext === 'csv') {
                $writer = new \PhpOffice\PhpSpreadsheet\Writer\Csv($spreadsheet);
            } else {
                $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, ucfirst($file_ext));
            }
            
            $writer->save($output_path);
            
            aisheets_debug('Successfully modified and saved spreadsheet');
            return true;
            
        } catch (\Exception $e) {
            aisheets_debug('Error modifying spreadsheet: ' . $e->getMessage());
            return false;
        }
    }
    
    /**
     * Apply a single change to the spreadsheet
     * 
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet Spreadsheet object
     * @param array $change Change to apply
     */
    private function apply_change($spreadsheet, $change) {
        // Validate required fields
        if (!isset($change['worksheet']) || !isset($change['type']) || !isset($change['target'])) {
            aisheets_debug('Invalid change: missing required fields');
            return;
        }
        
        // Get the worksheet
        $worksheet_name = $change['worksheet'];
        $worksheet = null;
        
        // Find the worksheet
        foreach ($spreadsheet->getWorksheetIterator() as $sheet) {
            if ($sheet->getTitle() === $worksheet_name) {
                $worksheet = $sheet;
                break;
            }
        }
        
        // If worksheet not found, try to use the first one
        if (!$worksheet && $spreadsheet->getSheetCount() > 0) {
            $worksheet = $spreadsheet->getSheet(0);
            aisheets_debug('Worksheet not found, using first worksheet');
        }
        
        if (!$worksheet) {
            aisheets_debug('No worksheet available');
            return;
        }
        
        // Apply change based on type
        $type = $change['type'];
        $target = $change['target'];
        $value = isset($change['value']) ? $change['value'] : null;
        $parameters = isset($change['parameters']) ? $change['parameters'] : [];
        
        switch ($type) {
            case 'value':
                $this->apply_value_change($worksheet, $target, $value);
                break;
                
            case 'formula':
                $this->apply_formula_change($worksheet, $target, $value);
                break;
                
            case 'format':
                $this->apply_format_change($worksheet, $target, $value, $parameters);
                break;
                
            case 'sort':
                $this->apply_sort_change($worksheet, $target, $parameters);
                break;
                
            case 'add_column':
                $this->apply_add_column_change($worksheet, $target, $value, $parameters);
                break;
                
            case 'add_row':
                $this->apply_add_row_change($worksheet, $target, $value, $parameters);
                break;
                
            case 'delete_row':
                $this->apply_delete_row_change($worksheet, $target);
                break;
                
            case 'delete_column':
                $this->apply_delete_column_change($worksheet, $target);
                break;
                
            default:
                aisheets_debug('Unsupported change type: ' . $type);
                break;
        }
    }
    
    /**
     * Apply value change to cells
     * 
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param array $target Target cell(s)
     * @param mixed $value Value to set
     */
    private function apply_value_change($worksheet, $target, $value) {
        if (!isset($target['type']) || !isset($target['reference'])) {
            return;
        }
        
        // Handle JSON-like strings by parsing them
        if (is_string($value) && (strpos($value, '{') === 0 || strpos($value, '[') === 0)) {
            try {
                $decoded = json_decode($value, true);
                if (json_last_error() === JSON_ERROR_NONE && is_array($decoded)) {
                    // If it's a valid JSON array/object, use the more appropriate values
                    if (isset($decoded['Name']) && isset($decoded['Email']) && isset($decoded['Phone'])) {
                        aisheets_debug('Detected JSON object with contact fields, extracting Name field');
                        $value = $decoded['Name']; // Use Name as default or most appropriate field
                    } else if (count($decoded) === 1 && is_string(reset($decoded))) {
                        // If it's a single value array, use that value
                        $value = reset($decoded);
                    }
                }
            } catch (Exception $e) {
                // If parsing fails, keep the original value
                aisheets_debug('Failed to parse JSON-like string: ' . $e->getMessage());
            }
        }
        
        switch ($target['type']) {
            case 'cell':
                $worksheet->getCell($target['reference'])->setValue($value);
                break;
                
            case 'range':
                $this->set_range_values($worksheet, $target['reference'], $value);
                break;
                
            case 'column':
                $this->set_column_values($worksheet, $target['reference'], $value);
                break;
                
            case 'row':
                $this->set_row_values($worksheet, $target['reference'], $value);
                break;
        }
    }
    
    /**
     * Apply formula change to cells
     * 
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param array $target Target cell(s)
     * @param string $formula Formula to set
     */
    private function apply_formula_change($worksheet, $target, $formula) {
        // Make sure the formula starts with =
        if ($formula && substr($formula, 0, 1) !== '=') {
            $formula = '=' . $formula;
        }
        
        $this->apply_value_change($worksheet, $target, $formula);
    }
    
    /**
     * Apply format change to cells
     * 
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param array $target Target cell(s)
     * @param string $format Format to apply
     * @param array $parameters Additional format parameters
     */
    private function apply_format_change($worksheet, $target, $format, $parameters) {
        if (!isset($target['type']) || !isset($target['reference'])) {
            return;
        }
        
        // Get the cell range
        $cells = $this->get_cell_range($worksheet, $target);
        
        // Apply format based on format type
        if ($format === 'currency') {
            $symbol = isset($parameters['symbol']) ? $parameters['symbol'] : '$';
            foreach ($cells as $cell) {
                $worksheet->getStyle($cell)->getNumberFormat()->setFormatCode($symbol . '#,##0.00');
            }
        } elseif ($format === 'percentage') {
            foreach ($cells as $cell) {
                $worksheet->getStyle($cell)->getNumberFormat()->setFormatCode('0.00%');
            }
        } elseif ($format === 'date') {
            $dateFormat = isset($parameters['format']) ? $parameters['format'] : 'mm/dd/yyyy';
            foreach ($cells as $cell) {
                $worksheet->getStyle($cell)->getNumberFormat()->setFormatCode($dateFormat);
            }
        } elseif ($format === 'number') {
            $decimals = isset($parameters['decimals']) ? $parameters['decimals'] : 2;
            $format_code = '0';
            if ($decimals > 0) {
                $format_code .= '.' . str_repeat('0', $decimals);
            }
            foreach ($cells as $cell) {
                $worksheet->getStyle($cell)->getNumberFormat()->setFormatCode($format_code);
            }
        } elseif ($format === 'text_format') {
            foreach ($cells as $cell) {
                $style = $worksheet->getStyle($cell);
                
                if (isset($parameters['bold']) && $parameters['bold']) {
                    $style->getFont()->setBold(true);
                }
                
                if (isset($parameters['italic']) && $parameters['italic']) {
                    $style->getFont()->setItalic(true);
                }
                
                if (isset($parameters['underline']) && $parameters['underline']) {
                    $style->getFont()->setUnderline(true);
                }
                
                if (isset($parameters['color'])) {
                    $style->getFont()->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color($parameters['color']));
                }
                
                if (isset($parameters['size'])) {
                    $style->getFont()->setSize($parameters['size']);
                }
            }
        }
    }
    
    /**
     * Apply sort change to worksheet
     * 
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param array $target Target range
     * @param array $parameters Sort parameters
     */
    private function apply_sort_change($worksheet, $target, $parameters) {
        if (!isset($target['type']) || $target['type'] !== 'range' || !isset($target['reference'])) {
            return;
        }
        
        $range = $target['reference'];
        $column = isset($parameters['column']) ? $parameters['column'] : null;
        $direction = isset($parameters['direction']) ? strtolower($parameters['direction']) : 'ascending';
        
        if (!$column) {
            return;
        }
        
        $direction = ($direction === 'descending') ? 
            \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::SORT_DESC : 
            \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::SORT_ASC;
        
        $columnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($column);
        
        $worksheet->getAutoFilter($range)
            ->setRange($range)
            ->getColumn($column)
            ->setFilterType(\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column::AUTOFILTER_FILTERTYPE_FILTER)
            ->createRule()
            ->setRule(
                \PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_COLUMN_RULE_EQUAL,
                ''
            )
            ->setRuleType(\PhpOffice\PhpSpreadsheet\Worksheet\AutoFilter\Column\Rule::AUTOFILTER_RULETYPE_FILTER);
        
        $worksheet->getAutoFilter()->sort($column, $direction);
    }
    
    /**
     * Apply add column change
     * 
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param array $target Target position
     * @param mixed $value Value for the new column
     * @param array $parameters Additional parameters
     */
    private function apply_add_column_change($worksheet, $target, $value, $parameters) {
        if (!isset($target['reference'])) {
            return;
        }
        
        // Convert column name to index
        $columnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($target['reference']);
        
        // Insert column
        $worksheet->insertNewColumnBefore($target['reference']);
        
        // Set header value if provided
        if (!empty($value)) {
            $worksheet->setCellValueByColumnAndRow($columnIndex, 1, $value);
        }
        
        // Set values if provided in parameters
        if (isset($parameters['values']) && is_array($parameters['values'])) {
            $row = 2; // Start from row 2 (after header)
            foreach ($parameters['values'] as $cellValue) {
                // Handle JSON objects in values
                if (is_string($cellValue) && strpos($cellValue, '{') === 0) {
                    try {
                        $decoded = json_decode($cellValue, true);
                        if (json_last_error() === JSON_ERROR_NONE && is_array($decoded)) {
                            // If JSON parsed successfully, extract relevant value
                            if (isset($decoded[$value])) {
                                $cellValue = $decoded[$value]; // Use value matching column name
                            } else {
                                // Or pick first value if column name doesn't match
                                $cellValue = reset($decoded);
                            }
                        }
                    } catch (Exception $e) {
                        // Keep original value if parsing fails
                    }
                }
                
                $worksheet->setCellValueByColumnAndRow($columnIndex, $row, $cellValue);
                $row++;
            }
        }
    }
    
    /**
     * Apply add row change
     * 
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param array $target Target position
     * @param mixed $value Value for the new row
     * @param array $parameters Additional parameters
     */
    private function apply_add_row_change($worksheet, $target, $value, $parameters) {
        if (!isset($target['reference'])) {
            return;
        }
        
        $rowIndex = (int)$target['reference'];
        
        // Insert row
        $worksheet->insertNewRowBefore($rowIndex);
        
        // Handle values_by_column parameter for better structured data
        if (isset($parameters['values_by_column']) && is_array($parameters['values_by_column'])) {
            $headers = [];
            $headerRow = 1;
            $highestColumn = $worksheet->getHighestColumn();
            $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
            
            // Get column headers
            for ($col = 1; $col <= $highestColumnIndex; $col++) {
                $headers[$col] = $worksheet->getCellByColumnAndRow($col, $headerRow)->getValue();
            }
            
            // Set cell values based on column headers
            foreach ($headers as $colIndex => $header) {
                if (isset($parameters['values_by_column'][$header])) {
                    $worksheet->setCellValueByColumnAndRow($colIndex, $rowIndex, $parameters['values_by_column'][$header]);
                }
            }
        }
        // Set values if provided as array
        else if (isset($parameters['values']) && is_array($parameters['values'])) {
            $column = 1;
            foreach ($parameters['values'] as $cellValue) {
                // Parse JSON-like strings if needed
                if (is_string($cellValue) && strpos($cellValue, '{') === 0) {
                    try {
                        $decoded = json_decode($cellValue, true);
                        if (json_last_error() === JSON_ERROR_NONE && is_array($decoded)) {
                            // If it's a contact object with common fields, extract the appropriate one
                            if (isset($decoded['Name']) && $column === 1) {
                                $cellValue = $decoded['Name'];
                            } else if (isset($decoded['Email']) && $column === 2) {
                                $cellValue = $decoded['Email'];
                            } else if (isset($decoded['Phone']) && $column === 3) {
                                $cellValue = $decoded['Phone'];
                            } else {
                                // Use first value
                                $cellValue = reset($decoded);
                            }
                        }
                    } catch (Exception $e) {
                        // Keep original value if parsing fails
                    }
                }
                
                $worksheet->setCellValueByColumnAndRow($column, $rowIndex, $cellValue);
                $column++;
            }
        } elseif (!empty($value)) {
            // If value is a JSON string, try to parse and extract meaningful values
            if (is_string($value) && strpos($value, '{') === 0) {
                try {
                    $decoded = json_decode($value, true);
                    if (json_last_error() === JSON_ERROR_NONE && is_array($decoded)) {
                        // Get headers to match data to columns
                        $headers = [];
                        $headerRow = 1;
                        $highestColumn = $worksheet->getHighestColumn();
                        $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
                        
                        for ($col = 1; $col <= $highestColumnIndex; $col++) {
                            $headers[$col] = $worksheet->getCellByColumnAndRow($col, $headerRow)->getValue();
                        }
                        
                        // Map decoded values to columns
                        foreach ($headers as $colIndex => $header) {
                            if (isset($decoded[$header])) {
                                $worksheet->setCellValueByColumnAndRow($colIndex, $rowIndex, $decoded[$header]);
                            }
                        }
                        
                        // If no columns were matched, just set the first cell with a meaningful value
                        if (isset($decoded['Name'])) {
                            $worksheet->setCellValueByColumnAndRow(1, $rowIndex, $decoded['Name']);
                        }
                    } else {
                        $worksheet->setCellValueByColumnAndRow(1, $rowIndex, $value);
                    }
                } catch (Exception $e) {
                    $worksheet->setCellValueByColumnAndRow(1, $rowIndex, $value);
                }
            } else {
                // Set the same value in the first cell of the row
                $worksheet->setCellValueByColumnAndRow(1, $rowIndex, $value);
            }
        }
    }
    
    /**
     * Apply delete row change
     * 
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param array $target Target row
     */
    private function apply_delete_row_change($worksheet, $target) {
        if (!isset($target['reference'])) {
            return;
        }
        
        $rowIndex = (int)$target['reference'];
        $worksheet->removeRow($rowIndex);
    }
    
    /**
     * Apply delete column change
     * 
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param array $target Target column
     */
    private function apply_delete_column_change($worksheet, $target) {
        if (!isset($target['reference'])) {
            return;
        }
        
        $worksheet->removeColumn($target['reference']);
    }
    
    /**
     * Set values for a range of cells
     * 
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param string $range Cell range
     * @param mixed $value Value to set
     */
    private function set_range_values($worksheet, $range, $value) {
        $cells = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::extractAllCellReferencesInRange($range);
        
        if (is_array($value) && count($value) > 0) {
            // If value is a 2D array
            if (is_array($value[0])) {
                $row_index = 0;
                foreach ($cells as $cell) {
                    $coordinates = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::coordinateFromString($cell);
                    $col = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($coordinates[0]) - 1;
                    $row = (int)$coordinates[1] - 1;
                    
                    if (isset($value[$row_index][$col])) {
                        // Parse JSON-like strings if needed
                        $cell_value = $value[$row_index][$col];
                        if (is_string($cell_value) && strpos($cell_value, '{') === 0) {
                            try {
                                $decoded = json_decode($cell_value, true);
                                if (json_last_error() === JSON_ERROR_NONE && is_array($decoded)) {
                                    // Get the first value if it's a JSON object
                                    $cell_value = reset($decoded);
                                }
                            } catch (Exception $e) {
                                // Use original value if parsing fails
                            }
                        }
                        $worksheet->getCell($cell)->setValue($cell_value);
                    }
                    
                    $row_index++;
                }
            } 
            // If value is a 1D array
            else {
                $cell_index = 0;
                foreach ($cells as $cell) {
                    if (isset($value[$cell_index])) {
                        // Parse JSON-like strings if needed
                        $cell_value = $value[$cell_index];
                        if (is_string($cell_value) && strpos($cell_value, '{') === 0) {
                            try {
                                $decoded = json_decode($cell_value, true);
                                if (json_last_error() === JSON_ERROR_NONE && is_array($decoded)) {
                                    // Get the first value if it's a JSON object
                                    $cell_value = reset($decoded);
                                }
                            } catch (Exception $e) {
                                // Use original value if parsing fails
                            }
                        }
                        $worksheet->getCell($cell)->setValue($cell_value);
                    }
                    $cell_index++;
                }
            }
        } 
        // If value is a scalar, set it for all cells in the range
        else {
            // Parse JSON-like strings if needed
            $processed_value = $value;
            if (is_string($value) && strpos($value, '{') === 0) {
                try {
                    $decoded = json_decode($value, true);
                    if (json_last_error() === JSON_ERROR_NONE && is_array($decoded)) {
                        // Get the first value if it's a JSON object
                        $processed_value = reset($decoded);
                    }
                } catch (Exception $e) {
                    // Use original value if parsing fails
                }
            }
            
            foreach ($cells as $cell) {
                $worksheet->getCell($cell)->setValue($processed_value);
            }
        }
    }
    
    /**
     * Set values for a column
     * 
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param string $column Column reference
     * @param mixed $value Value to set
     */
    private function set_column_values($worksheet, $column, $value) {
        // Convert column name to index if needed
        $columnIndex = is_numeric($column) ? 
        (int)$column : 
        \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($column);
        
        $highestRow = $worksheet->getHighestRow();
        
        if (is_array($value)) {
            // If value is an array, set each cell to corresponding array value
            $row_index = 1;
            foreach ($value as $cell_value) {
                // Parse JSON-like strings if needed
                if (is_string($cell_value) && strpos($cell_value, '{') === 0) {
                    try {
                        $decoded = json_decode($cell_value, true);
                        if (json_last_error() === JSON_ERROR_NONE && is_array($decoded)) {
                            // Get the first value if it's a JSON object
                            $cell_value = reset($decoded);
                        }
                    } catch (Exception $e) {
                        // Use original value if parsing fails
                    }
                }
                
                if ($row_index <= $highestRow) {
                    $worksheet->setCellValueByColumnAndRow($columnIndex, $row_index, $cell_value);
                }
                $row_index++;
            }
        } else {
            // Parse JSON-like strings if needed
            $processed_value = $value;
            if (is_string($value) && strpos($value, '{') === 0) {
                try {
                    $decoded = json_decode($value, true);
                    if (json_last_error() === JSON_ERROR_NONE && is_array($decoded)) {
                        // Get the first value if it's a JSON object
                        $processed_value = reset($decoded);
                    }
                } catch (Exception $e) {
                    // Use original value if parsing fails
                }
            }
            
            // Set the same value for all cells in the column
            for ($row = 1; $row <= $highestRow; $row++) {
                $worksheet->setCellValueByColumnAndRow($columnIndex, $row, $processed_value);
            }
        }
    }
    
    /**
     * Set values for a row
     * 
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param int $row Row index
     * @param mixed $value Value to set
     */
    private function set_row_values($worksheet, $row, $value) {
        $rowIndex = (int)$row;
        $highestColumn = $worksheet->getHighestColumn();
        $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
        
        if (is_array($value)) {
            // If value is an array, set each cell to corresponding array value
            $col_index = 1;
            foreach ($value as $cell_value) {
                // Parse JSON-like strings if needed
                if (is_string($cell_value) && strpos($cell_value, '{') === 0) {
                    try {
                        $decoded = json_decode($cell_value, true);
                        if (json_last_error() === JSON_ERROR_NONE && is_array($decoded)) {
                            // Get appropriate value based on column
                            if ($col_index === 1 && isset($decoded['Name'])) {
                                $cell_value = $decoded['Name'];
                            } else if ($col_index === 2 && isset($decoded['Email'])) {
                                $cell_value = $decoded['Email'];
                            } else if ($col_index === 3 && isset($decoded['Phone'])) {
                                $cell_value = $decoded['Phone'];
                            } else {
                                // Get the first value if it's a JSON object
                                $cell_value = reset($decoded);
                            }
                        }
                    } catch (Exception $e) {
                        // Use original value if parsing fails
                    }
                }
                
                if ($col_index <= $highestColumnIndex) {
                    $worksheet->setCellValueByColumnAndRow($col_index, $rowIndex, $cell_value);
                }
                $col_index++;
            }
        } else {
            // If value is a JSON-like string
            if (is_string($value) && strpos($value, '{') === 0) {
                try {
                    $decoded = json_decode($value, true);
                    if (json_last_error() === JSON_ERROR_NONE && is_array($decoded)) {
                        // Try to map properties to columns
                        $headers = [];
                        $headerRow = 1;
                        
                        // Get column headers to match values
                        for ($col = 1; $col <= $highestColumnIndex; $col++) {
                            $headers[$col] = $worksheet->getCellByColumnAndRow($col, $headerRow)->getValue();
                        }
                        
                        // Set values matching header names
                        foreach ($headers as $colIndex => $header) {
                            if (isset($decoded[$header])) {
                                $worksheet->setCellValueByColumnAndRow($colIndex, $rowIndex, $decoded[$header]);
                            }
                        }
                        
                        // If we didn't have matching headers for common fields, use defaults
                        if (isset($decoded['Name']) && !in_array('Name', $headers)) {
                            $worksheet->setCellValueByColumnAndRow(1, $rowIndex, $decoded['Name']);
                        }
                        if (isset($decoded['Email']) && !in_array('Email', $headers)) {
                            $worksheet->setCellValueByColumnAndRow(2, $rowIndex, $decoded['Email']);
                        }
                        if (isset($decoded['Phone']) && !in_array('Phone', $headers)) {
                            $worksheet->setCellValueByColumnAndRow(3, $rowIndex, $decoded['Phone']);
                        }
                        
                        return; // We've handled the JSON object
                    }
                } catch (Exception $e) {
                    // Continue with default handling if parsing fails
                }
            }
            
            // Set the same value for all cells in the row
            for ($col = 1; $col <= $highestColumnIndex; $col++) {
                $worksheet->setCellValueByColumnAndRow($col, $rowIndex, $value);
            }
        }
    }
    
    /**
     * Get cell references for a range
     * 
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $worksheet
     * @param array $target Target specification
     * @return array Array of cell references
     */
    private function get_cell_range($worksheet, $target) {
        switch ($target['type']) {
            case 'cell':
                return [$target['reference']];
            
            case 'range':
                return \PhpOffice\PhpSpreadsheet\Cell\Coordinate::extractAllCellReferencesInRange($target['reference']);
            
            case 'column':
                $column = $target['reference'];
                $highestRow = $worksheet->getHighestRow();
                $cells = [];
                
                for ($row = 1; $row <= $highestRow; $row++) {
                    $cells[] = $column . $row;
                }
                
                return $cells;
            
            case 'row':
                $row = (int)$target['reference'];
                $highestColumn = $worksheet->getHighestColumn();
                $cells = [];
                
                $columnRange = range('A', $highestColumn);
                foreach ($columnRange as $column) {
                    $cells[] = $column . $row;
                }
                
                return $cells;
            
            default:
                return [];
        }
    }
    
    /**
     * Get system prompt for OpenAI
     * 
     * @return string System prompt
     */
    private function get_system_prompt() {
        return "You are a spreadsheet assistant that helps modify Excel and CSV files.

1. Always maintain proper data formatting in the spreadsheet.
2. For adding new rows, provide individual values for each cell, not JSON objects.
3. When asked to add people or data, return separate values for each column.
4. DO NOT return JSON structures as values - provide plain text values only.
5. Values should be appropriate for the column they belong in.
6. For names, email addresses, and phone numbers, use realistic but fictional data.

For example, instead of returning:
{\"Name\":\"John Doe\",\"Email\":\"john@example.com\",\"Phone\":\"555-123-4567\"}

Return separate values:
- Name column: John Doe
- Email column: john@example.com
- Phone column: 555-123-4567

This will ensure proper formatting in the spreadsheet.

When adding rows with structured data (like contact information):
1. Use the 'add_row' change type
2. Include a 'parameters' object with 'values_by_column' 
3. Provide individual values for each column 
4. Organize by column headers to ensure proper data placement

Keep all values as plain text strings, not as serialized objects or JSON.

Your task is to analyze the provided spreadsheet data and modify it according to the user's instructions.
The provided spreadsheet data is often a sample of the full data to avoid token limits.
Always respect the structure of the spreadsheet and column names.";
    }
}