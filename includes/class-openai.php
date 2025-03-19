<?php
/**
 * OpenAI Integration for AI Excel Editor
 */
class AISheets_OpenAI {
    private $api_key;
    private $api_endpoint = 'https://api.openai.com/v1/chat/completions';
    private $model = 'gpt-4';
    
    /**
     * Constructor
     * 
     * @param string $api_key OpenAI API key
     */
    public function __construct($api_key) {
        $this->api_key = $api_key;
    }
    
    /**
     * Process spreadsheet data with OpenAI
     * 
     * @param array $spreadsheet_data Spreadsheet data array
     * @param string $instructions User instructions
     * @return array Changes to apply to spreadsheet
     */
    public function process_spreadsheet($spreadsheet_data, $instructions) {
        try {
            // Prepare sample data to avoid token limits
            $sample_data = $this->prepare_sample_data($spreadsheet_data);
            
            // Log the user instructions and data preparation
            aisheets_debug('Processing spreadsheet with instructions: ' . $instructions);
            aisheets_debug('Sample data prepared with ' . count($sample_data['worksheets']) . ' worksheets');
            
            // TEMPORARY IMPLEMENTATION: Return empty changes to just pass through the file
            aisheets_debug('Using temporary pass-through solution without OpenAI API call');
            
            // Return empty changes array (no modifications to the file)
            return array();
            
            /*
            // UNCOMMENT BELOW FOR FULL IMPLEMENTATION
            
            // Prepare request data
            $request_data = [
                'model' => $this->model,
                'messages' => [
                    [
                        'role' => 'system',
                        'content' => $this->get_system_prompt()
                    ],
                    [
                        'role' => 'user',
                        'content' => "Here is my spreadsheet data:\n" . 
                                     json_encode($sample_data, JSON_PRETTY_PRINT) . 
                                     "\n\nInstructions: " . $instructions
                    ]
                ],
                'functions' => [$this->get_function_definition()],
                'function_call' => ['name' => 'apply_spreadsheet_changes']
            ];
            
            // Log API request (without sensitive data)
            aisheets_debug('Sending request to OpenAI API using model: ' . $this->model);
            
            // Send API request
            $response = $this->send_api_request($request_data);
            
            // Process response
            $result = $this->parse_response($response);
            
            // Log successful API response
            aisheets_debug('Successfully received and parsed OpenAI response');
            
            // Convert the response format to match what AISheets_Spreadsheet expects
            $formatted_changes = $this->format_changes_for_spreadsheet($result);
            
            return $formatted_changes;
            */
            
        } catch (Exception $e) {
            aisheets_debug('OpenAI processing error: ' . $e->getMessage());
            throw new Exception('OpenAI processing error: ' . $e->getMessage());
        }
    }
    
    /**
     * Prepare sample data from full spreadsheet data
     * 
     * @param array $spreadsheet_data Full spreadsheet data
     * @return array Sample data
     */
    private function prepare_sample_data($spreadsheet_data) {
        $sample = ['worksheets' => []];
        
        foreach ($spreadsheet_data['worksheets'] as $sheet_name => $sheet_data) {
            // Check if we have 'rows' or 'data' in the sheet data (handle both formats)
            $rows_key = isset($sheet_data['rows']) ? 'rows' : 'data';
            
            // Take just the first 10 rows to avoid token limits
            $sample_rows = array_slice($sheet_data[$rows_key], 0, 10);
            
            $sample['worksheets'][$sheet_name] = [
                'headers' => $sheet_data['headers'],
                'rows' => $sample_rows,
                'total_rows' => isset($sheet_data['total_rows']) ? $sheet_data['total_rows'] : count($sheet_data[$rows_key]),
                'total_columns' => isset($sheet_data['total_columns']) ? $sheet_data['total_columns'] : count($sheet_data['headers']),
                'is_sample' => count($sheet_data[$rows_key]) > 10,
                'full_row_count' => count($sheet_data[$rows_key])
            ];
        }
        
        return $sample;
    }
    
    /**
     * Get system prompt for OpenAI
     * 
     * @return string System prompt
     */
    private function get_system_prompt() {
        return <<<EOT
You are a spreadsheet expert assistant that helps modify Excel and CSV files based on natural language instructions. Your task is to interpret the user's instructions and return a structured set of changes to apply to their spreadsheet.

Guidelines:
1. Analyze the spreadsheet data structure carefully
2. Determine the specific operations needed based on instructions
3. Return a precisely formatted JSON response using the function calling format
4. Be explicit about cell references, formulas, and values
5. For formulas, use proper Excel syntax (e.g., =SUM(A1:A10))
6. Include a clear explanation of what changes were made

Common operations you can perform:
- Adding/modifying values or formulas
- Adding/removing columns or rows
- Sorting and filtering data
- Formatting cells (numbers, dates, colors)
- Creating calculations and summaries
- Data transformation and analysis

If the instructions are ambiguous, use your best judgment based on the spreadsheet structure and provide an explanation of your reasoning.
EOT;
    }
    
    /**
     * Get function definition for OpenAI
     * 
     * @return array Function definition
     */
    private function get_function_definition() {
        return [
            'name' => 'apply_spreadsheet_changes',
            'description' => 'Apply changes to the spreadsheet based on instructions',
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
                                'change_type' => [
                                    'type' => 'string',
                                    'enum' => ['formula', 'value', 'format', 'sort', 'filter', 'add_column', 'delete_column', 'add_row', 'delete_row'],
                                    'description' => 'Type of change to apply'
                                ],
                                'target' => [
                                    'type' => 'object',
                                    'description' => 'Target of the change (cell, range, column, row)',
                                    'properties' => [
                                        'type' => [
                                            'type' => 'string',
                                            'enum' => ['cell', 'range', 'column', 'row'],
                                            'description' => 'Type of target'
                                        ],
                                        'reference' => [
                                            'type' => 'string',
                                            'description' => 'Cell reference (e.g., A1), range (e.g., A1:C10), column letter/name, or row number/name'
                                        ]
                                    ]
                                ],
                                'value' => [
                                    'type' => 'string',
                                    'description' => 'New value, formula, or format to apply'
                                ],
                                'parameters' => [
                                    'type' => 'object',
                                    'description' => 'Additional parameters for the change',
                                ]
                            ],
                            'required' => ['worksheet', 'change_type', 'target']
                        ]
                    ],
                    'explanation' => [
                        'type' => 'string',
                        'description' => 'Explanation of changes made to the spreadsheet'
                    ]
                ],
                'required' => ['changes', 'explanation']
            ]
        ];
    }
    
    /**
     * Send request to OpenAI API
     * 
     * @param array $request_data Request data
     * @return array API response
     */
    private function send_api_request($request_data) {
        $ch = curl_init($this->api_endpoint);
        
        $headers = [
            'Content-Type: application/json',
            'Authorization: Bearer ' . $this->api_key
        ];
        
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        curl_setopt($ch, CURLOPT_POST, true);
        curl_setopt($ch, CURLOPT_POSTFIELDS, json_encode($request_data));
        curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
        curl_setopt($ch, CURLOPT_TIMEOUT, 120); // Increased timeout to 120 seconds for larger spreadsheets
        
        $response = curl_exec($ch);
        $http_code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
        
        if (curl_errno($ch)) {
            $error = curl_error($ch);
            curl_close($ch);
            throw new Exception('API request failed: ' . $error);
        }
        
        curl_close($ch);
        
        if ($http_code !== 200) {
            $error_details = json_decode($response, true);
            $error_message = isset($error_details['error']['message']) 
                ? $error_details['error']['message'] 
                : 'API request failed with status code ' . $http_code;
            
            throw new Exception($error_message);
        }
        
        $decoded_response = json_decode($response, true);
        
        if (json_last_error() !== JSON_ERROR_NONE) {
            throw new Exception('Failed to parse API response: ' . json_last_error_msg());
        }
        
        return $decoded_response;
    }
    
    /**
     * Parse OpenAI API response
     * 
     * @param array $response API response
     * @return array Parsed changes
     */
    private function parse_response($response) {
        // Check for function call
        if (!isset($response['choices'][0]['message']['function_call'])) {
            throw new Exception('Invalid response format from OpenAI API');
        }
        
        $function_call = $response['choices'][0]['message']['function_call'];
        
        if ($function_call['name'] !== 'apply_spreadsheet_changes') {
            throw new Exception('Unexpected function call: ' . $function_call['name']);
        }
        
        // Parse arguments
        $arguments = json_decode($function_call['arguments'], true);
        
        if (json_last_error() !== JSON_ERROR_NONE) {
            throw new Exception('Failed to parse function arguments: ' . json_last_error_msg());
        }
        
        if (!isset($arguments['changes']) || !is_array($arguments['changes'])) {
            throw new Exception('Missing or invalid changes in response');
        }
        
        // Store explanation for user feedback
        if (isset($arguments['explanation'])) {
            update_option('ai_excel_editor_last_explanation', $arguments['explanation']);
        }
        
        return $arguments;
    }
    
    /**
     * Format changes to match what the spreadsheet handler expects
     * 
     * @param array $result OpenAI response result
     * @return array Formatted changes
     */
    private function format_changes_for_spreadsheet($result) {
        $formatted_changes = [];
        
        // Group changes by worksheet
        foreach ($result['changes'] as $change) {
            $worksheet_name = $change['worksheet'];
            
            if (!isset($formatted_changes[$worksheet_name])) {
                $formatted_changes[$worksheet_name] = [
                    'cell_changes' => [],
                    'column_changes' => [],
                    'row_changes' => [],
                    'format_changes' => [],
                    'sort_changes' => null
                ];
            }
            
            // Process change based on type
            switch ($change['change_type']) {
                case 'formula':
                case 'value':
                    if ($change['target']['type'] === 'cell') {
                        $cell_ref = $change['target']['reference'];
                        $value = isset($change['value']) ? $change['value'] : '';
                        
                        if (strpos($value, '=') === 0) {
                            $formatted_changes[$worksheet_name]['cell_changes'][$cell_ref] = ['formula' => $value];
                        } else {
                            $formatted_changes[$worksheet_name]['cell_changes'][$cell_ref] = ['value' => $value];
                        }
                    }
                    break;
                    
                case 'add_column':
                    if ($change['target']['type'] === 'column') {
                        $column_ref = $change['target']['reference'];
                        $header = isset($change['parameters']['header']) ? $change['parameters']['header'] : '';
                        $values = isset($change['parameters']['values']) ? $change['parameters']['values'] : [];
                        $formula = isset($change['value']) && strpos($change['value'], '=') === 0 ? $change['value'] : null;
                        
                        $formatted_changes[$worksheet_name]['column_changes'][$column_ref] = [
                            'operation' => 'add',
                            'header' => $header,
                            'values' => $values
                        ];
                        
                        if ($formula) {
                            $formatted_changes[$worksheet_name]['column_changes'][$column_ref]['formula'] = $formula;
                        }
                    }
                    break;
                    
                case 'delete_column':
                    if ($change['target']['type'] === 'column') {
                        $column_ref = $change['target']['reference'];
                        $formatted_changes[$worksheet_name]['column_changes'][$column_ref] = [
                            'operation' => 'delete'
                        ];
                    }
                    break;
                    
                case 'add_row':
                    if ($change['target']['type'] === 'row') {
                        $row_ref = $change['target']['reference'];
                        $values = isset($change['parameters']['values']) ? $change['parameters']['values'] : [];
                        
                        $formatted_changes[$worksheet_name]['row_changes'][$row_ref] = [
                            'operation' => 'add',
                            'values' => $values
                        ];
                    }
                    break;
                    
                case 'delete_row':
                    if ($change['target']['type'] === 'row') {
                        $row_ref = $change['target']['reference'];
                        $formatted_changes[$worksheet_name]['row_changes'][$row_ref] = [
                            'operation' => 'delete'
                        ];
                    }
                    break;
                    
                case 'format':
                    if (in_array($change['target']['type'], ['cell', 'range'])) {
                        $range_ref = $change['target']['reference'];
                        $format_params = isset($change['parameters']) ? $change['parameters'] : [];
                        
                        $formatted_changes[$worksheet_name]['format_changes'][$range_ref] = $format_params;
                    }
                    break;
                    
                case 'sort':
                    if ($change['target']['type'] === 'range') {
                        $range = $change['target']['reference'];
                        $column = isset($change['parameters']['column']) ? $change['parameters']['column'] : null;
                        $direction = isset($change['parameters']['direction']) ? $change['parameters']['direction'] : 'ASC';
                        
                        $formatted_changes[$worksheet_name]['sort_changes'] = [
                            'range' => $range,
                            'column' => $column,
                            'direction' => $direction
                        ];
                    }
                    break;
            }
        }
        
        return $formatted_changes;
    }
}