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
            
            // Send API request
            $response = $this->send_api_request($request_data);
            
            // Process response
            return $this->parse_response($response);
            
        } catch (Exception $e) {
            error_log('OpenAI processing error: ' . $e->getMessage());
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
            // Take just the first 10 rows to avoid token limits
            $sample_rows = array_slice($sheet_data['rows'], 0, 10);
            
            $sample['worksheets'][$sheet_name] = [
                'headers' => $sheet_data['headers'],
                'rows' => $sample_rows,
                'rowCount' => $sheet_data['rowCount'],
                'columnCount' => $sheet_data['columnCount'],
                'is_sample' => count($sheet_data['rows']) > 10,
                'total_rows' => count($sheet_data['rows'])
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
        curl_setopt($ch, CURLOPT_TIMEOUT, 60); // 60-second timeout
        
        $response = curl_exec($ch);
        $http_code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
        
        if (curl_errno($ch)) {
            $error = curl_error($ch);
            curl_close($ch);
            throw new Exception('API request failed: ' . $error);
        }
        
        curl_close($ch);
        
        if ($http_code !== 200) {
            throw new Exception('API request failed with status code ' . $http_code . ': ' . $response);
        }
        
        return json_decode($response, true);
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
        
        return $arguments;
    }
}