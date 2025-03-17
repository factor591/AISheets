# AISheets

![AISheets Logo](ai-excel-editor/img/ai-icon.svg)

AISheets is a WordPress plugin that leverages artificial intelligence to transform spreadsheet editing. Upload Excel or CSV files, describe changes in plain English, and let AI do the work for you - no complex formulas required!

## Features

- **Natural Language Spreadsheet Editing**: Describe what you want to do in plain English
- **Support for Multiple Formats**: Works with Excel (.xlsx, .xls) and CSV files
- **Secure File Handling**: Files are processed securely and automatically cleaned up
- **User-Friendly Interface**: Simple drag-and-drop upload and intuitive controls
- **Instant Downloads**: Processed files are available for immediate download

## How It Works

1. **Upload** your Excel or CSV file through the intuitive drag-and-drop interface
2. **Describe** the changes you want using natural language instructions
3. **Process** your file with the power of AI
4. **Download** your transformed spreadsheet instantly

## Examples of What You Can Do

### Calculations
- "Calculate the total for each row and add it as a new column"
- "Find the average of column C and add it to cell C20"
- "Calculate profit margin (Revenue - Cost)/Revenue as a percentage"

### Formatting
- "Format all values in column D as currency with $ symbol"
- "Add red background to all negative numbers"
- "Make the header row bold and centered"

### Data Manipulation
- "Sort the data by the Date column in descending order"
- "Filter out rows where the Status is 'Cancelled'"
- "Remove duplicate entries based on the Email column"

## Installation

### Requirements
- WordPress 5.6 or higher
- PHP 7.4 or higher
- OpenAI API key

### Installation Steps
1. Download the plugin zip file
2. Upload to your WordPress site via the plugin uploader
3. Activate the plugin through the 'Plugins' menu in WordPress
4. Go to AISheets settings and enter your OpenAI API key

## Usage

1. Create a new page or post
2. Add the shortcode `[ai_excel_editor]` where you want the editor to appear
3. Save the page and visit it to use the AISheets editor
4. Upload a file, enter instructions, and click "Process & Download"

## Configuration

### API Key Setup
1. Go to WordPress Admin → AISheets
2. Enter your OpenAI API key in the provided field
3. Save changes

### File Limitations
- Maximum file size: 5MB
- Supported formats: .xlsx, .xls, .csv

## Development Roadmap

- [x] Basic file upload and validation
- [x] User interface with drag-and-drop
- [x] Secure file handling system
- [x] PhpSpreadsheet integration
- [ ] Advanced AI processing capabilities
- [ ] Preview functionality
- [ ] Support for larger files
- [ ] User accounts and saved preferences

## Security

AISheets takes security seriously:
- Files are stored in protected directories
- Automatic file cleanup runs hourly
- Strict file validation prevents malicious uploads
- API keys are stored securely in WordPress options

## For Developers

The plugin is built with extensibility in mind:
- Clear class structure
- PhpSpreadsheet for file manipulation
- OpenAI API for natural language processing
- WordPress standards compliant

## License

This project is licensed under the GPL v2 or later - see the LICENSE file for details.

## Credits

- [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) for Excel file handling
- [OpenAI](https://openai.com/) for AI processing capabilities
- Icons and interface components by [Your Name]