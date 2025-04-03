# ASCII Value Calculator with PowerPoint Visualization

## Overview
This project demonstrates an interactive system that calculates the sum of exponentials of ASCII values for a given string and visualizes the result in Microsoft PowerPoint. It uses a combination of Python scripts, MCP (Model-Client-Protocol) server, and PowerPoint automation.

## Features
- Converts text to ASCII values
- Calculates exponential sums
- Automated PowerPoint presentation creation
- Visual representation with:
  - Centered rectangle
  - Formatted text display
  - Clean presentation layout

## Project Structure
```
.
├── talk2mcp-2.py      # Main client script with LLM integration
├── example2-3.py      # MCP server implementation with tools
└── README.md         # Project documentation
```

## Prerequisites
- Python 3.x
- Microsoft PowerPoint
- Required Python packages:
  ```
  python-dotenv
  mcp
  google-genai
  python-pptx
  pywinauto
  ```
- Gemini API key (stored in .env file)

## Setup
1. Clone the repository
2. Install required packages:
   ```bash
   pip install python-dotenv mcp google-genai python-pptx pywinauto
   ```
3. Create a `.env` file with your Gemini API key:
   ```
   GEMINI_API_KEY=your_api_key_here
   ```

## Usage
1. Run the main script:
   ```bash
   python talk2mcp-2.py
   ```
2. The program will:
   - Convert "INDIA" to ASCII values
   - Calculate exponential sums
   - Create a PowerPoint presentation
   - Display the result in a formatted rectangle
   - Close PowerPoint automatically

## Process Flow
1. Text to ASCII Conversion
   - Input text is converted to ASCII values
   - Example: "INDIA" → [73, 78, 68, 73, 65]

2. Mathematical Calculation
   - Calculates exponential sum of ASCII values
   - Uses math.exp() for each value

3. PowerPoint Visualization
   - Opens PowerPoint
   - Creates centered rectangle
   - Displays formatted result
   - Auto-closes after completion

## Functions
### Main Tools
- `strings_to_chars_to_int`: Converts string to ASCII values
- `int_list_to_exponential_sum`: Calculates sum of exponentials
- `draw_rectangle`: Creates visual rectangle in PowerPoint
- `add_text_in_powerpoint`: Adds formatted text to presentation

### PowerPoint Operations
- `open_powerpoint()`: Creates new presentation
- `draw_rectangle(x1, y1, x2, y2)`: Draws rectangle with specified coordinates
- `add_text_in_powerpoint(text)`: Adds formatted text
- `close_powerpoint()`: Closes presentation

## Error Handling
- Robust parameter validation
- Proper error messages
- Automatic cleanup on interruption
- Retry mechanism for server connection

## Notes
- Rectangle coordinates are limited to 1-8 range
- Text is automatically centered and formatted
- Program exits cleanly after PowerPoint closes
- Backup files (.bak) are maintained for safety

## Troubleshooting
If you encounter issues:
1. Ensure PowerPoint is installed and accessible
2. Check Gemini API key in .env file
3. Verify all required packages are installed
4. Ensure no existing PowerPoint instances are running

## License
This project is provided as-is for educational purposes. 