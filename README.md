# Excel Processor

A Python script for batch processing Excel files with text processing, data calculations, and feature filtering capabilities.

## Features

- Batch processing of multiple Excel files
- Text cleaning and normalization
- Basic statistical calculations
- Feature filtering based on columns and conditions
- Outputs processed data and statistics to new Excel files

## Requirements

- Python 3.8+
- Packages listed in requirements.txt

## Installation

1. Clone this repository
2. Create a virtual environment:
   ```bash
   python -m venv venv
   ```
3. Activate the virtual environment:
   - Windows: `venv\Scripts\activate`
   - macOS/Linux: `source venv/bin/activate`
4. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Place Excel files in the `input` folder
2. Run the script:
   ```bash
   python process_excel.py
   ```
3. Processed files will be saved in the `output` folder

## License

MIT
