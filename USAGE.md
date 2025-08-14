# Markdown to DOCX Converter Usage Guide

This project provides a Python-based markdown to DOCX converter with support for tables, headers, lists, and formatting.

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Method 1: Using the Converter Class

```python
from markdown2docx import Markdown2Docx

# Basic usage
project = Markdown2Docx('input.md')
project.eat_soup()
project.save()  # Creates input.docx

# With custom output filename
project = Markdown2Docx('input.md', 'output.docx')
project.eat_soup()
project.save()

# Generate HTML preview
project.save_html_preview()  # Creates input_preview.html
project.save_html_preview('custom_preview.html')
```

### Method 2: Command Line Interface

```bash
# Basic conversion
python markdown2docx.py input.md

# With custom output
python markdown2docx.py input.md -o output.docx

# With HTML preview
python markdown2docx.py input.md --html-preview

# Help
python markdown2docx.py --help
```

### Method 3: Test Script

```bash
# Run the test script with sample content
python test_converter.py
```

## Supported Features

### Markdown Elements
- **Headers**: `#` through `######`
- **Tables**: GitHub Flavored Markdown tables
- **Lists**: Ordered (`1.`) and unordered (`-`, `*`) lists
- **Code blocks**: Fenced code blocks with syntax highlighting
- **Blockquotes**: `>` syntax
- **Text formatting**: Bold (**text**) and italic (*text*)

### Table Support
Tables are supported using GitHub Flavored Markdown syntax:

```markdown
| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Cell 1   | Cell 2   | Cell 3   |
| Cell 4   | Cell 5   | Cell 6   |
```

### Code Blocks
Fenced code blocks are supported:

```markdown
```python
def hello():
    print("Hello, World!")
```
```

## Examples

### Basic Conversion
```python
from markdown2docx import Markdown2Docx

# Convert README.md to README.docx
converter = Markdown2Docx('README.md')
converter.eat_soup()
converter.save()
```

### Advanced Usage
```python
from markdown2docx import Markdown2Docx

# Convert with custom settings
converter = Markdown2Docx(
    input_file='Materi_Pelatihan.md',
    output_file='Training_Materials.docx'
)

# Convert and save
if converter.eat_soup():
    converter.save()
    converter.save_html_preview('preview.html')
```

## Troubleshooting

### Common Issues

1. **Import Error**: Make sure all dependencies are installed
   ```bash
   pip install python-docx markdown beautifulsoup4
   ```

2. **Encoding Issues**: The converter uses UTF-8 encoding by default

3. **Table Formatting**: Ensure tables use proper GitHub Flavored Markdown syntax

### Error Handling

The converter includes comprehensive error handling:
- File existence checks
- UTF-8 encoding support
- Graceful error messages

## File Structure

```
.
├── markdown2docx.py      # Main converter module
├── requirements.txt      # Python dependencies
├── test_converter.py     # Test script with sample content
├── USAGE.md             # This usage guide
├── Materi_Pelatihan.md  # Sample markdown file
└── README.md            # Project documentation
```

## Dependencies

- **python-docx**: For DOCX file creation
- **markdown**: For Markdown parsing
- **beautifulsoup4**: For HTML processing

## License

This is an open-source project for educational and commercial use.