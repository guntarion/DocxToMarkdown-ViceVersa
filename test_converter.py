#!/usr/bin/env python3
"""
Test script for markdown2docx converter
"""

from markdown2docx import Markdown2Docx
import os

def test_conversion():
    """Test the markdown to docx conversion"""
    
    # Check if Materi_Pelatihan.md exists
    test_file = "Materi_Pelatihan.md"
    
    if not os.path.isfile(test_file):
        print(f"Test file '{test_file}' not found. Creating a sample markdown file...")
        
        # Create a sample markdown file
        sample_content = """# Training Materials

## Introduction
This document contains training materials for our technical workshop.

## Table of Contents
- [Introduction](#introduction)
- [Technical Content](#technical-content)
- [Examples](#examples)

## Technical Content

### Python Basics
Python is a powerful programming language with the following features:

| Feature | Description | Example |
|---------|-------------|---------|
| Simple | Easy to read and write | `print("Hello World")` |
| Powerful | Can handle complex tasks | Data analysis, web dev |
| Popular | Widely used in industry | Google, Netflix, NASA |

### Code Examples

Here's a simple Python function:

```python
def greet(name):
    \"\"\"Greet a person by name\"\"\"
    return f"Hello, {name}!"

# Usage
print(greet("World"))
```

### Lists and Formatting

- **Bold text** and *italic text*
- Numbered items:
  1. First item
  2. Second item
  3. Third item

### Blockquotes

> "The best way to predict the future is to invent it."
> - Alan Kay

## Conclusion
This concludes our basic training materials."""
        
        with open(test_file, 'w', encoding='utf-8') as f:
            f.write(sample_content)
        print(f"Created sample markdown file: {test_file}")
    
    # Convert the markdown file to docx
    print(f"Converting {test_file} to DOCX...")
    
    converter = Markdown2Docx(test_file)
    
    if converter.eat_soup():
        if converter.save():
            print("✅ Conversion completed successfully!")
            print(f"📄 Output file: {converter.output_file}")
        else:
            print("❌ Failed to save the document")
    else:
        print("❌ Failed to convert the markdown file")
    
    # Also create HTML preview
    print("\nGenerating HTML preview...")
    if converter.save_html_preview():
        print("✅ HTML preview generated successfully!")
    else:
        print("❌ Failed to generate HTML preview")

if __name__ == "__main__":
    test_conversion()