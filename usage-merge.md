# DOCX File Merging Guide

This guide explains how to use the DOCX merging functionality to combine multiple Word documents into a single file.

## Overview

The `merge_docx.py` module provides functionality to merge multiple Microsoft Word documents (.docx) into a single combined document while preserving formatting, styles, images, and sections.

## Installation

### Required Dependencies

Make sure you have the required dependencies installed:

```bash
pip install python-docx docxcompose
```

Or install from requirements.txt:
```bash
pip install -r requirements.txt
```

## Usage Methods

### Method 1: Using the DocxMerger Class

```python
from merge_docx import DocxMerger

# Create merger instance
merger = DocxMerger()

# Merge specific files
files_to_merge = [
    'document1.docx',
    'document2.docx',
    'document3.docx'
]
success = merger.merge_multiple_docx(files_to_merge, 'merged_output.docx')

if success:
    print("Files merged successfully!")
```

### Method 2: Merge from Directory

```python
from merge_docx import DocxMerger

# Create merger instance
merger = DocxMerger()

# Merge all DOCX files in a directory
success = merger.merge_from_directory(
    '/path/to/documents', 
    'merged_output.docx',
    recursive=True  # Include subdirectories
)

if success:
    print("Directory merge completed!")
```

### Method 3: Command Line Interface

#### Basic Usage
```bash
# Merge specific files
python merge_docx.py file1.docx file2.docx file3.docx -o merged.docx

# Merge from directory
python merge_docx.py -d /path/to/documents -o merged.docx

# With verbose logging
python merge_docx.py file1.docx file2.docx -o merged.docx -v

# Recursive directory merge
python merge_docx.py -d /path/to/documents -o merged.docx -r
```

## Examples

### Example 1: Merging Specific Files
```bash
python merge_docx.py \
  "Chapter 1.docx" \
  "Chapter 2.docx" \
  "Chapter 3.docx" \
  -o "Complete Book.docx"
```

### Example 2: Merging All Files in Current Directory
```bash
python merge_docx.py -d . -o "combined.docx"
```

### Example 3: Python Script Integration
```python
import os
from merge_docx import DocxMerger

# Merge chapters for a book
chapters = [f"Chapter {i}.docx" for i in range(1, 6)]
output_file = "Complete_Book.docx"

merger = DocxMerger()
success = merger.merge_multiple_docx(chapters, output_file)

if not success:
    print("Error merging files")
```

## Advanced Usage

### Handling File Paths
The merger accepts both absolute and relative file paths:

```python
files = [
    '/Users/user/Documents/report1.docx',
    './reports/report2.docx',
    '../archive/report3.docx'
]
```

### Error Handling
```python
from merge_docx import DocxMerger

merger = DocxMerger()

try:
    success = merger.merge_multiple_docx(files, 'output.docx')
    if not success:
        print("Merge failed - check logs")
except Exception as e:
    print(f"Unexpected error: {e}")
```

## Features

- **Preserves Formatting**: Maintains styles, fonts, and formatting from all source documents
- **Handles Images**: Includes images from all documents
- **Maintains Sections**: Preserves page breaks and section formatting
- **Flexible Input**: Accepts specific file lists or directory-based merging
- **Recursive Support**: Can merge files from subdirectories
- **Error Handling**: Comprehensive validation and error reporting
- **Logging**: Detailed logging for debugging and monitoring

## Limitations

- Only works with `.docx` files (not `.doc` format)
- All files must exist and be readable
- Merged document will have the page setup of the first document
- Header and footer styles follow the first document's settings

## Troubleshooting

### Common Issues

**Issue**: "File not found" error
- **Solution**: Check that all file paths are correct and files exist

**Issue**: "No DOCX files found" error
- **Solution**: Verify the directory contains `.docx` files (not `.doc`)

**Issue**: Formatting issues in merged document
- **Solution**: Ensure all source documents use compatible styles

### Debug Mode
Enable verbose logging for detailed information:
```bash
python merge_docx.py -v -d . -o output.docx
```

## Testing

### Quick Test
1. Create test DOCX files with different content
2. Run: `python merge_docx.py test1.docx test2.docx -o result.docx`
3. Open `result.docx` to verify the merge

### Batch Testing
```python
# test_merge.py
from merge_docx import DocxMerger
import tempfile
import os

# Create test files
with tempfile.TemporaryDirectory() as tmpdir:
    # Your test code here
    pass