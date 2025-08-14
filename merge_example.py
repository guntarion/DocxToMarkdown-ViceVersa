#!/usr/bin/env python3
"""
merge_example.py - Example showing how to integrate DOCX merging with the existing project

This demonstrates how the new merge functionality can be used alongside
the existing markdown2docx functionality.
"""

import os
from merge_docx import DocxMerger


def merge_markdown_conversions():
    """
    Example: Merge multiple DOCX files created from markdown conversions
    """
    # Example: If you have converted multiple markdown files to DOCX
    docx_files = []
    
    # Look for converted files in the doc directory
    doc_dir = "doc"
    if os.path.exists(doc_dir):
        for file in sorted(os.listdir(doc_dir)):
            if file.endswith('.docx'):
                docx_files.append(os.path.join(doc_dir, file))
    
    if docx_files:
        print(f"Found {len(docx_files)} DOCX files to merge:")
        for f in docx_files:
            print(f"  - {os.path.basename(f)}")
        
        merger = DocxMerger()
        output_file = "merged_document.docx"
        success = merger.merge_multiple_docx(docx_files, output_file)
        
        if success:
            print(f"Successfully merged into: {output_file}")
        else:
            print("Merge failed")
    else:
        print("No DOCX files found in doc directory")


def merge_with_workflow():
    """
    Example: Complete workflow - convert markdown to DOCX then merge
    """
    # This shows how you might integrate with existing functionality
    
    # 1. First, convert markdown files to DOCX (using existing markdown2docx)
    print("Step 1: Convert markdown files to DOCX...")
    # This would use your existing markdown2docx functionality
    
    # 2. Then merge the resulting DOCX files
    print("Step 2: Merge the converted DOCX files...")
    
    merger = DocxMerger()
    
    # Example file list (these would be your actual converted files)
    converted_files = [
        'output/chapter1.docx',
        'output/chapter2.docx',
        'output/chapter3.docx'
    ]
    
    # Filter existing files
    existing_files = [f for f in converted_files if os.path.exists(f)]
    
    if existing_files:
        merger.merge_multiple_docx(existing_files, 'complete_document.docx')
    else:
        print("No converted files found. Create some DOCX files first.")


def batch_merge_example():
    """Example of batch processing multiple sets of documents."""
    
    merger = DocxMerger()
    
    # Example: Merge different sections
    sections = {
        'introduction': ['intro.docx', 'preface.docx'],
        'main_content': ['chapter1.docx', 'chapter2.docx', 'chapter3.docx'],
        'appendices': ['appendix_a.docx', 'appendix_b.docx']
    }
    
    for section_name, files in sections.items():
        # Filter existing files
        existing_files = [f for f in files if os.path.exists(f)]
        
        if existing_files:
            output_file = f"{section_name}_merged.docx"
            merger.merge_multiple_docx(existing_files, output_file)
            print(f"Created {output_file}")


if __name__ == "__main__":
    print("=== DOCX Merge Integration Examples ===")
    
    print("\n1. Merge existing DOCX files:")
    merge_markdown_conversions()
    
    print("\n2. Workflow integration example:")
    merge_with_workflow()
    
    print("\n3. Batch processing example:")
    batch_merge_example()
    
    print("\nDone! Check the usage-merge.md file for detailed documentation.")