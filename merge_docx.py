#!/usr/bin/env python3
"""
merge_docx.py - DOCX file merging utility for combining multiple .docx files.

This module provides functionality to merge multiple Microsoft Word documents (.docx)
into a single combined document while preserving formatting, styles, images, and sections.
"""

import os
from typing import List, Optional
from docx import Document
from docxcompose.composer import Composer
import argparse
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class DocxMerger:
    """A class to handle merging of multiple DOCX files."""
    
    def __init__(self):
        """Initialize the DOCX merger."""
        self.composer = None
        self.master_doc = None
    
    def merge_multiple_docx(self, file_paths: List[str], output_path: str, 
                          remove_blank_pages: bool = True) -> bool:
        """
        Merge multiple DOCX files into a single document.
        
        Args:
            file_paths: List of paths to DOCX files to merge
            output_path: Path where the merged document will be saved
            remove_blank_pages: Whether to remove blank pages between documents
            
        Returns:
            bool: True if merge was successful, False otherwise
        """
        if not file_paths:
            logger.error("No files provided for merging")
            return False
        
        # Validate all files exist
        for file_path in file_paths:
            if not os.path.exists(file_path):
                logger.error(f"File not found: {file_path}")
                return False
        
        try:
            logger.info(f"Merging {len(file_paths)} documents...")
            
            # Start with the first document as master
            master = Document(file_paths[0])
            composer = Composer(master)
            
            # Append remaining documents
            for i, file_path in enumerate(file_paths[1:], 1):
                logger.info(f"Appending document {i+1}/{len(file_paths)}: {os.path.basename(file_path)}")
                doc = Document(file_path)
                composer.append(doc)
            
            # Save the merged document
            composer.save(output_path)
            logger.info(f"Successfully merged documents into: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error during merge: {str(e)}")
            return False
    
    def merge_from_directory(self, directory_path: str, output_path: str, 
                           file_pattern: str = "*.docx", recursive: bool = False) -> bool:
        """
        Merge all DOCX files in a directory.
        
        Args:
            directory_path: Path to directory containing DOCX files
            output_path: Path where the merged document will be saved
            file_pattern: File pattern to match (default: "*.docx")
            recursive: Whether to search recursively in subdirectories
            
        Returns:
            bool: True if merge was successful, False otherwise
        """
        if not os.path.exists(directory_path):
            logger.error(f"Directory not found: {directory_path}")
            return False
        
        # Find all DOCX files
        docx_files = []
        if recursive:
            for root, dirs, files in os.walk(directory_path):
                for file in files:
                    if file.lower().endswith('.docx') and not file.startswith('~'):
                        docx_files.append(os.path.join(root, file))
        else:
            for file in os.listdir(directory_path):
                if file.lower().endswith('.docx') and not file.startswith('~'):
                    docx_files.append(os.path.join(directory_path, file))
        
        if not docx_files:
            logger.error("No DOCX files found in directory")
            return False
        
        # Sort files alphabetically for consistent ordering
        docx_files.sort()
        
        logger.info(f"Found {len(docx_files)} DOCX files to merge")
        return self.merge_multiple_docx(docx_files, output_path)


def main():
    """Command line interface for DOCX merging."""
    parser = argparse.ArgumentParser(description='Merge multiple DOCX files into one')
    parser.add_argument('files', nargs='*', help='DOCX files to merge')
    parser.add_argument('-o', '--output', required=True, help='Output file path')
    parser.add_argument('-d', '--directory', help='Directory containing DOCX files')
    parser.add_argument('-r', '--recursive', action='store_true', 
                       help='Search directories recursively')
    parser.add_argument('-v', '--verbose', action='store_true', 
                       help='Enable verbose logging')
    
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    merger = DocxMerger()
    
    if args.directory:
        success = merger.merge_from_directory(args.directory, args.output, 
                                            recursive=args.recursive)
    elif args.files:
        success = merger.merge_multiple_docx(args.files, args.output)
    else:
        print("Error: Either provide files to merge or use -d for directory")
        return 1
    
    if success:
        print("Merge completed successfully!")
    else:
        print("Merge failed. Check logs for details.")
        return 1


if __name__ == "__main__":
    main()