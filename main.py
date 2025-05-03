#!/usr/bin/env python3
"""
dso-document-revision-history-remover.py

Removes revision history from document files (e.g., .docx, .odt) to anonymize
the evolution of the document.

Focused on Tools for sanitizing and obfuscating sensitive data within text
files and structured data formats.
"""

import argparse
import logging
import os
import sys
import zipfile
import xml.etree.ElementTree as ET
from tempfile import TemporaryDirectory
import shutil
import chardet

# Configure logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Define a custom exception
class RevisionHistoryRemovalError(Exception):
    """Custom exception for revision history removal errors."""
    pass


def setup_argparse():
    """Sets up the argument parser for the command-line interface."""
    parser = argparse.ArgumentParser(
        description="Removes revision history from document files (docx, odt)."
    )
    parser.add_argument(
        "input_file",
        help="Path to the input document file (docx or odt)."
    )
    parser.add_argument(
        "output_file",
        help="Path to the output document file (will overwrite if exists)."
    )
    parser.add_argument(
        "--log_level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        help="Set the logging level (default: INFO)."
    )
    return parser


def remove_docx_revision_history(input_file, output_file):
    """
    Removes revision history from a .docx file.
    Args:
        input_file (str): Path to the input .docx file.
        output_file (str): Path to the output .docx file.

    Raises:
        RevisionHistoryRemovalError: If an error occurs during processing.
    """
    try:
        # Create a temporary directory to extract the contents of the docx file.
        with TemporaryDirectory() as tmp_dir:
            # Extract the docx file.
            with zipfile.ZipFile(input_file, 'r') as zip_ref:
                zip_ref.extractall(tmp_dir)

            # Identify revision files and remove them.
            revision_files = [
                f for f in os.listdir(tmp_dir + '/word')
                if f.startswith('revisions')
            ]
            for rev_file in revision_files:
                os.remove(os.path.join(tmp_dir, 'word', rev_file))
                logging.debug(f"Removed revision file: {rev_file}")

            # Modify document.xml to remove track revisions.
            document_xml_path = os.path.join(tmp_dir, 'word', 'document.xml')
            if os.path.exists(document_xml_path):
                try:
                    with open(document_xml_path, 'rb') as f:
                        raw_data = f.read()
                        result = chardet.detect(raw_data)
                        encoding = result['encoding']
                        if encoding is None:
                            encoding = 'utf-8'
                        doc_xml = raw_data.decode(encoding)
                    root = ET.fromstring(doc_xml)

                    # Namespace mapping
                    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

                    # Remove all 'ins' and 'del' elements (track changes)
                    for ins in root.findall('.//w:ins', ns):
                        parent = ins.getparent()
                        parent.remove(ins)
                    for dele in root.findall('.//w:del', ns):
                         parent = dele.getparent()
                         parent.remove(dele)
                   
                    xmlstr = ET.tostring(root, encoding=encoding).decode(encoding)
                    with open(document_xml_path, 'w', encoding=encoding) as f:
                        f.write(xmlstr)


                except ET.ParseError as e:
                    raise RevisionHistoryRemovalError(f"Error parsing document.xml: {e}")
                except Exception as e:
                     raise RevisionHistoryRemovalError(f"Error processing document.xml: {e}")



            # Create a new docx file without revision history.
            with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                for root, _, files in os.walk(tmp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        archive_path = os.path.relpath(file_path, tmp_dir)  # Relative path within the archive
                        zip_out.write(file_path, archive_path)  # Write to the output zip file

    except FileNotFoundError:
        raise RevisionHistoryRemovalError(f"Input file not found: {input_file}")
    except zipfile.BadZipFile:
        raise RevisionHistoryRemovalError(f"Invalid zip file: {input_file}")
    except OSError as e:
        raise RevisionHistoryRemovalError(f"OS error: {e}")
    except Exception as e:
        raise RevisionHistoryRemovalError(f"An unexpected error occurred: {e}")


def remove_odt_revision_history(input_file, output_file):
    """
    Removes revision history from an .odt file.

    Args:
        input_file (str): Path to the input .odt file.
        output_file (str): Path to the output .odt file.

    Raises:
        RevisionHistoryRemovalError: If an error occurs during processing.
    """
    try:
        # Create a temporary directory
        with TemporaryDirectory() as tmp_dir:
            # Extract the odt file
            with zipfile.ZipFile(input_file, 'r') as zip_ref:
                zip_ref.extractall(tmp_dir)

            # Remove meta.xml, settings.xml, and styles.xml (common places for metadata).  This is an aggressive approach.
            for file in ['meta.xml', 'settings.xml', 'styles.xml']:
                file_path = os.path.join(tmp_dir, file)
                if os.path.exists(file_path):
                    os.remove(file_path)
                    logging.debug(f"Removed file: {file}")


            # Modify content.xml to remove track revisions and other metadata.
            content_xml_path = os.path.join(tmp_dir, 'content.xml')
            if os.path.exists(content_xml_path):
                try:
                    with open(content_xml_path, 'rb') as f:
                        raw_data = f.read()
                        result = chardet.detect(raw_data)
                        encoding = result['encoding']
                        if encoding is None:
                            encoding = 'utf-8'
                        content_xml = raw_data.decode(encoding)
                    root = ET.fromstring(content_xml)
                    ns = {'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
                          'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'}


                    # Remove track changes.  This is ODT specific logic.  It may not catch everything.
                    for change in root.findall('.//text:tracked-changes', ns):
                        parent = change.getparent()
                        parent.remove(change)
                    
                    xmlstr = ET.tostring(root, encoding=encoding).decode(encoding)
                    with open(content_xml_path, 'w', encoding=encoding) as f:
                        f.write(xmlstr)


                except ET.ParseError as e:
                    raise RevisionHistoryRemovalError(f"Error parsing content.xml: {e}")
                except Exception as e:
                    raise RevisionHistoryRemovalError(f"Error processing content.xml: {e}")

            # Recreate the odt file
            with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                for root, _, files in os.walk(tmp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        archive_path = os.path.relpath(file_path, tmp_dir)
                        zip_out.write(file_path, archive_path)

    except FileNotFoundError:
        raise RevisionHistoryRemovalError(f"Input file not found: {input_file}")
    except zipfile.BadZipFile:
        raise RevisionHistoryRemovalError(f"Invalid zip file: {input_file}")
    except OSError as e:
        raise RevisionHistoryRemovalError(f"OS error: {e}")
    except Exception as e:
        raise RevisionHistoryRemovalError(f"An unexpected error occurred: {e}")


def main():
    """
    Main function to parse arguments and process the document.
    """
    parser = setup_argparse()
    args = parser.parse_args()

    # Set logging level
    logging.getLogger().setLevel(args.log_level.upper())

    input_file = args.input_file
    output_file = args.output_file

    # Validate input file exists
    if not os.path.exists(input_file):
        logging.error(f"Input file does not exist: {input_file}")
        sys.exit(1)

    # Determine file type and call appropriate function
    try:
        if input_file.endswith('.docx'):
            remove_docx_revision_history(input_file, output_file)
            logging.info(f"Revision history removed from {input_file}, output saved to {output_file}")
        elif input_file.endswith('.odt'):
            remove_odt_revision_history(input_file, output_file)
            logging.info(f"Revision history removed from {input_file}, output saved to {output_file}")
        else:
            logging.error("Unsupported file type. Only .docx and .odt are supported.")
            sys.exit(1)
    except RevisionHistoryRemovalError as e:
        logging.error(f"Error removing revision history: {e}")
        sys.exit(1)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()

# Example usage:
# ./dso-document-revision-history-remover.py input.docx output.docx
# ./dso-document-revision-history-remover.py input.odt output.odt --log_level DEBUG