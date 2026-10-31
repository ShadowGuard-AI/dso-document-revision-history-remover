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

# Register the common prefixes so rewritten parts keep readable names.
for _prefix, _uri in [
    ('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'),
    ('cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'),
    ('dc', 'http://purl.org/dc/elements/1.1/'),
    ('office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'),
    ('text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'),
    ('meta', 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0'),
    ('style', 'urn:oasis:names:tc:opendocument:xmlns:style:1.0'),
]:
    ET.register_namespace(_prefix, _uri)


# Define a custom exception
class RevisionHistoryRemovalError(Exception):
    """Custom exception for revision history removal errors."""
    pass


def _parent_map(root):
    """Map each element to its parent (ElementTree has no getparent)."""
    return {child: parent for parent in root.iter() for child in parent}


def _remove_element(root, element):
    parent = _parent_map(root).get(element)
    if parent is None:
        return False
    if element.tail:
        # Text after an element is stored in its tail; keep it by moving
        # it onto the previous sibling or the parent, or it would be lost
        # with the element (this is where ODT body text around change
        # markers lives).
        index = list(parent).index(element)
        if index > 0:
            previous = parent[index - 1]
            previous.tail = (previous.tail or '') + element.tail
        else:
            parent.text = (parent.text or '') + element.tail
    parent.remove(element)
    return True


def _unwrap_element(root, element):
    """Replace an element with its children, keeping their text and order."""
    parent = _parent_map(root).get(element)
    if parent is None:
        return False
    index = list(parent).index(element)
    children = list(element)
    parent.remove(element)
    for offset, child in enumerate(children):
        parent.insert(index + offset, child)
    if element.tail:
        if children:
            last = children[-1]
            last.tail = (last.tail or '') + element.tail
        else:
            previous = parent[index - 1] if index > 0 else None
            if previous is not None:
                previous.tail = (previous.tail or '') + element.tail
            else:
                parent.text = (parent.text or '') + element.tail
    elif not children and element.text:
        parent.text = (parent.text or '') + element.text
    return True


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
    removed = 0
    try:
        # Create a temporary directory to extract the contents of the docx file.
        with TemporaryDirectory() as tmp_dir:
            # Extract the docx file.
            with zipfile.ZipFile(input_file, 'r') as zip_ref:
                zip_ref.extractall(tmp_dir)

            # Identify revision files and remove them.
            word_dir = os.path.join(tmp_dir, 'word')
            revision_files = [
                f for f in os.listdir(word_dir)
                if f.startswith('revisions')
            ] if os.path.isdir(word_dir) else []
            for rev_file in revision_files:
                os.remove(os.path.join(tmp_dir, 'word', rev_file))
                removed += 1
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

                    # Accept tracked changes. A deletion is dropped with
                    # its deleted text; an insertion is unwrapped so the
                    # inserted text stays as normal content. ElementTree
                    # elements have no getparent, so parent maps are built.
                    for dele in root.findall('.//w:del', ns):
                        if _remove_element(root, dele):
                            removed += 1
                    for move_from in root.findall('.//w:moveFrom', ns):
                        if _remove_element(root, move_from):
                            removed += 1
                    for tag in ('.//w:ins', './/w:moveTo'):
                        while True:
                            found = root.findall(tag, ns)
                            if not found:
                                break
                            if _unwrap_element(root, found[0]):
                                removed += 1
                            else:
                                break

                    xmlstr = ET.tostring(root, encoding=encoding).decode(encoding)
                    with open(document_xml_path, 'w', encoding=encoding) as f:
                        f.write(xmlstr)


                except ET.ParseError as e:
                    raise RevisionHistoryRemovalError(f"Error parsing document.xml: {e}")
                except Exception as e:
                     raise RevisionHistoryRemovalError(f"Error processing document.xml: {e}")



            # Clear the revision metadata in docProps/core.xml: the
            # creator, the last modifier, and the revision count.
            core_xml_path = os.path.join(tmp_dir, 'docProps', 'core.xml')
            if os.path.exists(core_xml_path):
                try:
                    core_root = ET.parse(core_xml_path).getroot()
                    core_ns = {
                        'dc': 'http://purl.org/dc/elements/1.1/',
                        'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                    }
                    for tag in ('dc:creator', 'cp:lastModifiedBy'):
                        for el in core_root.findall(tag, core_ns):
                            if el.text:
                                el.text = None
                                removed += 1
                    for el in core_root.findall('cp:revision', core_ns):
                        if el.text and el.text.strip() not in ('', '0'):
                            removed += 1
                        core_root.remove(el)
                    ET.ElementTree(core_root).write(
                        core_xml_path, encoding='UTF-8', xml_declaration=True)
                except ET.ParseError as e:
                    raise RevisionHistoryRemovalError(f"Error parsing docProps/core.xml: {e}")

            # Create a new docx file without revision history.
            with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                for root, _, files in os.walk(tmp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        archive_path = os.path.relpath(file_path, tmp_dir)  # Relative path within the archive
                        zip_out.write(file_path, archive_path)  # Write to the output zip file
        return removed

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
    removed = 0
    try:
        # Create a temporary directory
        with TemporaryDirectory() as tmp_dir:
            # Extract the odt file
            with zipfile.ZipFile(input_file, 'r') as zip_ref:
                zip_ref.extractall(tmp_dir)

            # Keep styles.xml and every other part. Only the tracked
            # changes in content.xml and the revision fields in meta.xml
            # are touched. In meta.xml, clear the creator fields and the
            # editing cycle count instead of deleting the part.
            meta_xml_path = os.path.join(tmp_dir, 'meta.xml')
            if os.path.exists(meta_xml_path):
                try:
                    meta_root = ET.parse(meta_xml_path).getroot()
                    meta_ns = {
                        'dc': 'http://purl.org/dc/elements/1.1/',
                        'meta': 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
                    }
                    for tag in ('meta:initial-creator', 'dc:creator'):
                        for el in meta_root.findall(f'.//{tag}', meta_ns):
                            if el.text:
                                el.text = None
                                removed += 1
                    for el in meta_root.findall('.//meta:editing-cycles', meta_ns):
                        if el.text and el.text.strip() not in ('', '0'):
                            removed += 1
                        _remove_element(meta_root, el)
                    ET.ElementTree(meta_root).write(
                        meta_xml_path, encoding='UTF-8', xml_declaration=True)
                except ET.ParseError as e:
                    raise RevisionHistoryRemovalError(f"Error parsing meta.xml: {e}")


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


                    # Remove tracked changes. The tracked-changes section
                    # holds the change definitions and, for a deletion, the
                    # deleted text itself, so dropping that section drops
                    # the deleted text. Inserted text lives in the body
                    # between change-start and change-end markers; those
                    # markers are empty, so removing them keeps the text.
                    for change in root.findall('.//text:tracked-changes', ns):
                        regions = change.findall('.//text:changed-region', ns)
                        if _remove_element(root, change):
                            removed += max(1, len(regions))
                    for marker_tag in ('text:change', 'text:change-start', 'text:change-end'):
                        for marker in root.findall(f'.//{marker_tag}', ns):
                            _remove_element(root, marker)

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
        return removed

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

    # Refuse in place output: the resolved output path must not be the
    # input file, under any spelling (same string, ./ alias, symlink).
    if os.path.realpath(input_file) == os.path.realpath(output_file):
        logging.error(
            f"Output file resolves to the input file, refusing to overwrite in place: {output_file}")
        sys.exit(1)

    # Determine file type and call appropriate function
    try:
        if input_file.endswith('.docx'):
            removed = remove_docx_revision_history(input_file, output_file)
        elif input_file.endswith('.odt'):
            removed = remove_odt_revision_history(input_file, output_file)
        else:
            logging.error("Unsupported file type. Only .docx and .odt are supported.")
            sys.exit(1)
        if removed == 0:
            logging.info(
                f"No revision history removed from {input_file}: no tracked changes "
                f"or revision metadata found, output copy saved to {output_file}")
            sys.exit(1)
        logging.info(
            f"Revision history removed from {input_file}: {removed} item(s) removed, "
            f"output saved to {output_file}")
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