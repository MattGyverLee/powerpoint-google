#!/usr/bin/env python3
"""
PowerPoint Cleaner - Fix Google Slides Compatibility Issues

This script fixes two common issues that prevent PowerPoint files from being
read by Google Slides:
1. Malformed ZIP structure with extra padding bytes
2. LibreOffice metadata that Google Slides rejects

Usage:
    python clean_powerpoint.py <input.pptx> [output.pptx]
    python clean_powerpoint.py --batch <directory>

Examples:
    # Clean a single file (creates input-CLEAN.pptx)
    python clean_powerpoint.py "presentation.pptx"

    # Clean with custom output name
    python clean_powerpoint.py "presentation.pptx" "fixed.pptx"

    # Clean all .pptx files in a directory
    python clean_powerpoint.py --batch "C:\\Users\\Documents\\Presentations"
"""

import zipfile
import os
import sys
import argparse
from pathlib import Path


def clean_app_xml(xml_content):
    """
    Replace LibreOffice metadata with Microsoft Office metadata.

    Args:
        xml_content: The app.xml content as bytes or string

    Returns:
        Cleaned XML content as bytes
    """
    if isinstance(xml_content, bytes):
        xml_str = xml_content.decode('utf-8')
    else:
        xml_str = xml_content

    # Replace LibreOffice application strings
    replacements = [
        # LibreOffice Impress
        ('LibreOffice/', 'Microsoft Office PowerPoint'),
        # Remove LibreOffice project references
        ('Impress', 'PowerPoint'),
    ]

    for old, new in replacements:
        if old in xml_str:
            # Find and replace the entire Application tag content
            import re
            pattern = r'<Application>.*?LibreOffice.*?</Application>'
            xml_str = re.sub(pattern, '<Application>Microsoft Office PowerPoint</Application>', xml_str)
            break

    return xml_str.encode('utf-8')


def clean_powerpoint(input_path, output_path=None, verbose=True):
    """
    Clean a PowerPoint file to fix Google Slides compatibility issues.

    Args:
        input_path: Path to the input .pptx file
        output_path: Path for the output file (optional)
        verbose: Print progress messages

    Returns:
        Path to the cleaned file, or None if failed
    """
    input_path = Path(input_path)

    if not input_path.exists():
        print(f"Error: File not found: {input_path}")
        return None

    if not input_path.suffix.lower() in ['.pptx', '.ppt']:
        print(f"Error: Not a PowerPoint file: {input_path}")
        return None

    # Determine output path
    if output_path is None:
        output_path = input_path.parent / f"{input_path.stem}-CLEAN{input_path.suffix}"
    else:
        output_path = Path(output_path)

    if verbose:
        print(f"Cleaning: {input_path.name}")

    try:
        # Read all files from the original PowerPoint
        with zipfile.ZipFile(input_path, 'r') as zip_read:
            # Test the original file
            bad_file = zip_read.testzip()
            if bad_file:
                print(f"  Warning: Original file has corrupted entry: {bad_file}")

            # Read all files into memory
            all_files = {}
            for name in zip_read.namelist():
                all_files[name] = zip_read.read(name)

        # Fix app.xml if it exists and contains LibreOffice references
        if 'docProps/app.xml' in all_files:
            original_app_xml = all_files['docProps/app.xml']
            cleaned_app_xml = clean_app_xml(original_app_xml)

            if original_app_xml != cleaned_app_xml:
                all_files['docProps/app.xml'] = cleaned_app_xml
                if verbose:
                    print("  ✓ Fixed LibreOffice metadata")

        # Write to new file with clean ZIP structure
        # This automatically removes any extra padding/malformed structure
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_write:
            for filename, data in all_files.items():
                zip_write.writestr(filename, data)

        # Verify the new file
        with zipfile.ZipFile(output_path, 'r') as zip_verify:
            bad_file = zip_verify.testzip()
            if bad_file:
                print(f"  Error: Output file has corrupted entry: {bad_file}")
                return None

        # Compare sizes
        original_size = input_path.stat().st_size
        cleaned_size = output_path.stat().st_size
        size_diff = original_size - cleaned_size

        if verbose:
            print(f"  ✓ Original size: {original_size:,} bytes")
            print(f"  ✓ Cleaned size:  {cleaned_size:,} bytes")
            if size_diff > 0:
                print(f"  ✓ Removed {size_diff:,} bytes of padding/metadata")
            print(f"  ✓ Saved to: {output_path.name}")

        return output_path

    except zipfile.BadZipFile:
        print(f"  Error: Not a valid ZIP/PowerPoint file")
        return None
    except Exception as e:
        print(f"  Error: {str(e)}")
        return None


def batch_clean(directory, verbose=True):
    """
    Clean all PowerPoint files in a directory.

    Args:
        directory: Path to directory containing .pptx files
        verbose: Print progress messages

    Returns:
        Tuple of (success_count, failure_count)
    """
    directory = Path(directory)

    if not directory.exists() or not directory.is_dir():
        print(f"Error: Directory not found: {directory}")
        return 0, 0

    # Find all PowerPoint files
    pptx_files = list(directory.glob("*.pptx")) + list(directory.glob("*.ppt"))

    if not pptx_files:
        print(f"No PowerPoint files found in: {directory}")
        return 0, 0

    print(f"Found {len(pptx_files)} PowerPoint file(s)")
    print()

    success_count = 0
    failure_count = 0

    for pptx_file in pptx_files:
        # Skip already cleaned files
        if '-CLEAN' in pptx_file.stem:
            if verbose:
                print(f"Skipping (already cleaned): {pptx_file.name}")
            continue

        result = clean_powerpoint(pptx_file, verbose=verbose)
        if result:
            success_count += 1
        else:
            failure_count += 1

        if verbose:
            print()  # Blank line between files

    return success_count, failure_count


def main():
    parser = argparse.ArgumentParser(
        description='Clean PowerPoint files for Google Slides compatibility',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s presentation.pptx
  %(prog)s "My Slides.pptx" "Fixed Slides.pptx"
  %(prog)s --batch "C:\\Users\\Documents\\Presentations"
  %(prog)s --batch . --quiet

This tool fixes:
  1. Malformed ZIP structure (extra padding bytes)
  2. LibreOffice metadata incompatible with Google Slides
        """
    )

    parser.add_argument('input', help='Input PowerPoint file or directory (with --batch)')
    parser.add_argument('output', nargs='?', help='Output file path (optional)')
    parser.add_argument('--batch', action='store_true', help='Process all .pptx files in directory')
    parser.add_argument('--quiet', '-q', action='store_true', help='Suppress progress messages')

    args = parser.parse_args()

    verbose = not args.quiet

    if args.batch:
        if args.output:
            print("Warning: Output path ignored in batch mode")

        success, failure = batch_clean(args.input, verbose=verbose)

        print("=" * 50)
        print(f"Batch processing complete:")
        print(f"  ✓ Successfully cleaned: {success}")
        if failure > 0:
            print(f"  ✗ Failed: {failure}")

        sys.exit(0 if failure == 0 else 1)
    else:
        result = clean_powerpoint(args.input, args.output, verbose=verbose)
        sys.exit(0 if result else 1)


if __name__ == '__main__':
    # If no arguments provided, show help
    if len(sys.argv) == 1:
        print(__doc__)
        sys.exit(0)

    main()
