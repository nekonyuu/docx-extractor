#!/usr/bin/env python3
"""
Extract embedded files from a .docx file.

Usage:
    python extract_docx_embeddings.py <docx_file_path> [output_folder]
"""

import argparse
import os
import shutil
import sys
import zipfile
from pathlib import Path

try:
    import olefile
except ImportError:
    olefile = None

try:
    import py7zr
except ImportError:
    py7zr = None


def extract_ole_package_content(ole_file_path):
    """
    Extract the actual file content from an OLE Package object.

    Args:
        ole_file_path: Path to the OLE file

    Returns:
        Tuple of (filename, file_data) or (None, None) if extraction fails
    """
    if olefile is None:
        return None, None

    try:
        ole = olefile.OleFileIO(ole_file_path)

        if ole.exists("\x01Ole10Native"):
            data = ole.openstream("\x01Ole10Native").read()

            # Known file signatures to search for
            signatures = [
                (b"\x37\x7a\xbc\xaf\x27\x1c", ".7z"),  # 7z archive
                (b"PK\x03\x04", ".zip"),  # ZIP archive
                (b"PK\x05\x06", ".zip"),  # ZIP archive (empty)
                (b"\x1f\x8b", ".gz"),  # GZIP
                (b"BZh", ".bz2"),  # BZIP2
                (b"\x52\x61\x72\x21\x1a\x07", ".rar"),  # RAR
                (b"%PDF", ".pdf"),  # PDF
                (b"\x89PNG", ".png"),  # PNG
                (b"\xff\xd8\xff", ".jpg"),  # JPEG
                (b"GIF8", ".gif"),  # GIF
                (
                    b"\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x4d\x53\x43\x46",
                    ".cab",
                ),  # CAB
            ]

            # Search for file signature in the data
            file_offset = None
            detected_ext = None
            for sig, ext in signatures:
                idx = data.find(sig)
                if idx != -1:
                    file_offset = idx
                    detected_ext = ext
                    break

            # Try to extract filename from Ole10Native structure
            filename = None
            try:
                # Look for null-terminated strings that look like filenames
                # Search in first 500 bytes for efficiency
                search_data = data[: min(500, len(data))]
                strings = search_data.split(b"\x00")
                for s in strings:
                    try:
                        text = s.decode("latin-1", errors="ignore").strip()
                        # Look for strings with file extensions
                        if (
                            text
                            and "." in text
                            and len(text) > 3
                            and any(
                                text.lower().endswith(ext)
                                for ext in [
                                    ".7z",
                                    ".zip",
                                    ".rar",
                                    ".pdf",
                                    ".doc",
                                    ".docx",
                                    ".xls",
                                    ".xlsx",
                                    ".ppt",
                                    ".pptx",
                                    ".txt",
                                    ".jpg",
                                    ".png",
                                    ".gif",
                                ]
                            )
                        ):
                            # Extract basename if it's a path
                            basename = os.path.basename(text.replace("\\", "/"))
                            if len(basename) > 3 and "." in basename:
                                filename = basename
                                break
                    except Exception:
                        continue
            except Exception:
                pass

            # Extract file data from signature position
            if file_offset is not None:
                file_data = data[file_offset:]

                # If we didn't find filename but we detected extension, generate one
                if not filename and detected_ext:
                    filename = f"embedded_file{detected_ext}"

                ole.close()
                return filename, file_data

        ole.close()
    except Exception:
        pass

    return None, None


def get_ole_original_filename(ole_file_path):
    """
    Extract the original filename from an OLE package object.

    Args:
        ole_file_path: Path to the OLE file

    Returns:
        Original filename if found, None otherwise
    """
    if olefile is None:
        return None

    try:
        ole = olefile.OleFileIO(ole_file_path)

        # For Package objects, the filename is in the Ole10Native stream
        if ole.exists("\x01Ole10Native"):
            data = ole.openstream("\x01Ole10Native").read()

            # Try to find filenames by looking for null-terminated strings
            strings = data.split(b"\x00")
            candidates = []

            for s in strings:
                try:
                    text = s.decode("latin-1", errors="ignore").strip()
                    # Look for Windows-style paths or reasonable filenames
                    if text and 5 < len(text) < 300:
                        # Check for path separators or standalone filenames
                        if ("\\" in text or "/" in text) and "." in text:
                            # Extract basename from path
                            basename = os.path.basename(text.replace("\\", "/"))
                            if "." in basename and 3 < len(basename) < 100:
                                # Check extension
                                ext = basename.rsplit(".", 1)[-1]
                                # Valid extension: alphanumeric, 2-10 chars
                                # Make sure basename doesn't have weird chars (only printable ASCII)
                                if (
                                    1 < len(ext) <= 10
                                    and all(c.isalnum() for c in ext)
                                    and all(32 <= ord(c) < 127 for c in basename)
                                ):
                                    candidates.append(basename)
                        elif "." in text and not ("\\" in text or "/" in text):
                            # Standalone filename without path
                            if 3 < len(text) < 100:
                                ext = text.rsplit(".", 1)[-1]
                                if (
                                    1 < len(ext) <= 10
                                    and all(c.isalnum() for c in ext)
                                    and all(32 <= ord(c) < 127 for c in text)
                                ):
                                    candidates.append(text)
                except Exception:
                    continue

            # Prioritize filenames without temp indicators or parentheses
            if candidates:
                # Remove duplicates
                candidates = list(set(candidates))
                # Filter out temp files and files with version numbers in parentheses
                good_candidates = [
                    c
                    for c in candidates
                    if not any(x in c.lower() for x in ["~", "temp", "tmp"])
                    and "(" not in c
                    and ")" not in c
                ]
                if good_candidates:
                    # Return the longest good candidate
                    good_candidates.sort(key=len, reverse=True)
                    ole.close()
                    return good_candidates[0]
                elif candidates:
                    # Fallback to any candidate, prefer ones without parentheses
                    no_parens = [c for c in candidates if "(" not in c and ")" not in c]
                    if no_parens:
                        no_parens.sort(key=len, reverse=True)
                        ole.close()
                        return no_parens[0]
                    candidates.sort(key=len, reverse=True)
                    ole.close()
                    return candidates[0]

        ole.close()
    except Exception:
        pass

    return None


def extract_archive(archive_path, extract_to=None):
    """
    Extract an archive file (7z, zip, etc.) to a folder.

    Args:
        archive_path: Path to the archive file
        extract_to: Optional extraction folder. If not provided, creates a folder
                   named after the archive without extension

    Returns:
        Path to extraction folder if successful, None otherwise
    """
    archive_path = Path(archive_path)

    extract_to = archive_path.parent / archive_path.stem if extract_to is None else Path(extract_to)

    extract_to.mkdir(parents=True, exist_ok=True)

    try:
        # Try 7z extraction
        if archive_path.suffix.lower() == ".7z":
            if py7zr is None:
                print(f"  ⚠ Cannot extract {archive_path.name}: py7zr not available")
                return None

            with py7zr.SevenZipFile(archive_path, mode="r") as archive:
                archive.extractall(path=extract_to)
            return extract_to

        # Try ZIP extraction
        elif archive_path.suffix.lower() in [".zip"]:
            with zipfile.ZipFile(archive_path, "r") as archive:
                archive.extractall(path=extract_to)
            return extract_to

        else:
            # Unknown archive type
            return None

    except Exception as e:
        print(f"  ⚠ Failed to extract {archive_path.name}: {e}")
        return None


def extract_embedded_files(docx_path, output_folder=None, auto_extract_archives=False):
    """
    Extract all embedded files from a .docx file.

    Args:
        docx_path: Path to the .docx file
        output_folder: Optional output folder path. If not provided, creates a folder
                      named after the .docx file with '_extracted' suffix
        auto_extract_archives: If True, automatically extract any archive files (7z, zip)

    Returns:
        List of extracted file paths
    """
    docx_path = Path(docx_path)

    if not docx_path.exists():
        raise FileNotFoundError(f"File not found: {docx_path}")

    if not docx_path.suffix.lower() == ".docx":
        raise ValueError(f"File must be a .docx file, got: {docx_path.suffix}")

    # Create output folder
    if output_folder is None:
        output_folder = docx_path.parent / f"{docx_path.stem}_extracted"
    else:
        output_folder = Path(output_folder)

    output_folder.mkdir(parents=True, exist_ok=True)

    extracted_files = []

    # Open the .docx file as a ZIP archive
    with zipfile.ZipFile(docx_path, "r") as zip_ref:
        # Get all files in the archive
        all_files = zip_ref.namelist()

        # Extract embedded files (OLE objects)
        embeddings_folder = "word/embeddings/"
        embedding_files = [
            f for f in all_files if f.startswith(embeddings_folder) and not f.endswith("/")
        ]

        if embedding_files:
            print(f"Found {len(embedding_files)} embedded file(s) in word/embeddings/:")
            for file_path in embedding_files:
                filename = os.path.basename(file_path)
                temp_ole_path = output_folder / f"_temp_{filename}"

                # Extract the OLE file temporarily
                with zip_ref.open(file_path) as source, open(temp_ole_path, "wb") as target:
                    shutil.copyfileobj(source, target)

                # Extract the actual file content from the OLE package
                original_filename, file_data = extract_ole_package_content(temp_ole_path)

                # Clean up temp OLE file
                temp_ole_path.unlink()

                if file_data:
                    # Use original filename or fallback to oleObject name
                    if not original_filename:
                        original_filename = filename.replace(".bin", "")

                    # Determine final output path
                    final_output_path = output_folder / original_filename
                    # Handle duplicate filenames
                    counter = 1
                    while final_output_path.exists():
                        name, ext = os.path.splitext(original_filename)
                        final_output_path = output_folder / f"{name}_{counter}{ext}"
                        counter += 1

                    # Write the actual file content
                    with open(final_output_path, "wb") as f:
                        f.write(file_data)

                    extracted_files.append(final_output_path)
                    print(
                        f"  Extracted: {filename} → {final_output_path.name} ({final_output_path.stat().st_size:,} bytes)"
                    )

                    # Auto-extract archives if enabled
                    if auto_extract_archives and final_output_path.suffix.lower() in [
                        ".7z",
                        ".zip",
                    ]:
                        print(f"  Extracting archive: {final_output_path.name}")
                        archive_extract_path = extract_archive(final_output_path)
                        if archive_extract_path:
                            # Count extracted files
                            archive_files = list(archive_extract_path.rglob("*"))
                            file_count = sum(1 for f in archive_files if f.is_file())
                            print(
                                f"    → Extracted {file_count} file(s) to {archive_extract_path.name}/"
                            )
                else:
                    # Fallback: if we can't extract the package content, keep the OLE file
                    fallback_path = output_folder / filename
                    with zip_ref.open(file_path) as source, open(fallback_path, "wb") as target:
                        shutil.copyfileobj(source, target)
                    extracted_files.append(fallback_path)
                    print(
                        f"  Extracted (as OLE): {filename} ({fallback_path.stat().st_size:,} bytes)"
                    )
        else:
            print("No embedded files found in word/embeddings/")

    return extracted_files


def main():
    parser = argparse.ArgumentParser(
        description="Extract embedded files from Microsoft Word .docx documents",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  docx-extractor document.docx
  docx-extractor document.docx ./output_folder
  docx-extractor document.docx --output ./extracted
  docx-extractor document.docx --extract-archives
        """,
    )
    parser.add_argument("docx_file", help="Path to the .docx file to extract from")
    parser.add_argument(
        "output_folder",
        nargs="?",
        help="Output folder (default: <docx_filename>_extracted)",
    )
    parser.add_argument(
        "-o",
        "--output",
        dest="output_alt",
        help="Alternative way to specify output folder",
    )
    parser.add_argument(
        "-x",
        "--extract-archives",
        action="store_true",
        help="Automatically extract archive files (7z, zip)",
    )
    parser.add_argument("--version", action="version", version="%(prog)s 0.1.0")

    args = parser.parse_args()

    # Use output_alt if provided via -o flag, otherwise use positional argument
    output_folder = args.output_alt if args.output_alt else args.output_folder
    docx_path = args.docx_file

    try:
        print(f"Extracting embedded files from: {docx_path}")
        print(f"Output folder: {output_folder or Path(docx_path).stem + '_extracted'}")
        if args.extract_archives:
            print("Auto-extract archives: enabled")
        print()

        extracted_files = extract_embedded_files(
            docx_path, output_folder, auto_extract_archives=args.extract_archives
        )

        print(f"\n{'=' * 60}")
        print(f"Total files extracted: {len(extracted_files)}")
        print(
            f"Output folder: {Path(output_folder or Path(docx_path).stem + '_extracted').resolve()}"
        )
        print(f"{'=' * 60}")

    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
