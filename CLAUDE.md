# DOCX Embedded Files Extractor - Context for Claude

This document provides context for Claude Code sessions working on this project.

## Project Overview

A Python CLI tool that extracts embedded files from Microsoft Word `.docx` documents with automatic original filename detection and optional archive extraction.

**Key Features:**
- Extracts files from `word/embeddings/` folder within docx (ZIP) archives
- Parses OLE (Object Linking and Embedding) packages to extract embedded content
- Detects original filenames from OLE structure metadata
- Optionally auto-extracts archives (7z, ZIP) with `--extract-archives` flag
- Distributable as standalone macOS executable via PyInstaller

## Architecture & Design Decisions

### OLE Package Extraction

The core challenge: `.docx` files store embedded objects as OLE containers, not raw files.

**Problem:**
- Embedded files in `word/embeddings/` are stored as OLE packages (Composite Document File V2)
- Simply extracting these gives you the OLE wrapper, not the actual file

**Solution (src/docx_extractor/cli.py:27):**
```python
def extract_ole_package_content(ole_file_path):
```
- Uses `olefile` library to parse OLE structure
- Searches for known file signatures (7z: `\x37\x7A\xBC\xAF\x27\x1C`, ZIP: `PK\x03\x04`, etc.)
- Extracts content from signature position to end of stream
- Attempts to detect original filename from OLE metadata (null-terminated strings in first 500 bytes)

**Why signature-based extraction:**
- Ole10Native structure varies between Office versions
- Field-based parsing proved unreliable (incorrect field sizes)
- Signature search is robust and format-agnostic

### Archive Extraction

**Function (src/docx_extractor/cli.py:224):**
```python
def extract_archive(archive_path, extract_to=None):
```
- Uses `py7zr` for 7z archives
- Uses Python's built-in `zipfile` for ZIP archives
- Only extracts when `--extract-archives` / `-x` flag is provided
- Creates extraction folder named after archive (without extension)

## Project Structure

```
extract-embedded-files/
├── src/docx_extractor/
│   ├── __init__.py              # Package metadata
│   └── cli.py                    # Main CLI application (ALL code here)
├── build.sh                      # PyInstaller build script
├── package.sh                    # Distribution packaging script
├── pyproject.toml                # uv/pip package configuration
├── .mise.toml                    # mise tool configuration
├── .gitignore                    # Git ignore patterns
├── README.md                     # User documentation
└── CLAUDE.md                     # This file

After build:
├── dist/docx-extractor          # Standalone executable (~13MB)
├── build/                        # PyInstaller build artifacts
└── docx-extractor.spec          # PyInstaller spec file
```

## Key Files

### `src/docx_extractor/cli.py`
**Single-file architecture** - All functionality in one module:
- `extract_ole_package_content()`: OLE package parsing and content extraction
- `get_ole_original_filename()`: Legacy filename detection (currently unused but kept)
- `extract_archive()`: Archive extraction (7z, ZIP)
- `extract_embedded_files()`: Main extraction logic
- `main()`: CLI argument parsing and orchestration

### `pyproject.toml`
- Uses `uv` for dependency management
- Runtime dependencies: `olefile>=0.47`, `py7zr>=0.20.0`
- Dev dependencies: `ruff>=0.8.0`, `pyinstaller>=6.0.0`
- Entry point: `docx-extractor = "docx_extractor.cli:main"`

### `build.sh`
- Runs `uv sync --all-extras` to install dependencies
- Executes PyInstaller with `--onefile` to create standalone executable
- Final binary: `dist/docx-extractor` (~13MB with all crypto deps for 7z)

### `.mise.toml`
- Python version: 3.13
- Tasks: `lint`, `build`, `package`, `test`

## Development Workflow

### Setup
```bash
# Install dependencies
uv sync --all-extras

# Trust mise config (first time only)
mise trust
```

### Common Tasks
```bash
# Format and lint
mise run lint
# OR manually:
uv run ruff format
uv run ruff check

# Build executable
mise run build
# OR manually:
./build.sh

# Test locally
uv run docx-extractor document.docx --extract-archives

# Create distribution package
mise run package
# OR manually:
./package.sh
```

## Important Implementation Details

### Why Only Embeddings?

User explicitly requested **only** embedded files, not media (images).
- Embedded files: `word/embeddings/` folder
- Media files: `word/media/` folder (IGNORED)

### OLE Structure Challenges

**Initial approach (FAILED):**
Tried parsing Ole10Native structure with fixed field layout:
- Total size (4 bytes)
- Label size + label
- Filename size + filename
- Unknown field (2 bytes)
- Temp path size + temp path
- File size (4 bytes)
- File data

**Why it failed:**
- Field sizes were incorrect (e.g., filename_size: 29807 when should be ~40)
- Structure varies between Office versions
- Data at offset 208 showed 7z signature, not at calculated position

**Current approach (WORKS):**
- Search for file signatures directly
- Extract from signature to end of stream
- Parse filename separately from metadata area

### Archive Extraction Edge Cases

- Only extracts if `--extract-archives` / `-x` flag provided (not default)
- Silently skips if `py7zr` unavailable (prints warning)
- Creates extraction folder named after archive without extension
- Counts and reports number of files extracted

### File Signature Reference

From `extract_ole_package_content()`:
```python
signatures = [
    (b"\x37\x7A\xBC\xAF\x27\x1C", ".7z"),    # 7z archive
    (b"PK\x03\x04", ".zip"),                 # ZIP archive
    (b"PK\x05\x06", ".zip"),                 # ZIP archive (empty)
    (b"\x1F\x8B", ".gz"),                    # GZIP
    (b"BZh", ".bz2"),                        # BZIP2
    (b"\x52\x61\x72\x21\x1A\x07", ".rar"),  # RAR
    (b"%PDF", ".pdf"),                       # PDF
    (b"\x89PNG", ".png"),                    # PNG
    (b"\xFF\xD8\xFF", ".jpg"),              # JPEG
    (b"GIF8", ".gif"),                       # GIF
]
```

## Testing

### Test File
`proof_file_1BARKEA-SERVID01-RECORD-20260110204846-KZN4YD2C23YSQ625.docx`
- Contains 1 embedded file: `oleObject1.bin`
- OLE package contains: `Protect&Sign-FichierDePreuve.7z` (3.5MB)
- Archive contains: 256 files including `ProofData-Summary.html`

### Test Commands
```bash
# Extract only
./dist/docx-extractor proof_file_*.docx

# Extract + unpack archive
./dist/docx-extractor proof_file_*.docx -x

# Verify extraction
find *_extracted -name "ProofData-Summary.html"
```

## Known Limitations

1. **Only .docx format** - Does not support legacy .doc files
2. **Only embeddings** - Ignores media files (by design)
3. **Archive formats** - Only 7z and ZIP auto-extraction supported
4. **macOS executable** - PyInstaller build configured for macOS only
5. **No recursive docx** - Doesn't extract embedded docx files within docx

## Troubleshooting

### "Cannot extract: py7zr not available"
- Executable was built without py7zr included
- Rebuild with `./build.sh` after `uv sync --all-extras`

### Extracted file is OLE, not actual file
- Check if `extract_ole_package_content()` found file signature
- Verify OLE file with: `file extracted_file.bin`
- Should show specific format (7z, ZIP), not "Composite Document File"

### No files extracted
- Check if docx has `word/embeddings/` folder: `unzip -l document.docx`
- Tool only extracts from embeddings, not media

## CLI Usage Reference

```bash
# Basic extraction
docx-extractor document.docx

# Custom output folder
docx-extractor document.docx ./output
docx-extractor document.docx -o ./output

# Auto-extract archives
docx-extractor document.docx -x
docx-extractor document.docx --extract-archives

# Show help
docx-extractor --help

# Show version
docx-extractor --version
```

## Python API Reference

```python
from docx_extractor.cli import extract_embedded_files

# Basic extraction
files = extract_embedded_files("document.docx")

# Custom output folder
files = extract_embedded_files("document.docx", "output_folder")

# With archive extraction
files = extract_embedded_files(
    "document.docx",
    "output_folder",
    auto_extract_archives=True
)
```

## Version History

**v0.1.0** (Current)
- Initial release
- OLE package content extraction
- Original filename detection
- Optional archive extraction (7z, ZIP)
- Standalone macOS executable
- mise task integration

## Future Enhancement Ideas

**NOT currently requested, but potential improvements:**
- Support for other archive formats (RAR, tar.gz)
- Windows/Linux executable builds
- Recursive extraction (docx within docx)
- GUI version
- Batch processing of multiple docx files
- Progress bars for large files
- Extraction verification (checksum)

## Contact & Issues

For issues or questions, refer to project maintainer.
