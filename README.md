# DOCX Embedded Files Extractor

Extract embedded files from Microsoft Word `.docx` documents with automatic original filename detection.

## Features

- 🔍 Automatically detects and restores original filenames from embedded OLE objects
- 📦 Extracts only embedded files (not media/images)
- 📂 Automatically extracts archives (7z, zip) with the `--extract-archives` flag
- 🚀 Simple command-line interface
- 💻 Standalone executable for easy distribution

## Installation

### For Development

```bash
# Clone or navigate to the project directory
cd docx-extractor

# Install with uv (includes dev dependencies)
uv sync --all-extras

# Or install in editable mode
uv pip install -e ".[dev]"
```

### Using the Standalone Executable

Download the `docx-extractor` executable and run it directly:

```bash
./docx-extractor your_document.docx
```

## Usage

### Command Line

```bash
# Extract to default folder (document_name_extracted)
docx-extractor document.docx

# Extract to custom folder
docx-extractor document.docx ./output_folder

# Extract and automatically unpack archives (7z, zip)
docx-extractor document.docx --extract-archives

# Short form
docx-extractor document.docx -x

# Show help
docx-extractor --help
```

### Python API

```python
from docx_extractor.cli import extract_embedded_files

# Extract embedded files
extracted = extract_embedded_files("document.docx", "output_folder")
print(f"Extracted {len(extracted)} files")

# Extract and automatically unpack archives
extracted = extract_embedded_files(
    "document.docx",
    "output_folder",
    auto_extract_archives=True
)
print(f"Extracted {len(extracted)} files")
```

## Building the Executable

To create a standalone executable for distribution:

```bash
# Build the executable
./build.sh

# The executable will be in dist/docx-extractor
```

The executable is self-contained and can be distributed to collaborators without requiring Python installation.

## Development

### Using mise (Recommended)

This project includes a `.mise.toml` configuration file:

```bash
# Format and lint code
mise run lint

# Build executable
mise run build

# Create distribution package
mise run package
```

### Manual Commands

```bash
# Format code with ruff
uv run ruff format

# Lint code
uv run ruff check

# Fix linting issues
uv run ruff check --fix
```

## Requirements

- Python 3.10 or higher
- olefile library (automatically installed)
- py7zr library for 7z extraction (automatically installed)

## License

MIT
