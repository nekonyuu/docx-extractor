#!/bin/bash
set -e

echo "Building docx-extractor executable for macOS..."

# Clean previous builds
rm -rf build dist *.spec

# Install dependencies
echo "Installing dependencies..."
uv sync --all-extras

# Build with PyInstaller
echo "Building executable with PyInstaller..."
uv run pyinstaller \
    --name docx-extractor \
    --onefile \
    --console \
    --clean \
    --noconfirm \
    --add-data "src/docx_extractor/__init__.py:docx_extractor" \
    src/docx_extractor/cli.py

echo ""
echo "✅ Build complete!"
echo "📦 Executable location: dist/docx-extractor"
echo ""
echo "Test it with:"
echo "  ./dist/docx-extractor --help"
