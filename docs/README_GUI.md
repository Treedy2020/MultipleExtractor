# DOCX Document Extractor - GUI Application

A user-friendly GUI application for extracting structured information from DOCX documents using Azure OpenAI.

## Features

- **Simple GUI Interface**: Easy-to-use graphical interface with pre-configured Azure OpenAI settings
- **Batch Processing**: Async batch processing with dynamic token windowing
- **Progress Tracking**: Real-time progress bar and detailed logging
- **Multiple Output Formats**: Export to CSV or JSON
- **Flexible Configuration**: Customize batch size, concurrency, and token limits

## Pre-configured Settings

The application comes with default Azure OpenAI configuration:
- **Endpoint**: `https://aidt-eastus2-ai.openai.azure.com`
- **API Version**: `2025-03-01-preview`
- **Deployment**: `gpt-5-chat`

**You only need to provide your API Key!**

## Installation

### Option 1: Download Pre-built Binaries (Recommended)

Download the latest release for your platform:

**Windows:**
1. Go to [Releases](https://github.com/yourusername/HandWriting/releases)
2. Download `DOCX_Extractor_Windows_x64.zip`
3. Extract and run `DOCX_Extractor.exe`

**macOS:**
1. Go to [Releases](https://github.com/yourusername/HandWriting/releases)
2. Download `DOCX_Extractor_macOS_x64.tar.gz`
3. Extract: `tar -xzf DOCX_Extractor_macOS_x64.tar.gz`
4. Move to Applications: `mv DOCX_Extractor.app /Applications/`
5. If macOS blocks the app, go to System Preferences > Security & Privacy and click "Open Anyway"

### Option 2: Run from Source

```bash
# Install dependencies with uv
uv sync

# Run GUI
uv run python gui_app.py

# Or install and run
uv pip install -e .
docx-extract-gui
```

## Usage

1. **Launch the application**
   - Windows: Double-click `DOCX_Extractor.exe`
   - macOS: Open `DOCX_Extractor.app` from Applications
   - From source: `uv run python gui_app.py`

2. **Enter API Key**
   - Paste your Azure OpenAI API Key in the first field
   - Other fields are pre-configured but can be modified if needed

3. **Select Files**
   - Click "Browse" next to "Input File" and select your DOCX file
   - Click "Browse" next to "Output File" to choose where to save results
   - Select output format (CSV or JSON)

4. **Configure Processing Options** (Optional)
   - **Enable Batch Processing Mode**: Enable for large documents
   - **Batch Size**: Number of table rows per batch (default: 10)
   - **Max Concurrent API Calls**: Number of parallel API requests (default: 5)
   - **Max Tokens per Batch**: Token limit per batch to prevent truncation (default: 20000)

5. **Start Processing**
   - Click "Start Processing"
   - Monitor progress in the log area
   - Wait for completion message

## Building from Source

### Prerequisites

- Python 3.12+
- uv package manager

### Build Instructions

```bash
# Install build dependencies
uv sync
uv pip install pyinstaller

# Build for your platform
uv run pyinstaller docx_extractor_gui.spec

# Executable will be in dist/ folder
# Windows: dist/DOCX_Extractor.exe
# macOS: dist/DOCX_Extractor.app
```

## Creating a Release

To create a new release with automated builds for Windows and macOS:

1. **Create and push a version tag:**
   ```bash
   git tag -a v1.0.0 -m "Release version 1.0.0"
   git push origin v1.0.0
   ```

2. **GitHub Actions will automatically:**
   - Build executables for Windows and macOS
   - Create a new GitHub Release
   - Upload the binaries as release assets

3. **The release will include:**
   - `DOCX_Extractor_Windows_x64.zip` - Windows executable
   - `DOCX_Extractor_macOS_x64.tar.gz` - macOS application bundle
   - Detailed installation and usage instructions

## Command Line Alternative

If you prefer using the command line:

```bash
# Basic usage with batch mode
uv run python main.py document.docx --batch-mode -o output.csv

# With custom settings
uv run python main.py document.docx \
  --batch-mode \
  --batch-size 20 \
  --max-concurrent 10 \
  --max-tokens 25000 \
  -o output.csv

# See all options
uv run python main.py --help
```

## Troubleshooting

### Windows: "Windows protected your PC"
- Click "More info" → "Run anyway"

### macOS: "Cannot be opened because it is from an unidentified developer"
- Right-click the app → "Open" → "Open" again
- Or go to System Preferences → Security & Privacy → Click "Open Anyway"

### API Key Issues
- Ensure your API key is valid and has access to the specified deployment
- Check that the endpoint and deployment name match your Azure OpenAI setup

### Large Documents
- Increase "Max Tokens per Batch" if processing very long rows
- Decrease "Batch Size" if hitting token limits
- Adjust "Max Concurrent API Calls" based on your API rate limits

## License

[Your License Here]

## Support

For issues and questions, please [open an issue](https://github.com/yourusername/HandWriting/issues).
