# DOCX Document Extractor

A powerful tool for extracting structured information from DOCX documents using Azure OpenAI API with intelligent batch processing and dynamic token management.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.12+](https://img.shields.io/badge/python-3.12+-blue.svg)](https://www.python.org/downloads/)

## Features

- 🚀 **Batch Processing with Async API Calls**: Process large documents efficiently with concurrent API requests
- 🎯 **Dynamic Token Windowing**: Automatically adjusts batch sizes to prevent output truncation
- 🖥️ **User-Friendly GUI**: Simple graphical interface with pre-configured Azure OpenAI settings
- 📊 **Multiple Output Formats**: Export results to CSV or JSON
- ⚡ **Real-time Progress Tracking**: Monitor processing with progress bars and detailed logging
- 🔄 **Smart Content Extraction**: Intelligently extracts data from table columns with validation
- 🏷️ **Image Detection**: Identifies and marks cells containing images (labels, symbols, etc.)

## Quick Start

### Command Line Interface

```bash
# Install dependencies
uv sync

# Set your API key (or configure in .env file)
export AZURE_OPENAI_API_KEY="your-api-key-here"

# Basic usage with batch mode (recommended)
uv run python main.py document.docx --batch-mode -o output.csv

# With custom settings
uv run python main.py document.docx \
  --batch-mode \
  --batch-size 20 \
  --max-concurrent 10 \
  --max-tokens 25000 \
  -o output.csv

# Process entire folder
uv run python main.py ./documents --batch-mode -o results.csv
```

## Installation

### Prerequisites

- Python 3.12 or higher
- Azure OpenAI API access with deployment
- [uv](https://github.com/astral-sh/uv) package manager (recommended)

### Install from Source

```bash
# Clone the repository
git clone https://github.com/yourusername/HandWriting.git
cd HandWriting

# Install with uv
uv sync

# Or with pip
pip install -e .
```

## Configuration

### Environment Variables

Create a `.env` file in the project root:

```env
# Azure OpenAI Configuration (default)
USE_AZURE_OPENAI=true
AZURE_OPENAI_API_KEY=your-api-key-here
AZURE_OPENAI_ENDPOINT=https://aidt-eastus2-ai.openai.azure.com
AZURE_OPENAI_API_VERSION=2025-03-01-preview
AZURE_OPENAI_DEPLOYMENT=gpt-5-chat

# Or use Standard OpenAI
USE_AZURE_OPENAI=false
OPENAI_API_KEY=your-api-key-here
OPENAI_API_BASE=https://api.openai.com/v1
OPENAI_MODEL=gpt-4o-2024-08-06
```

### Default Settings

Default configuration for batch processing:
- **Batch Size**: 10 rows per batch
- **Max Concurrent**: 5 API calls
- **Max Tokens**: 20,000 per batch
- **API Version**: 2025-03-01-preview
- **Deployment**: gpt-5-chat

## Usage Examples

### Command Line Interface

#### Single File Processing

```bash
# Standard mode (processes entire document at once)
uv run python main.py document.docx -o output.json

# Batch mode (recommended for large documents)
uv run python main.py document.docx --batch-mode -o output.csv
```

#### Batch Processing with Custom Settings

```bash
uv run python main.py document.docx \
  --batch-mode \
  --batch-size 15 \          # Rows per batch
  --max-concurrent 8 \        # Concurrent API calls
  --max-tokens 30000 \        # Max tokens per batch
  -o results.csv
```

#### Folder Processing

```bash
# Process all DOCX files in a folder
uv run python main.py ./documents --batch-mode -o combined_results.csv
```

#### Convert to Markdown (No API needed)

```bash
# Convert tables to Markdown format
uv run python main.py document.docx --to-markdown -o output.md
```

### Python API

```python
from main import DocxExtractor

# Initialize extractor
extractor = DocxExtractor(
    api_key="your-api-key",
    model="gpt-5-chat",
    use_azure=True,
    azure_endpoint="https://aidt-eastus2-ai.openai.azure.com",
    azure_api_version="2025-03-01-preview"
)

# Standard processing
extraction = extractor.process_file("document.docx", "output.json")

# Batch processing
extraction = extractor.process_file_with_batches(
    "document.docx",
    batch_size=10,
    max_concurrent=5,
    max_tokens_per_batch=20000,
    output_path="output.csv"
)

# Access results
for record in extraction.records:
    print(f"TL EA: {record.tl_ea}")
    print(f"Test Standard: {record.test_standard}")
    print(f"Test Analytes: {record.test_analytes}")
    # ... etc
```

## Extracted Fields

The tool extracts the following structured information from each table row:

| Field | Description | Source |
|-------|-------------|--------|
| **TL EA** | Test Lab EA information | Column 1 |
| **Test Standard** | Test standard name (excluding URLs) | Column 2 |
| **Test Analytes** | Test analytes information | Column 5 only |
| **PP Notes** | Product/Process notes | Column 3 |
| **Source Link** | Website URL if present | Column 2 |
| **Label & Symbol** | Whether row contains images | Any column |

### Output Format

**CSV Output:**
```csv
Source File,TL EA,Test Standard,Test Analytes,PP Notes,Source Link,Label and Symbol
document.docx,Physical and Mechanical,EN 71-1,,"Shall comply with requirements",https://example.com,yes
```

**JSON Output:**
```json
{
  "records": [
    {
      "tl_ea": "Physical and Mechanical",
      "test_standard": "EN 71-1",
      "test_analytes": "",
      "pp_notes": "Shall comply with requirements",
      "source_link": "https://example.com",
      "label_and_symbol": "yes"
    }
  ]
}
```

## Advanced Features

### Dynamic Token Management

The tool automatically adjusts batch sizes based on content length to prevent API token limits:

```bash
# Set maximum tokens per batch (default: 20000)
uv run python main.py document.docx --batch-mode --max-tokens 30000
```

**How it works:**
1. Estimates token count for each row
2. Groups rows into batches under the token limit
3. Reserves tokens for system prompt and output
4. Automatically splits large batches

### Concurrent Processing

Process multiple batches simultaneously:

```bash
# Adjust concurrency based on your API rate limits
uv run python main.py document.docx \
  --batch-mode \
  --max-concurrent 10  # Up to 10 concurrent API calls
```

### Progress Monitoring

**CLI:**
```
Reading file: document.docx
  Batch 1: 8 rows, ~15234 tokens
  Batch 2: 6 rows, ~18956 tokens
  Batch 3: 10 rows, ~12445 tokens
Document split into 3 batch(es)
Processing with up to 5 concurrent API calls...

Processing batches: 100%|████████| 3/3 [00:15<00:00,  5.23s/batch]
```

**GUI:**
- Real-time log with detailed progress
- Progress bar animation
- Success/error notifications

## All Command Options

View all available options:

```bash
uv run python main.py --help
```

Key options:
- `--batch-mode`: Enable batch processing with async API calls
- `--batch-size N`: Rows per batch (default: 10)
- `--max-concurrent N`: Max concurrent API calls (default: 5)
- `--max-tokens N`: Max input tokens per batch (default: 20000)
- `-o, --output FILE`: Output file path (.json or .csv)
- `--to-markdown`: Convert tables to Markdown (no API needed)
- `--api-key KEY`: Azure OpenAI API key
- `--model NAME`: Model/deployment name

## Development

### Project Structure

```
HandWriting/
├── main.py           # CLI tool and core extraction logic
├── pyproject.toml    # Project configuration and dependencies
├── .env              # Environment variables (create this)
└── README.md         # This file
```

### Running Tests

```bash
# Syntax check
python -m py_compile main.py

# Test with a sample document
uv run python main.py sample.docx --batch-mode -o test_output.csv

# Test without API (Markdown conversion)
uv run python main.py sample.docx --to-markdown -o test_output.md
```

### Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Troubleshooting

### Common Issues

**Connection Errors:**
- Verify your API key is valid
- Check endpoint URL and API version
- Ensure internet connection

**Token Limit Errors:**
- Reduce `--batch-size`
- Increase `--max-tokens`
- Check row content length

**Empty Output:**
- Check input file format (must be .docx)
- Verify table structure in document
- Review extraction logs for errors

### API Rate Limits

If you encounter rate limit errors:
```bash
# Reduce concurrent requests
uv run python main.py document.docx --batch-mode --max-concurrent 3

# Add delays between batches (modify source if needed)
```

## Requirements

- **Python**: 3.12 or higher
- **Dependencies**:
  - openai >= 2.7.1
  - python-docx >= 1.2.0
  - pydantic >= 2.0.0
  - tqdm >= 4.66.0
  - httpx[socks] >= 0.28.1

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Built with [OpenAI API](https://openai.com/)
- Uses [python-docx](https://python-docx.readthedocs.io/) for document parsing
- GUI built with [Tkinter](https://docs.python.org/3/library/tkinter.html)
- Package management with [uv](https://github.com/astral-sh/uv)

## Support

For issues, questions, or feature requests:
- [Open an issue](https://github.com/yourusername/HandWriting/issues)
- Check existing documentation
- Review closed issues for solutions

## Changelog

### Version 0.1.0 (Current)
- Initial release
- Command-line interface for document processing
- Batch processing with async API calls
- Dynamic token management to prevent truncation
- CSV and JSON export formats
- Image detection in table cells
- Support for Azure OpenAI and standard OpenAI
- Folder batch processing support

---

**Note**: Remember to replace `yourusername` with your actual GitHub username in all URLs.
