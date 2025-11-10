# Usage Guide - DOCX Document Extractor

Complete guide for using the command-line tool to extract structured information from DOCX documents.

## Table of Contents

- [Basic Setup](#basic-setup)
- [Simple Usage](#simple-usage)
- [Batch Processing](#batch-processing)
- [Configuration](#configuration)
- [Advanced Usage](#advanced-usage)
- [Tips & Best Practices](#tips--best-practices)

## Basic Setup

### 1. Install Dependencies

```bash
# Install using uv (recommended)
uv sync

# Or using pip
pip install -e .
```

### 2. Configure API Access

Create a `.env` file in the project directory:

```env
# For Azure OpenAI (default)
USE_AZURE_OPENAI=true
AZURE_OPENAI_API_KEY=your-api-key-here
AZURE_OPENAI_ENDPOINT=https://aidt-eastus2-ai.openai.azure.com
AZURE_OPENAI_API_VERSION=2025-03-01-preview
AZURE_OPENAI_DEPLOYMENT=gpt-5-chat
```

Or set environment variables directly:

```bash
# On Linux/macOS
export AZURE_OPENAI_API_KEY="your-api-key-here"

# On Windows
set AZURE_OPENAI_API_KEY=your-api-key-here
```

## Simple Usage

### Process a Single File

```bash
# Standard mode - process entire document at once
uv run python main.py document.docx -o output.json

# Output to CSV
uv run python main.py document.docx -o output.csv
```

### View Help

```bash
uv run python main.py --help
```

## Batch Processing

Batch mode is **recommended** for large documents as it:
- Processes documents faster with concurrent API calls
- Prevents output truncation with dynamic token management
- Shows real-time progress

### Basic Batch Processing

```bash
uv run python main.py document.docx --batch-mode -o output.csv
```

### Custom Batch Settings

```bash
uv run python main.py document.docx \
  --batch-mode \
  --batch-size 15 \          # Process 15 rows per batch
  --max-concurrent 8 \        # Use 8 concurrent API calls
  --max-tokens 25000 \        # Allow up to 25000 tokens per batch
  -o results.csv
```

### Process Multiple Files

```bash
# Process all DOCX files in a folder
uv run python main.py ./documents --batch-mode -o combined_results.csv
```

## Configuration

### Command-Line Options

| Option | Description | Default |
|--------|-------------|---------|
| `--batch-mode` | Enable batch processing | Disabled |
| `--batch-size N` | Rows per batch | 10 |
| `--max-concurrent N` | Max concurrent API calls | 5 |
| `--max-tokens N` | Max input tokens per batch | 20000 |
| `-o, --output FILE` | Output file (.json or .csv) | None |
| `--api-key KEY` | API key (overrides env var) | From env |
| `--model NAME` | Model/deployment name | From env |
| `--to-markdown` | Convert to Markdown (no API) | - |
| `--json` | Print JSON to stdout | - |

### Adjusting for Your Use Case

**For Large Documents:**
```bash
# Increase token limit, reduce batch size
uv run python main.py large_doc.docx \
  --batch-mode \
  --batch-size 5 \
  --max-tokens 30000
```

**For Fast Processing:**
```bash
# Increase concurrency (watch your rate limits!)
uv run python main.py document.docx \
  --batch-mode \
  --max-concurrent 15
```

**For API Rate Limits:**
```bash
# Reduce concurrency
uv run python main.py document.docx \
  --batch-mode \
  --max-concurrent 2
```

## Advanced Usage

### Convert to Markdown (No API Required)

```bash
# Convert tables to Markdown format
uv run python main.py document.docx --to-markdown -o output.md
```

This is useful for:
- Previewing document structure
- Testing without API costs
- Creating documentation

### Custom API Configuration

```bash
# Use custom endpoint
uv run python main.py document.docx \
  --api-key "your-key" \
  --api-base "https://custom-endpoint.com" \
  --model "custom-model"
```

### Output to Console

```bash
# Print JSON to stdout (for piping)
uv run python main.py document.docx --json

# Process and save, then display
uv run python main.py document.docx --batch-mode -o output.csv --json
```

### Using with Standard OpenAI

```bash
# Set in .env file
USE_AZURE_OPENAI=false
OPENAI_API_KEY=your-openai-key
OPENAI_MODEL=gpt-4o-2024-08-06

# Then run normally
uv run python main.py document.docx --batch-mode -o output.csv
```

## Tips & Best Practices

### 1. Start with Batch Mode

Always use `--batch-mode` for production:
```bash
uv run python main.py document.docx --batch-mode -o output.csv
```

### 2. Monitor Token Usage

The tool shows token estimates for each batch:
```
Batch 1: 8 rows, ~15234 tokens
Batch 2: 6 rows, ~18956 tokens
```

If batches are close to your token limit, adjust `--max-tokens`.

### 3. Adjust Concurrency Based on Rate Limits

Azure OpenAI rate limits vary by tier:
- **Free/Basic**: Use `--max-concurrent 2-3`
- **Standard**: Use `--max-concurrent 5-10`
- **Enterprise**: Use `--max-concurrent 15-20`

### 4. Use CSV for Multiple Files

When processing folders, CSV format consolidates all results:
```bash
uv run python main.py ./documents --batch-mode -o all_results.csv
```

### 5. Test Without API First

Use Markdown conversion to verify document structure:
```bash
uv run python main.py document.docx --to-markdown
```

### 6. Handle Errors Gracefully

The tool continues processing even if some files fail:
```bash
# Folder processing skips failed files and continues
uv run python main.py ./documents --batch-mode -o results.csv
```

Check the console output for warnings about failed files.

## Understanding the Output

### CSV Format

```csv
Source File,TL EA,Test Standard,Test Analytes,PP Notes,Source Link,Label and Symbol
doc1.docx,Physical Tests,EN 71-1,Lead,"Shall comply...",http://example.com,yes
```

Each row represents one extracted record from the document tables.

### JSON Format

```json
{
  "records": [
    {
      "tl_ea": "Physical Tests",
      "test_standard": "EN 71-1",
      "test_analytes": "Lead",
      "pp_notes": "Shall comply with requirements",
      "source_link": "http://example.com",
      "label_and_symbol": "yes"
    }
  ]
}
```

### Progress Output

```
Reading file: document.docx
  Batch 1: 8 rows, ~15234 tokens
  Batch 2: 6 rows, ~18956 tokens
Document split into 2 batch(es)
Processing with up to 5 concurrent API calls...

Processing batches: 100%|████████| 2/2 [00:10<00:00,  5.23s/batch]

Extraction complete!
Total records extracted: 25
```

## Troubleshooting

### "Connection Error"

**Cause**: Network issue or invalid API key

**Solution**:
```bash
# Verify API key is set
echo $AZURE_OPENAI_API_KEY

# Test with a simple request
uv run python main.py small_doc.docx -o test.json
```

### "Token Limit Exceeded"

**Cause**: Batch contains too much text

**Solution**:
```bash
# Reduce batch size or increase token limit
uv run python main.py document.docx \
  --batch-mode \
  --batch-size 5 \
  --max-tokens 30000
```

### "Rate Limit Error"

**Cause**: Too many concurrent requests

**Solution**:
```bash
# Reduce concurrency
uv run python main.py document.docx \
  --batch-mode \
  --max-concurrent 2
```

### Empty Output

**Causes**:
1. Document has no tables
2. Tables don't match expected structure
3. API returned empty responses

**Debug**:
```bash
# Check document structure first
uv run python main.py document.docx --to-markdown

# Enable verbose output
uv run python main.py document.docx --batch-mode -o output.csv --json
```

## Example Workflows

### Workflow 1: Single Document

```bash
# 1. Check document structure
uv run python main.py report.docx --to-markdown | head -20

# 2. Process with defaults
uv run python main.py report.docx --batch-mode -o report_data.csv

# 3. Review results
cat report_data.csv
```

### Workflow 2: Batch Processing Multiple Files

```bash
# 1. Set API key
export AZURE_OPENAI_API_KEY="your-key"

# 2. Process all documents
uv run python main.py ./documents --batch-mode -o all_documents.csv

# 3. Check results
wc -l all_documents.csv  # Count rows
head all_documents.csv   # Preview
```

### Workflow 3: High-Volume Processing

```bash
# Optimize for speed with high concurrency
uv run python main.py ./large_batch \
  --batch-mode \
  --batch-size 20 \
  --max-concurrent 10 \
  --max-tokens 30000 \
  -o results.csv
```

## Support

For additional help:
- Run `uv run python main.py --help`
- Check [README.md](README.md)
- Open an issue on GitHub

---

**Pro Tip**: Save common configurations as shell aliases:
```bash
# Add to ~/.bashrc or ~/.zshrc
alias docx-extract='uv run python /path/to/main.py --batch-mode'

# Then use:
docx-extract document.docx -o output.csv
```
