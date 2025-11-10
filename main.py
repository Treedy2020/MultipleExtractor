"""DOCX document information extraction tool.

This module extracts structured information from DOCX documents using OpenAI or
Azure OpenAI structured outputs. It supports extracting multiple records from a
single document (e.g., multiple rows in a table) and batch processing multiple
documents in a folder.

Features:
    - Single file processing: Extract structured information from a single DOCX
    - Batch processing: Process all DOCX files in a folder
    - Multi-format export: Support JSON and CSV format outputs
    - Multi-record extraction: Automatically identify multiple rows of data
    - Dual API support: Support both Azure OpenAI and standard OpenAI API

Typical usage examples:
    # Use Azure OpenAI (default)
    python main.py document.docx
    python main.py document.docx -o output.json

    # Batch process folder and export to CSV
    python main.py ./documents -o results.csv

    # Use custom Azure configuration
    python main.py document.docx --api-key your-azure-key --model your-deployment-name

    # Use standard OpenAI API
    set USE_AZURE_OPENAI=false
    python main.py document.docx --api-key your-openai-key

Environment variables (Azure OpenAI - enabled by default):
    USE_AZURE_OPENAI: Whether to use Azure OpenAI (default: true)
    AZURE_OPENAI_API_KEY: Azure OpenAI API key
    AZURE_OPENAI_ENDPOINT: Azure OpenAI endpoint URL
    AZURE_OPENAI_API_VERSION: Azure OpenAI API version (default: 2024-08-01-preview)
    AZURE_OPENAI_DEPLOYMENT: Azure OpenAI deployment name (default: gpt-4o)

Environment variables (Standard OpenAI - when USE_AZURE_OPENAI=false):
    OPENAI_API_KEY: OpenAI API key
    OPENAI_API_BASE: OpenAI API base URL
    OPENAI_MODEL: Model name to use (default: gpt-4o-2024-08-06)

Note:
    - Each row of data in the document will be extracted as a separate record
    - When batch processing folders, CSV format is recommended for consolidation
    - CSV output includes source filename for tracking record origins
    - Azure OpenAI is used by default, set USE_AZURE_OPENAI=false to switch to standard OpenAI
"""

import argparse
import asyncio
import csv
import json
import os
import sys
from pathlib import Path
from typing import Optional

from docx import Document
from openai import AzureOpenAI, OpenAI, AsyncAzureOpenAI, AsyncOpenAI
from pydantic import BaseModel, Field
from tqdm import tqdm


class DocumentFields(BaseModel):
    """Structured model for document fields.

    Attributes:
        tl_ea: TL EA information from Column 1 of attached protocol.
        test_standard: Test standard from Column 2 (excluding website links).
        test_analytes: Test analytes from Column 5.
        pp_notes: PP notes information from Column 3.
        source_link: Website link from Column 2 if found, otherwise None.
        label_and_symbol: Whether any label is found in this row (yes/no).
    """
    tl_ea: str = Field(description="Column 1 of attached protocol - TL EA information")
    test_standard: str = Field(description="Column 2 but not website - test standard (excluding website links)")
    test_analytes: str = Field(description="Column 5 - test analytes")
    pp_notes: str = Field(description="Column 3 - PP notes information")
    source_link: Optional[str] = Field(default=None, description="Column 2 website if found - source link")
    label_and_symbol: str = Field(description="Any label found in this row, just state yes/no")


class DocumentExtraction(BaseModel):
    """Document extraction result model containing multiple records.

    Attributes:
        records: List of all records extracted from the document.
    """
    records: list[DocumentFields] = Field(
        description="All records extracted from the document, each record corresponds to a row of data"
    )


class DocxExtractor:
    """DOCX document extractor for extracting structured information from Word documents.

    This class uses OpenAI or Azure OpenAI structured output functionality to
    extract specific fields from DOCX documents.

    Attributes:
        client: OpenAI or AzureOpenAI client instance.
        model: Name of the OpenAI model or Azure deployment to use.
        use_azure: Whether Azure OpenAI is being used.
    """

    def __init__(
        self,
        api_key: str,
        model: str = "gpt-5-chat",
        api_base: Optional[str] = None,
        use_azure: bool = False,
        azure_endpoint: Optional[str] = None,
        azure_api_version: Optional[str] = None
    ):
        """Initializes a DocxExtractor instance.

        Args:
            api_key: OpenAI or Azure OpenAI API key.
            model: Model name or Azure deployment name. Defaults to "gpt-4o-2024-08-06".
            api_base: API base URL. Defaults to None (uses OpenAI official address).
            use_azure: Whether to use Azure OpenAI. Defaults to False.
            azure_endpoint: Azure OpenAI endpoint URL (required when use_azure=True).
            azure_api_version: Azure OpenAI API version (required when use_azure=True).

        Raises:
            openai.OpenAIError: If the API key is invalid.
            ValueError: If using Azure but missing required parameters.
        """
        if use_azure:
            if not azure_endpoint:
                raise ValueError("azure_endpoint is required when using Azure OpenAI")
            if not azure_api_version:
                raise ValueError("azure_api_version is required when using Azure OpenAI")

            self.client = AzureOpenAI(
                api_key=api_key,
                azure_endpoint=azure_endpoint,
                api_version=azure_api_version
            )
            # Create async client for batch processing
            self.async_client = AsyncAzureOpenAI(
                api_key=api_key,
                azure_endpoint=azure_endpoint,
                api_version=azure_api_version
            )
        else:
            client_kwargs = {"api_key": api_key}
            if api_base:
                client_kwargs["base_url"] = api_base

            self.client = OpenAI(**client_kwargs)
            # Create async client for batch processing
            self.async_client = AsyncOpenAI(**client_kwargs)

        self.model = model
        self.use_azure = use_azure

    def read_docx(self, file_path: str) -> str:
        """Reads a DOCX file and extracts all text content.

        This method extracts paragraph text and table content from the document,
        combining them into a single string. It also detects and marks cells
        that contain images.

        Args:
            file_path: Path to the DOCX file (relative or absolute).

        Returns:
            A string containing all document content, with paragraphs and tables
            separately annotated.

        Raises:
            FileNotFoundError: If the specified file does not exist.
            docx.opc.exceptions.PackageNotFoundError: If the file is not a valid DOCX format.
        """
        doc = Document(file_path)

        # Extract all paragraph text
        paragraphs = [para.text for para in doc.paragraphs if para.text.strip()]

        # Extract table content with image detection
        tables_content = []
        for table in doc.tables:
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    # Check if cell contains images
                    if self._cell_has_images(cell):
                        # Append image marker to cell content
                        if cell_text:
                            cell_text += " [IMAGE: Contains label/symbol image]"
                        else:
                            cell_text = "[IMAGE: Contains label/symbol image]"
                    row_data.append(cell_text)
                if any(row_data):  # Only add non-empty rows
                    tables_content.append(" | ".join(row_data))

        # Combine all content
        full_text = "\n".join(paragraphs)
        if tables_content:
            full_text += "\n\n=== Table Content ===\n" + "\n".join(tables_content)

        return full_text

    def _cell_has_images(self, cell) -> bool:
        """Check if a table cell contains any images.

        Args:
            cell: A docx table cell object.

        Returns:
            True if the cell contains one or more images, False otherwise.
        """
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                # Check if the run contains any drawing elements (images)
                if run._element.xpath('.//w:drawing'):
                    return True
                # Also check for inline shapes (another way images can be embedded)
                if run._element.xpath('.//w:pict'):
                    return True
        return False

    @staticmethod
    def estimate_tokens(text: str) -> int:
        """Estimates the number of tokens in a text string.

        This uses a simple heuristic:
        - For ASCII characters: ~4 characters per token
        - For non-ASCII (including Chinese): ~2 characters per token

        Args:
            text: The text to estimate tokens for.

        Returns:
            Estimated number of tokens.
        """
        if not text:
            return 0

        ascii_chars = sum(1 for c in text if ord(c) < 128)
        non_ascii_chars = len(text) - ascii_chars

        # Rough estimation: 4 chars/token for ASCII, 2 chars/token for non-ASCII
        estimated_tokens = (ascii_chars / 4.0) + (non_ascii_chars / 2.0)

        return int(estimated_tokens) + 1  # Add 1 for safety margin

    def read_docx_in_batches(
        self,
        file_path: str,
        batch_size: int = 10,
        max_tokens_per_batch: int = 20000
    ) -> list[str]:
        """Reads a DOCX file and splits table content into batches by rows with dynamic token window.

        This method extracts table content and splits it into multiple batches.
        It uses a dynamic window strategy that considers both row count and token limits
        to ensure the output doesn't get truncated due to exceeding token limits.

        Args:
            file_path: Path to the DOCX file (relative or absolute).
            batch_size: Maximum number of data rows per batch. Defaults to 10.
                This serves as an upper limit; actual batch size may be smaller
                if token limit is reached.
            max_tokens_per_batch: Maximum tokens per batch input. Defaults to 20000.
                This ensures enough room for model output (typically ~4-8k tokens).

        Returns:
            A list of strings, each containing a batch of table content with header.

        Raises:
            FileNotFoundError: If the specified file does not exist.
            docx.opc.exceptions.PackageNotFoundError: If the file is not a valid DOCX format.
        """
        doc = Document(file_path)
        batches = []

        # Extract paragraph text (will be prepended to each batch)
        paragraphs = [para.text for para in doc.paragraphs if para.text.strip()]
        paragraph_text = "\n".join(paragraphs) if paragraphs else ""

        # Reserve tokens for system prompt (~2000), paragraph text, and formatting
        system_prompt_tokens = 2000
        paragraph_tokens = self.estimate_tokens(paragraph_text)
        reserved_tokens = system_prompt_tokens + paragraph_tokens + 500  # 500 for formatting

        # Process each table
        for table_idx, table in enumerate(doc.tables):
            if len(table.rows) == 0:
                continue

            # Extract header row with image detection
            header_row = table.rows[0]
            header_data = []
            for cell in header_row.cells:
                cell_text = cell.text.strip()
                if self._cell_has_images(cell):
                    if cell_text:
                        cell_text += " [IMAGE: Contains label/symbol image]"
                    else:
                        cell_text = "[IMAGE: Contains label/symbol image]"
                header_data.append(cell_text)

            header_line = " | ".join(header_data)
            header_tokens = self.estimate_tokens(header_line)

            # Extract data rows with image detection
            data_rows = []
            data_row_tokens = []
            for row in table.rows[1:]:  # Skip header
                row_data = []
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if self._cell_has_images(cell):
                        if cell_text:
                            cell_text += " [IMAGE: Contains label/symbol image]"
                        else:
                            cell_text = "[IMAGE: Contains label/symbol image]"
                    row_data.append(cell_text)
                if any(row_data):  # Only add non-empty rows
                    row_line = " | ".join(row_data)
                    data_rows.append(row_line)
                    data_row_tokens.append(self.estimate_tokens(row_line))

            # Dynamic batching based on token count
            batch_idx = 0
            i = 0
            while i < len(data_rows):
                batch_data_rows = []
                current_batch_tokens = reserved_tokens + header_tokens

                # Add rows until we hit batch_size or token limit
                rows_in_batch = 0
                while i < len(data_rows) and rows_in_batch < batch_size:
                    row_tokens = data_row_tokens[i]

                    # Check if adding this row would exceed token limit
                    if current_batch_tokens + row_tokens > max_tokens_per_batch:
                        # If this is the first row in batch, we must include it anyway
                        if rows_in_batch == 0:
                            batch_data_rows.append(data_rows[i])
                            i += 1
                            print(f"Warning: Single row exceeds token limit ({row_tokens} tokens)")
                        break

                    batch_data_rows.append(data_rows[i])
                    current_batch_tokens += row_tokens
                    rows_in_batch += 1
                    i += 1

                # Construct batch content
                if batch_data_rows:
                    batch_content = []
                    if paragraph_text:
                        batch_content.append(paragraph_text)

                    batch_content.append(f"\n=== Table {table_idx + 1} (Batch {batch_idx + 1}) ===")
                    batch_content.append(header_line)
                    batch_content.extend(batch_data_rows)

                    batch_text = "\n".join(batch_content)
                    batches.append(batch_text)

                    # Log batch info
                    actual_tokens = self.estimate_tokens(batch_text)
                    print(f"  Batch {batch_idx + 1}: {rows_in_batch} rows, ~{actual_tokens} tokens")

                    batch_idx += 1

        # If no tables found, return the whole document as one batch
        if not batches:
            batches.append(self.read_docx(file_path))

        return batches

    def convert_tables_to_markdown(self, file_path: str) -> str:
        """Converts all tables in a DOCX file to Markdown format.

        This method extracts all tables from the document and converts them
        to Markdown table syntax. Non-table content is preserved as plain text.
        It also detects and marks cells that contain images.

        Args:
            file_path: Path to the DOCX file (relative or absolute).

        Returns:
            A string containing the document content with tables in Markdown format.

        Raises:
            FileNotFoundError: If the specified file does not exist.
            docx.opc.exceptions.PackageNotFoundError: If the file is not a valid DOCX format.
        """
        doc = Document(file_path)
        output = []

        # Extract paragraphs
        for para in doc.paragraphs:
            if para.text.strip():
                output.append(para.text)
                output.append("")  # Add blank line

        # Extract and convert tables to Markdown
        for table_idx, table in enumerate(doc.tables, 1):
            output.append(f"### Table {table_idx}")
            output.append("")

            # Get table data with image detection
            table_data = []
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    # Check if cell contains images
                    if self._cell_has_images(cell):
                        # Append image marker to cell content
                        if cell_text:
                            cell_text += " [IMAGE: Contains label/symbol image]"
                        else:
                            cell_text = "[IMAGE: Contains label/symbol image]"
                    row_data.append(cell_text)
                table_data.append(row_data)

            if not table_data:
                continue

            # Assume first row is header
            if len(table_data) > 0:
                header = table_data[0]
                num_cols = len(header)

                # Write header
                output.append("| " + " | ".join(header) + " |")

                # Write separator
                output.append("| " + " | ".join(["---"] * num_cols) + " |")

                # Write data rows
                for row in table_data[1:]:
                    # Pad row if it has fewer columns than header
                    while len(row) < num_cols:
                        row.append("")
                    output.append("| " + " | ".join(row[:num_cols]) + " |")

                output.append("")  # Add blank line after table

        return "\n".join(output)

    def extract_fields(self, text: str) -> DocumentExtraction:
        """Extracts document fields using OpenAI structured output API.

        This method sends document text to the OpenAI API and uses structured
        output functionality to extract predefined fields. A document may contain
        multiple rows of data, each extracted as a DocumentFields record.

        Args:
            text: Text content extracted from the DOCX document.

        Returns:
            A DocumentExtraction instance containing all extracted records.

        Raises:
            openai.APIError: If the API call fails.
            openai.RateLimitError: If the API rate limit is exceeded.
        """
        completion = self.client.responses.parse(
            model=self.model,
            input=[
                {
                    "role": "system",
                    "content": """You are a professional document information extraction assistant. Please extract the following fields from the provided document content:
                    1. TL EA: Extract the attached protocol information from Column 1
                       - Extract ALL content completely and accurately from Column 1
                       - Do not omit, summarize, or truncate any text
                    2. Test standard: Extract non-website content from Column 2 (test standard)
                       - Extract ALL text content from Column 2 completely
                       - Exclude only website URLs (which should go to Source link field)
                       - Do not omit, summarize, or truncate any text
                    3. Test analytes: Extract test analyte information ONLY from Column 5
                       - IMPORTANT: Only extract content that appears in Column 5
                       - Do NOT extract test analytes or chemical names from Column 3 (PP notes) or any other columns
                       - If Column 5 is empty, return an empty string
                       - Only use the exact content from Column 5, do not infer or analyze from other columns
                       - Extract ALL content from Column 5 completely when present
                    4. PP notes: Extract notes information from Column 3
                       - CRITICAL: Extract ALL content from Column 3 completely and accurately
                       - Do NOT omit any text even if it contains chemical names or test information
                       - Do NOT summarize or truncate the content
                       - Column 3 often contains detailed requirements, standards, and notes - preserve everything
                    5. Source link: If there is a website link in Column 2, extract it; otherwise return null
                    6. Label and symbol: This field indicates whether this row contains any label or symbol images (such as certification marks, safety labels, warning symbols, etc.).
                       - In the document, cells containing images will be marked with "[IMAGE: Contains label/symbol image]"
                       - If you see this marker in any cell of the current row, return "yes"
                       - If no such marker is found in the row, return "no"
                       - Note: The label/symbol refers to visual graphics or icons embedded in the document, not just text descriptions

                    Important notes:
                    - The document may contain multiple rows of data (e.g., multiple rows in a table)
                    - Please create a separate record for each row of data
                    - Put all records in the records list
                    - Please carefully analyze the document content and accurately extract this information
                    - Pay special attention to the "[IMAGE: Contains label/symbol image]" markers to determine the Label and symbol field
                    - CRITICAL: Test analytes must ONLY come from Column 5, never from other columns
                    - CRITICAL: All other columns (1, 2, 3) must be extracted COMPLETELY without omission, summarization, or truncation"""
                },
                {
                    "role": "user",
                    "content": f"Please extract information from all rows in the following document content:\n\n{text}"
                }
            ],
            text_format=DocumentExtraction,
        )

        return completion.output_parsed

    async def extract_fields_async(self, text: str) -> DocumentExtraction:
        """Asynchronously extracts document fields using OpenAI structured output API.

        This is the async version of extract_fields, used for batch processing.

        Args:
            text: Text content extracted from the DOCX document.

        Returns:
            A DocumentExtraction instance containing all extracted records.

        Raises:
            openai.APIError: If the API call fails.
            openai.RateLimitError: If the API rate limit is exceeded.
        """
        completion = await self.async_client.responses.parse(
            model=self.model,
            input=[
                {
                    "role": "system",
                    "content": """You are a professional document information extraction assistant. Please extract the following fields from the provided document content:
                    1. TL EA: Extract the attached protocol information from Column 1
                       - Extract ALL content completely and accurately from Column 1
                       - Do not omit, summarize, or truncate any text
                    2. Test standard: Extract non-website content from Column 2 (test standard)
                       - Extract ALL text content from Column 2 completely
                       - Exclude only website URLs (which should go to Source link field)
                       - Do not omit, summarize, or truncate any text
                    3. Test analytes: Extract test analyte information ONLY from Column 5
                       - IMPORTANT: Only extract content that appears in Column 5
                       - Do NOT extract test analytes or chemical names from Column 3 (PP notes) or any other columns
                       - If Column 5 is empty, return an empty string
                       - Only use the exact content from Column 5, do not infer or analyze from other columns
                       - Extract ALL content from Column 5 completely when present
                    4. PP notes: Extract notes information from Column 3
                       - CRITICAL: Extract ALL content from Column 3 completely and accurately
                       - Do NOT omit any text even if it contains chemical names or test information
                       - Do NOT summarize or truncate the content
                       - Column 3 often contains detailed requirements, standards, and notes - preserve everything
                    5. Source link: If there is a website link in Column 2, extract it; otherwise return null
                    6. Label and symbol: This field indicates whether this row contains any label or symbol images (such as certification marks, safety labels, warning symbols, etc.).
                       - In the document, cells containing images will be marked with "[IMAGE: Contains label/symbol image]"
                       - If you see this marker in any cell of the current row, return "yes"
                       - If no such marker is found in the row, return "no"
                       - Note: The label/symbol refers to visual graphics or icons embedded in the document, not just text descriptions

                    Important notes:
                    - The document may contain multiple rows of data (e.g., multiple rows in a table)
                    - Please create a separate record for each row of data
                    - Put all records in the records list
                    - Please carefully analyze the document content and accurately extract this information
                    - Pay special attention to the "[IMAGE: Contains label/symbol image]" markers to determine the Label and symbol field
                    - CRITICAL: Test analytes must ONLY come from Column 5, never from other columns
                    - CRITICAL: All other columns (1, 2, 3) must be extracted COMPLETELY without omission, summarization, or truncation"""
                },
                {
                    "role": "user",
                    "content": f"Please extract information from all rows in the following document content:\n\n{text}"
                }
            ],
            text_format=DocumentExtraction,
        )

        return completion.output_parsed

    async def process_batches_async(
        self,
        batches: list[str],
        max_concurrent: int = 5
    ) -> DocumentExtraction:
        """Processes multiple batches of document content asynchronously with concurrency control.

        This method processes multiple batches concurrently using asyncio, with a limit
        on the number of concurrent requests to avoid overwhelming the API.

        Args:
            batches: List of text batches to process.
            max_concurrent: Maximum number of concurrent API calls. Defaults to 5.

        Returns:
            A DocumentExtraction instance containing all records from all batches,
            maintaining the order of batches.

        Raises:
            openai.APIError: If any API call fails.
        """
        semaphore = asyncio.Semaphore(max_concurrent)

        async def process_with_semaphore(batch_text: str, batch_idx: int) -> tuple[int, DocumentExtraction]:
            async with semaphore:
                result = await self.extract_fields_async(batch_text)
                return batch_idx, result

        # Create tasks for all batches with progress tracking
        tasks = [process_with_semaphore(batch, idx) for idx, batch in enumerate(batches)]

        # Execute all tasks with progress bar
        all_records = []
        with tqdm(total=len(batches), desc="Processing batches", unit="batch") as pbar:
            for coro in asyncio.as_completed(tasks):
                batch_idx, extraction = await coro
                all_records.append((batch_idx, extraction))
                pbar.update(1)

        # Sort by batch index to maintain order
        all_records.sort(key=lambda x: x[0])

        # Combine all records from all batches
        combined_records = []
        for _, extraction in all_records:
            combined_records.extend(extraction.records)

        return DocumentExtraction(records=combined_records)

    @staticmethod
    def export_to_csv(extractions: list[tuple[str, DocumentExtraction]], output_path: str) -> None:
        """Exports multiple extraction results to a CSV file.

        Args:
            extractions: List of (filename, DocumentExtraction) tuples.
            output_path: Output CSV file path.

        Raises:
            IOError: If unable to write to the file.
        """
        with open(output_path, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)

            # Write header
            writer.writerow([
                'Source File',
                'TL EA',
                'Test Standard',
                'Test Analytes',
                'PP Notes',
                'Source Link',
                'Label and Symbol'
            ])

            # Write all records from each file
            for filename, extraction in extractions:
                for record in extraction.records:
                    writer.writerow([
                        filename,
                        record.tl_ea,
                        record.test_standard,
                        record.test_analytes,
                        record.pp_notes,
                        record.source_link or '',
                        record.label_and_symbol
                    ])

    def process_file(self, file_path: str, output_path: Optional[str] = None) -> DocumentExtraction:
        """Processes a DOCX file and extracts structured information.

        This is the main workflow method that reads the DOCX file, extracts fields,
        and optionally saves the results to a JSON file. Processing progress and
        results are printed to standard output. The document may contain multiple
        rows of data, each of which will be extracted.

        Args:
            file_path: Input DOCX file path (relative or absolute).
            output_path: Optional output JSON file path. If provided, results will
                be saved in JSON format.

        Returns:
            A DocumentExtraction instance containing all extracted records.

        Raises:
            FileNotFoundError: If the input file does not exist.
            openai.APIError: If the OpenAI API call fails.
            IOError: If unable to write to the output file.
        """
        print(f"Reading file: {file_path}")
        text = self.read_docx(file_path)

        print(f"Document content length: {len(text)} characters")
        print("\nExtracting structured information using OpenAI...")

        extraction = self.extract_fields(text)

        print("\nExtraction complete!")
        print("=" * 80)
        print(f"Total records extracted: {len(extraction.records)}\n")

        for idx, record in enumerate(extraction.records, 1):
            print(f"Record #{idx}")
            print("-" * 80)
            print(f"  TL EA:           {record.tl_ea}")
            print(f"  Test standard:   {record.test_standard}")
            print(f"  Test analytes:   {record.test_analytes}")
            print(f"  PP notes:        {record.pp_notes}")
            print(f"  Source link:     {record.source_link}")
            print(f"  Label & symbol:  {record.label_and_symbol}")
            print()

        print("=" * 80)

        # Save as JSON if output path is specified
        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(extraction.model_dump(), f, ensure_ascii=False, indent=2)
            print(f"\nResults saved to: {output_path}")

        return extraction

    def process_file_with_batches(
        self,
        file_path: str,
        batch_size: int = 10,
        max_concurrent: int = 5,
        max_tokens_per_batch: int = 20000,
        output_path: Optional[str] = None
    ) -> DocumentExtraction:
        """Processes a DOCX file using batch processing with async API calls.

        This method splits the document into batches by table rows, then processes
        all batches concurrently using async API calls. Progress is shown via a
        progress bar. Uses dynamic token windowing to prevent output truncation.

        Args:
            file_path: Input DOCX file path (relative or absolute).
            batch_size: Maximum number of table rows per batch. Defaults to 10.
                Actual batch size may be smaller if token limit is reached.
            max_concurrent: Maximum number of concurrent API calls. Defaults to 5.
            max_tokens_per_batch: Maximum input tokens per batch. Defaults to 20000.
                This ensures enough room for model output without truncation.
            output_path: Optional output file path (.json or .csv). If provided, results
                will be saved in the specified format based on file extension.

        Returns:
            A DocumentExtraction instance containing all extracted records from all batches.

        Raises:
            FileNotFoundError: If the input file does not exist.
            openai.APIError: If the OpenAI API call fails.
            IOError: If unable to write to the output file.
        """
        print(f"Reading file: {file_path}")
        batches = self.read_docx_in_batches(file_path, batch_size, max_tokens_per_batch)

        print(f"Document split into {len(batches)} batch(es)")
        print(f"Processing with up to {max_concurrent} concurrent API calls...\n")

        # Run async processing
        extraction = asyncio.run(self.process_batches_async(batches, max_concurrent))

        print("\n\nExtraction complete!")
        print("=" * 80)
        print(f"Total records extracted: {len(extraction.records)}\n")

        for idx, record in enumerate(extraction.records, 1):
            print(f"Record #{idx}")
            print("-" * 80)
            print(f"  TL EA:           {record.tl_ea}")
            print(f"  Test standard:   {record.test_standard}")
            print(f"  Test analytes:   {record.test_analytes}")
            print(f"  PP notes:        {record.pp_notes}")
            print(f"  Source link:     {record.source_link}")
            print(f"  Label & symbol:  {record.label_and_symbol}")
            print()

        print("=" * 80)

        # Save based on file extension
        if output_path:
            output_path_obj = Path(output_path)
            if output_path_obj.suffix.lower() == '.csv':
                # Export to CSV - use filename from file_path
                filename = Path(file_path).name
                DocxExtractor.export_to_csv([(filename, extraction)], output_path)
                print(f"\nResults saved to CSV: {output_path}")
            else:
                # Default to JSON
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(extraction.model_dump(), f, ensure_ascii=False, indent=2)
                print(f"\nResults saved to JSON: {output_path}")

        return extraction


def parse_args() -> argparse.Namespace:
    """Parses command line arguments.

    Returns:
        A Namespace object containing the parsed arguments.
    """
    parser = argparse.ArgumentParser(
        description="Extract structured information from DOCX documents, supports single file or batch folder processing",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process a single file
  %(prog)s document.docx
  %(prog)s document.docx -o output.json

  # Process all DOCX files in a folder and export to CSV
  %(prog)s ./documents -o results.csv

  # Use batch mode for concurrent processing with progress bar
  %(prog)s document.docx --batch-mode --batch-size 10 --max-concurrent 5
  %(prog)s ./documents --batch-mode -o results.csv

  # Convert DOCX tables to Markdown (no AI processing)
  %(prog)s document.docx --to-markdown
  %(prog)s document.docx --to-markdown -o output.md

  # Use custom API configuration
  %(prog)s document.docx --api-key your-api-key --model gpt-4o
  %(prog)s document.docx --api-base https://api.openai.com/v1

Batch Mode:
  When --batch-mode is enabled, the tool splits tables by rows and processes
  them concurrently with async API calls. This is useful for large documents
  with many table rows, as it can significantly speed up processing while
  maintaining consistent table headers across all batches. Progress is shown
  via a progress bar.

Environment variables (Azure OpenAI - enabled by default):
  USE_AZURE_OPENAI           Whether to use Azure OpenAI (default: true)
  AZURE_OPENAI_API_KEY       Azure OpenAI API key
  AZURE_OPENAI_ENDPOINT      Azure OpenAI endpoint URL
  AZURE_OPENAI_API_VERSION   Azure OpenAI API version (default: 2024-08-01-preview)
  AZURE_OPENAI_DEPLOYMENT    Azure OpenAI deployment name (default: gpt-4o)

Environment variables (Standard OpenAI - when USE_AZURE_OPENAI=false):
  OPENAI_API_KEY     OpenAI API key (if not specified via --api-key)
  OPENAI_API_BASE    OpenAI API base URL (if not specified via --api-base)
  OPENAI_MODEL       OpenAI model name (if not specified via --model, default: gpt-4o-2024-08-06)
        """
    )

    parser.add_argument(
        "input_path",
        type=str,
        help="Input DOCX file path or folder path containing DOCX files"
    )

    parser.add_argument(
        "-o", "--output",
        type=str,
        default=None,
        help="Output file path (.json or .csv format). CSV format is recommended for folder processing"
    )

    parser.add_argument(
        "--api-key",
        type=str,
        default=None,
        help="API key (Azure mode: AZURE_OPENAI_API_KEY, OpenAI mode: OPENAI_API_KEY)"
    )

    parser.add_argument(
        "--api-base",
        type=str,
        default=None,
        help="API endpoint URL (Azure mode: AZURE_OPENAI_ENDPOINT, OpenAI mode: OPENAI_API_BASE)"
    )

    parser.add_argument(
        "--model",
        type=str,
        default=None,
        help="Model/deployment name (Azure mode: AZURE_OPENAI_DEPLOYMENT, OpenAI mode: OPENAI_MODEL)"
    )

    parser.add_argument(
        "--json",
        action="store_true",
        help="Output in JSON format to standard output"
    )

    parser.add_argument(
        "--to-markdown",
        action="store_true",
        help="Convert DOCX tables to Markdown format (no AI processing)"
    )

    parser.add_argument(
        "--batch-mode",
        action="store_true",
        help="Enable batch processing mode: split tables by rows and process concurrently"
    )

    parser.add_argument(
        "--batch-size",
        type=int,
        default=10,
        help="Number of table rows per batch (default: 10)"
    )

    parser.add_argument(
        "--max-concurrent",
        type=int,
        default=5,
        help="Maximum number of concurrent API calls (default: 5)"
    )

    parser.add_argument(
        "--max-tokens",
        type=int,
        default=20000,
        help="Maximum input tokens per batch (default: 20000). Prevents output truncation by limiting input size."
    )

    return parser.parse_args()


def main():
    """Main entry point for the CLI tool.

    Parses command line arguments, validates input, and processes DOCX files to
    extract structured information. Supports single file processing and batch
    folder processing. Results are printed to standard output and can optionally
    be saved to JSON or CSV files.

    Returns:
        Exit code (0 for success, 1 for failure).
    """
    args = parse_args()

    # Validate input path
    input_path = Path(args.input_path)
    if not input_path.exists():
        print(f"Error: Path does not exist - {args.input_path}", file=sys.stderr)
        return 1

    # Handle --to-markdown option (no API required)
    if args.to_markdown:
        if not input_path.is_file():
            print("Error: --to-markdown only supports single file conversion", file=sys.stderr)
            return 1

        if input_path.suffix.lower() not in ['.docx', '.doc']:
            print(f"Warning: File may not be in DOCX format - {args.input_path}", file=sys.stderr)

        try:
            # Create a temporary extractor just for conversion (no API needed)
            temp_extractor = DocxExtractor(
                api_key="dummy",  # Not used for markdown conversion
                use_azure=False
            )

            print(f"Converting tables to Markdown: {input_path}")
            markdown_content = temp_extractor.convert_tables_to_markdown(str(input_path))

            # Save to file or print to stdout
            if args.output:
                output_path = Path(args.output)
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(markdown_content)
                print(f"\nMarkdown content saved to: {output_path}")
            else:
                print("\n" + "=" * 80)
                print("Markdown Output:")
                print("=" * 80)
                # Handle encoding for console output
                try:
                    print(markdown_content)
                except UnicodeEncodeError:
                    # If console doesn't support UTF-8, encode with error handling
                    print(markdown_content.encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding))

            return 0
        except Exception as e:
            print(f"Error: Failed to convert to Markdown - {e}", file=sys.stderr)
            import traceback
            traceback.print_exc()
            return 1

    # Check whether to use Azure OpenAI (default is true)
    use_azure = os.getenv("USE_AZURE_OPENAI", "true").lower() in ("true", "1", "yes")

    if use_azure:
        # Get Azure OpenAI configuration
        api_key = args.api_key or os.getenv("AZURE_OPENAI_API_KEY")
        azure_endpoint = args.api_base or os.getenv("AZURE_OPENAI_ENDPOINT")
        azure_api_version = os.getenv("AZURE_OPENAI_API_VERSION", "2024-08-01-preview")
        model = args.model or os.getenv("AZURE_OPENAI_DEPLOYMENT") or "gpt-4o"

        if not api_key:
            print("Error: Azure OpenAI API key not provided", file=sys.stderr)
            print("Please provide the API key in one of the following ways:", file=sys.stderr)
            print("  1. Use the --api-key parameter", file=sys.stderr)
            print("  2. Set the AZURE_OPENAI_API_KEY environment variable", file=sys.stderr)
            print("     Example: set AZURE_OPENAI_API_KEY=your-api-key-here", file=sys.stderr)
            return 1

        if not azure_endpoint:
            print("Error: Azure OpenAI endpoint not provided", file=sys.stderr)
            print("Please provide the endpoint in one of the following ways:", file=sys.stderr)
            print("  1. Use the --api-base parameter", file=sys.stderr)
            print("  2. Set the AZURE_OPENAI_ENDPOINT environment variable", file=sys.stderr)
            print("     Example: set AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/", file=sys.stderr)
            return 1
    else:
        # Get standard OpenAI configuration
        api_key = args.api_key or os.getenv("OPENAI_API_KEY")
        api_base = args.api_base or os.getenv("OPENAI_API_BASE")
        model = args.model or os.getenv("OPENAI_MODEL") or "gpt-4o-2024-08-06"
        azure_endpoint = None
        azure_api_version = None

        if not api_key:
            print("Error: OpenAI API key not provided", file=sys.stderr)
            print("Please provide the API key in one of the following ways:", file=sys.stderr)
            print("  1. Use the --api-key parameter", file=sys.stderr)
            print("  2. Set the OPENAI_API_KEY environment variable", file=sys.stderr)
            print("     Example: export OPENAI_API_KEY='your-api-key-here'", file=sys.stderr)
            return 1

    # Create extractor instance
    if use_azure:
        extractor = DocxExtractor(
            api_key=api_key,
            model=model,
            use_azure=True,
            azure_endpoint=azure_endpoint,
            azure_api_version=azure_api_version
        )
    else:
        extractor = DocxExtractor(
            api_key=api_key,
            model=model,
            api_base=api_base,
            use_azure=False
        )

    try:
        # Determine whether it's a file or folder
        if input_path.is_file():
            # Process single file
            if input_path.suffix.lower() not in ['.docx', '.doc']:
                print(f"Warning: File may not be in DOCX format - {args.input_path}", file=sys.stderr)

            # Choose processing method based on batch mode flag
            if args.batch_mode:
                extraction = extractor.process_file_with_batches(
                    str(input_path),
                    batch_size=args.batch_size,
                    max_concurrent=args.max_concurrent,
                    max_tokens_per_batch=args.max_tokens,
                    output_path=args.output
                )
            else:
                extraction = extractor.process_file(str(input_path), args.output)

            # Output JSON format if --json flag is specified
            if args.json:
                print("\n" + "=" * 80)
                print("JSON output:")
                print(json.dumps(extraction.model_dump(), ensure_ascii=False, indent=2))

        elif input_path.is_dir():
            # Process all DOCX files in the folder
            docx_files = list(input_path.glob("*.docx")) + list(input_path.glob("*.doc"))

            if not docx_files:
                print(f"Error: No DOCX files found in folder - {args.input_path}", file=sys.stderr)
                return 1

            print(f"Found {len(docx_files)} DOCX file(s)")
            if args.batch_mode:
                print(f"Batch mode enabled: {args.batch_size} rows per batch, {args.max_concurrent} max concurrent")
            print("=" * 80)

            extractions = []
            for idx, docx_file in enumerate(docx_files, 1):
                print(f"\n[{idx}/{len(docx_files)}] Processing file: {docx_file.name}")
                print("-" * 80)

                try:
                    if args.batch_mode:
                        extraction = extractor.process_file_with_batches(
                            str(docx_file),
                            batch_size=args.batch_size,
                            max_concurrent=args.max_concurrent,
                            max_tokens_per_batch=args.max_tokens
                        )
                    else:
                        extraction = extractor.process_file(str(docx_file))
                    extractions.append((docx_file.name, extraction))
                except Exception as e:
                    print(f"Warning: Error processing file {docx_file.name}: {e}", file=sys.stderr)
                    continue

            # Save results
            if args.output:
                output_path = Path(args.output)
                if output_path.suffix.lower() == '.csv':
                    # Export to CSV
                    DocxExtractor.export_to_csv(extractions, str(output_path))
                    print(f"\nAll results saved to CSV file: {output_path}")
                elif output_path.suffix.lower() == '.json':
                    # Export to JSON
                    all_data = {
                        "files": [
                            {
                                "filename": filename,
                                "records": extraction.model_dump()["records"]
                            }
                            for filename, extraction in extractions
                        ]
                    }
                    with open(output_path, 'w', encoding='utf-8') as f:
                        json.dump(all_data, f, ensure_ascii=False, indent=2)
                    print(f"\nAll results saved to JSON file: {output_path}")
                else:
                    print(f"Warning: Unsupported output format {output_path.suffix}, please use .csv or .json", file=sys.stderr)

            # Output JSON format to standard output if --json flag is specified
            if args.json:
                all_data = {
                    "files": [
                        {
                            "filename": filename,
                            "records": extraction.model_dump()["records"]
                        }
                        for filename, extraction in extractions
                    ]
                }
                print("\n" + "=" * 80)
                print("JSON output:")
                print(json.dumps(all_data, ensure_ascii=False, indent=2))

        else:
            print(f"Error: Input path is neither a file nor a folder - {args.input_path}", file=sys.stderr)
            return 1

        return 0

    except FileNotFoundError as e:
        print(f"Error: File not found - {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Error: An error occurred during processing - {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())