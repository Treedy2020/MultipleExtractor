"""DOCX Document Extractor - GUI Application

A graphical user interface for extracting structured information from DOCX documents
using Azure OpenAI or OpenAI API with batch processing support.
"""

import asyncio
import json
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Optional

from main import DocxExtractor


class ExtractorGUI:
    """GUI application for DOCX document extraction."""

    def __init__(self, root: tk.Tk):
        """Initialize the GUI application.

        Args:
            root: The root tkinter window.
        """
        self.root = root
        self.root.title("DOCX Document Extractor")
        self.root.geometry("800x700")

        # Variables
        self.api_key_var = tk.StringVar()
        self.endpoint_var = tk.StringVar(value="https://aidt-eastus2-ai.openai.azure.com")
        self.api_version_var = tk.StringVar(value="2025-03-01-preview")
        self.deployment_var = tk.StringVar(value="gpt-5-chat")
        self.batch_mode_var = tk.BooleanVar(value=True)
        self.batch_size_var = tk.IntVar(value=10)
        self.max_concurrent_var = tk.IntVar(value=5)
        self.max_tokens_var = tk.IntVar(value=20000)
        self.input_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.output_format_var = tk.StringVar(value="csv")

        self.processing = False
        self.extractor: Optional[DocxExtractor] = None

        self._create_widgets()

    def _create_widgets(self):
        """Create and layout all GUI widgets."""
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        row = 0

        # Title
        title_label = ttk.Label(
            main_frame,
            text="DOCX Document Extractor",
            font=("Helvetica", 16, "bold")
        )
        title_label.grid(row=row, column=0, columnspan=3, pady=(0, 10))
        row += 1

        # Azure OpenAI Configuration Section
        config_frame = ttk.LabelFrame(main_frame, text="Azure OpenAI Configuration", padding="10")
        config_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)
        row += 1

        config_row = 0

        # API Key
        ttk.Label(config_frame, text="API Key:").grid(row=config_row, column=0, sticky=tk.W, pady=2)
        api_key_entry = ttk.Entry(config_frame, textvariable=self.api_key_var, show="*", width=50)
        api_key_entry.grid(row=config_row, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2)
        config_row += 1

        # Endpoint
        ttk.Label(config_frame, text="Endpoint:").grid(row=config_row, column=0, sticky=tk.W, pady=2)
        ttk.Entry(config_frame, textvariable=self.endpoint_var, width=50).grid(
            row=config_row, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2
        )
        config_row += 1

        # API Version
        ttk.Label(config_frame, text="API Version:").grid(row=config_row, column=0, sticky=tk.W, pady=2)
        ttk.Entry(config_frame, textvariable=self.api_version_var, width=50).grid(
            row=config_row, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2
        )
        config_row += 1

        # Deployment
        ttk.Label(config_frame, text="Deployment:").grid(row=config_row, column=0, sticky=tk.W, pady=2)
        ttk.Entry(config_frame, textvariable=self.deployment_var, width=50).grid(
            row=config_row, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=2
        )

        # File Selection Section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        row += 1

        file_row = 0

        # Input file
        ttk.Label(file_frame, text="Input File:").grid(row=file_row, column=0, sticky=tk.W, pady=2)
        ttk.Entry(file_frame, textvariable=self.input_file_var, width=40).grid(
            row=file_row, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2
        )
        ttk.Button(file_frame, text="Browse...", command=self._browse_input).grid(
            row=file_row, column=2, pady=2
        )
        file_row += 1

        # Output file
        ttk.Label(file_frame, text="Output File:").grid(row=file_row, column=0, sticky=tk.W, pady=2)
        ttk.Entry(file_frame, textvariable=self.output_file_var, width=40).grid(
            row=file_row, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2
        )
        ttk.Button(file_frame, text="Browse...", command=self._browse_output).grid(
            row=file_row, column=2, pady=2
        )
        file_row += 1

        # Output format
        ttk.Label(file_frame, text="Output Format:").grid(row=file_row, column=0, sticky=tk.W, pady=2)
        format_frame = ttk.Frame(file_frame)
        format_frame.grid(row=file_row, column=1, sticky=tk.W, padx=(5, 0), pady=2)
        ttk.Radiobutton(format_frame, text="CSV", variable=self.output_format_var, value="csv").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(format_frame, text="JSON", variable=self.output_format_var, value="json").pack(side=tk.LEFT)

        # Processing Options Section
        options_frame = ttk.LabelFrame(main_frame, text="Processing Options", padding="10")
        options_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        options_frame.columnconfigure(1, weight=1)
        row += 1

        options_row = 0

        # Batch mode
        ttk.Checkbutton(
            options_frame,
            text="Enable Batch Processing Mode",
            variable=self.batch_mode_var,
            command=self._toggle_batch_options
        ).grid(row=options_row, column=0, columnspan=2, sticky=tk.W, pady=2)
        options_row += 1

        # Batch size
        ttk.Label(options_frame, text="Batch Size (rows):").grid(row=options_row, column=0, sticky=tk.W, pady=2)
        self.batch_size_spinbox = ttk.Spinbox(
            options_frame,
            from_=1,
            to=100,
            textvariable=self.batch_size_var,
            width=10
        )
        self.batch_size_spinbox.grid(row=options_row, column=1, sticky=tk.W, padx=(5, 0), pady=2)
        options_row += 1

        # Max concurrent
        ttk.Label(options_frame, text="Max Concurrent API Calls:").grid(
            row=options_row, column=0, sticky=tk.W, pady=2
        )
        self.max_concurrent_spinbox = ttk.Spinbox(
            options_frame,
            from_=1,
            to=20,
            textvariable=self.max_concurrent_var,
            width=10
        )
        self.max_concurrent_spinbox.grid(row=options_row, column=1, sticky=tk.W, padx=(5, 0), pady=2)
        options_row += 1

        # Max tokens
        ttk.Label(options_frame, text="Max Tokens per Batch:").grid(
            row=options_row, column=0, sticky=tk.W, pady=2
        )
        self.max_tokens_spinbox = ttk.Spinbox(
            options_frame,
            from_=5000,
            to=100000,
            increment=1000,
            textvariable=self.max_tokens_var,
            width=10
        )
        self.max_tokens_spinbox.grid(row=options_row, column=1, sticky=tk.W, padx=(5, 0), pady=2)

        # Progress Section
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(0, weight=1)
        row += 1

        # Configure main frame row weight for progress section
        main_frame.rowconfigure(row - 1, weight=1)

        # Log text area
        self.log_text = scrolledtext.ScrolledText(
            progress_frame,
            wrap=tk.WORD,
            width=70,
            height=12,
            font=("Courier", 9)
        )
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Progress bar
        self.progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))

        # Action Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, columnspan=3, pady=(0, 0))

        self.process_button = ttk.Button(
            button_frame,
            text="Start Processing",
            command=self._start_processing,
            width=20
        )
        self.process_button.pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="Clear Log",
            command=self._clear_log,
            width=15
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="Exit",
            command=self.root.quit,
            width=15
        ).pack(side=tk.LEFT, padx=5)

        # Initialize batch options state
        self._toggle_batch_options()

    def _toggle_batch_options(self):
        """Enable or disable batch processing options based on checkbox."""
        state = 'normal' if self.batch_mode_var.get() else 'disabled'
        self.batch_size_spinbox.config(state=state)
        self.max_concurrent_spinbox.config(state=state)
        self.max_tokens_spinbox.config(state=state)

    def _browse_input(self):
        """Open file dialog to select input DOCX file."""
        filename = filedialog.askopenfilename(
            title="Select DOCX file",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if filename:
            self.input_file_var.set(filename)
            # Auto-suggest output filename
            if not self.output_file_var.get():
                input_path = Path(filename)
                output_path = input_path.parent / f"{input_path.stem}_output.{self.output_format_var.get()}"
                self.output_file_var.set(str(output_path))

    def _browse_output(self):
        """Open file dialog to select output file location."""
        ext = self.output_format_var.get()
        filename = filedialog.asksaveasfilename(
            title="Save output as",
            defaultextension=f".{ext}",
            filetypes=[
                (f"{ext.upper()} Files", f"*.{ext}"),
                ("All Files", "*.*")
            ]
        )
        if filename:
            self.output_file_var.set(filename)

    def _log(self, message: str):
        """Add a message to the log text area.

        Args:
            message: The message to log.
        """
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def _clear_log(self):
        """Clear the log text area."""
        self.log_text.delete(1.0, tk.END)

    def _validate_inputs(self) -> bool:
        """Validate all input fields.

        Returns:
            True if all inputs are valid, False otherwise.
        """
        if not self.api_key_var.get():
            messagebox.showerror("Error", "Please enter your Azure OpenAI API Key")
            return False

        if not self.endpoint_var.get():
            messagebox.showerror("Error", "Please enter the Azure OpenAI Endpoint")
            return False

        if not self.input_file_var.get():
            messagebox.showerror("Error", "Please select an input DOCX file")
            return False

        if not Path(self.input_file_var.get()).exists():
            messagebox.showerror("Error", "Input file does not exist")
            return False

        if not self.output_file_var.get():
            messagebox.showerror("Error", "Please specify an output file")
            return False

        return True

    def _start_processing(self):
        """Start the document processing in a separate thread."""
        if self.processing:
            messagebox.showwarning("Warning", "Processing is already in progress")
            return

        if not self._validate_inputs():
            return

        # Disable process button
        self.process_button.config(state='disabled', text='Processing...')
        self.processing = True
        self.progress_bar.start()

        # Clear previous log
        self._clear_log()
        self._log("Starting extraction process...")

        # Start processing in a separate thread
        thread = threading.Thread(target=self._process_document, daemon=True)
        thread.start()

    def _process_document(self):
        """Process the document (runs in a separate thread)."""
        try:
            # Get configuration values
            api_key = self.api_key_var.get()
            model = self.deployment_var.get()
            endpoint = self.endpoint_var.get()
            api_version = self.api_version_var.get()
            input_file = self.input_file_var.get()
            output_file = self.output_file_var.get()
            batch_mode = self.batch_mode_var.get()
            batch_size = self.batch_size_var.get()
            max_concurrent = self.max_concurrent_var.get()
            max_tokens = self.max_tokens_var.get()

            # Create extractor in this thread
            self._log("Initializing Azure OpenAI client...")
            self.extractor = DocxExtractor(
                api_key=api_key,
                model=model,
                use_azure=True,
                azure_endpoint=endpoint,
                azure_api_version=api_version
            )

            # Process based on mode
            if batch_mode:
                self._log(f"Using batch mode: {batch_size} rows per batch")
                self._log(f"Max concurrent API calls: {max_concurrent}")
                self._log(f"Max tokens per batch: {max_tokens}")

                # Run async processing with proper event loop handling
                extraction = self._run_async_processing(input_file, batch_size, max_tokens, max_concurrent)
            else:
                self._log("Using standard processing mode")
                extraction = self.extractor.process_file(input_file)

            # Save results
            self._log(f"\nSaving results to: {output_file}")
            output_path = Path(output_file)

            if output_path.suffix.lower() == '.csv':
                filename = Path(input_file).name
                DocxExtractor.export_to_csv([(filename, extraction)], str(output_path))
                self._log("Results saved to CSV format")
            else:
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(extraction.model_dump(), f, ensure_ascii=False, indent=2)
                self._log("Results saved to JSON format")

            self._log(f"\n✓ Processing completed successfully!")
            self._log(f"Total records extracted: {len(extraction.records)}")

            # Show success message
            record_count = len(extraction.records)
            self.root.after(0, lambda count=record_count, file=output_file: messagebox.showinfo(
                "Success",
                f"Processing completed!\n\nTotal records: {count}\nOutput: {file}"
            ))

        except Exception as e:
            error_msg = str(e)
            self._log(f"\n✗ Error: {error_msg}")
            self.root.after(0, lambda msg=error_msg: messagebox.showerror("Error", f"Processing failed:\n\n{msg}"))

        finally:
            # Re-enable process button
            self.root.after(0, self._finish_processing)

    def _run_async_processing(self, input_file: str, batch_size: int, max_tokens: int, max_concurrent: int):
        """Run async batch processing in a thread-safe way.

        Args:
            input_file: Path to input DOCX file.
            batch_size: Batch size.
            max_tokens: Max tokens per batch.
            max_concurrent: Max concurrent requests.

        Returns:
            DocumentExtraction with all extracted records.
        """
        # Create a new event loop for this thread
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

        try:
            # Run the async processing
            extraction = loop.run_until_complete(
                self._process_with_batches_async(input_file, batch_size, max_tokens, max_concurrent)
            )
            return extraction
        finally:
            # Clean up
            loop.close()

    async def _process_with_batches_async(
        self,
        input_file: str,
        batch_size: int,
        max_tokens: int,
        max_concurrent: int
    ):
        """Process document with batch mode (async).

        Args:
            input_file: Path to input DOCX file.
            batch_size: Batch size.
            max_tokens: Max tokens per batch.
            max_concurrent: Max concurrent requests.

        Returns:
            DocumentExtraction with all extracted records.
        """
        # Import here to avoid circular imports
        from openai import AsyncAzureOpenAI, AsyncOpenAI

        # Recreate async client in the current event loop to avoid connection issues
        if self.extractor.use_azure:
            self.extractor.async_client = AsyncAzureOpenAI(
                api_key=self.api_key_var.get(),
                azure_endpoint=self.endpoint_var.get(),
                api_version=self.api_version_var.get()
            )
        else:
            client_kwargs = {"api_key": self.api_key_var.get()}
            if hasattr(self.extractor.client, 'base_url'):
                client_kwargs["base_url"] = str(self.extractor.client.base_url)
            self.extractor.async_client = AsyncOpenAI(**client_kwargs)

        self._log(f"Reading file: {input_file}")
        batches = self.extractor.read_docx_in_batches(
            input_file,
            batch_size,
            max_tokens
        )

        self._log(f"Document split into {len(batches)} batch(es)\n")

        # Process batches
        extraction = await self.extractor.process_batches_async(
            batches,
            max_concurrent
        )

        return extraction

    def _finish_processing(self):
        """Clean up after processing completes."""
        self.progress_bar.stop()
        self.process_button.config(state='normal', text='Start Processing')
        self.processing = False


def main():
    """Main entry point for the GUI application."""
    root = tk.Tk()
    app = ExtractorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
