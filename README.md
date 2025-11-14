# PDF Flattener

A command-line tool that converts PDF files into image-based documents, simulating scanned documents while compressing them.

## Installation

### Prerequisites

- **Poppler**: Required for PDF processing
  - macOS: `brew install poppler`
  - Linux: `sudo apt-get install poppler-utils`
  - Windows: Download and install from [poppler releases](https://github.com/oschwartz10612/poppler-windows/releases/)

### Install with pipx

```bash
pipx install git+https://github.com/robog-two/pdf-flattener.git
```

## Usage

```bash
flatten-pdf input.pdf
```

### Options

- `--output`, `-o`: Output PDF file name (default: "flat-{input_filename}")
- `--dpi`, `-d`: DPI for image extraction (default: 200)
- `--creation-date`, `-c`: Creation date in YYYY-MM-DD format
- `--modification-date`, `-m`: Modification date in YYYY-MM-DD format

### Example

```bash
flatten-pdf input.pdf --output output.pdf --dpi 300
```
