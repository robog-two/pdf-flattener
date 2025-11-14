import platform
import sys
import os
import shutil
import tempfile
import fitz
from pdf2image import convert_from_path
from datetime import datetime
import subprocess
from contextlib import contextmanager
import logging
import time

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

DPI = 200

@contextmanager
def safe_temp_file(suffix):
    """Context manager for safely handling temporary files."""
    temp_file = None
    try:
        # On Windows, use a more reliable temporary file creation
        if platform.system() == "Windows":
            temp_dir = os.environ.get('TEMP') or os.environ.get('TMP') or tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, f"pdf_flattener_{os.urandom(8).hex()}{suffix}")
            # Create an empty file
            with open(temp_path, 'wb') as f:
                pass
            temp_file = type('TempFile', (), {'name': temp_path, 'close': lambda: None})
        else:
            temp_file = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
        
        yield temp_file.name
    finally:
        if temp_file and os.path.exists(temp_file.name):
            try:
                # On Windows, we need to ensure the file is not in use
                if platform.system() == "Windows":
                    # Close any open handles
                    if hasattr(temp_file, 'close'):
                        temp_file.close()
                    # Wait a bit to ensure file is released
                    time.sleep(0.1)
                    # Try to force close any handles
                    try:
                        import ctypes
                        from ctypes import wintypes
                        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
                        CloseHandle = kernel32.CloseHandle
                        CloseHandle.argtypes = [wintypes.HANDLE]
                        CloseHandle.restype = wintypes.BOOL
                        # Try to close any open handles to the file
                        handle = kernel32.CreateFileW(
                            temp_file.name,
                            0x80000000,  # GENERIC_READ
                            0,  # No sharing
                            None,
                            3,  # OPEN_EXISTING
                            0x02000000,  # FILE_FLAG_DELETE_ON_CLOSE
                            None
                        )
                        if handle != -1:
                            CloseHandle(handle)
                    except Exception as e:
                        logger.warning(f"Failed to force close handles: {e}")
                os.remove(temp_file.name)
            except OSError as e:
                logger.warning(f"Failed to remove temporary file {temp_file.name}: {e}")
                # On Windows, if we can't delete the file, try to mark it for deletion on reboot
                if platform.system() == "Windows":
                    try:
                        import ctypes
                        from ctypes import wintypes
                        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
                        DeleteFile = kernel32.DeleteFileW
                        DeleteFile.argtypes = [wintypes.LPCWSTR]
                        DeleteFile.restype = wintypes.BOOL
                        DeleteFile(temp_file.name)
                    except Exception as e:
                        logger.warning(f"Failed to mark file for deletion: {e}")

def get_poppler_path():
    if platform.system() == "Windows":
        # First check if user has specified a custom path via environment variable
        custom_path = os.environ.get("POPPLER_PATH")
        logger.debug(f"Checking POPPLER_PATH environment variable: {custom_path}")
        if custom_path and os.path.exists(custom_path):
            logger.debug(f"Found Poppler at custom path: {custom_path}")
            return custom_path

        # Common installation paths to check
        possible_paths = [
            r"C:\poppler\Library\bin",
            r"C:\Program Files\poppler-0.68.0\bin",
            r"C:\Program Files\poppler-23.11.0\bin",
            r"C:\Program Files\poppler\bin",
            r"C:\poppler-0.68.0\bin",
            r"C:\poppler\bin",
        ]

        # Check each possible path
        for path in possible_paths:
            logger.debug(f"Checking path: {path}")
            if os.path.exists(path):
                logger.debug(f"Found Poppler at: {path}")
                return path

        # If no path found, provide detailed error message
        logger.error(
            "Error: Poppler not found. Please install Poppler using one of these methods:\n"
            "1. Download from: https://github.com/oschwartz10612/poppler-windows/releases/\n"
            "2. Extract to C:\\poppler\n"
            "3. Add the bin directory to your PATH\n"
            "4. Or set POPPLER_PATH environment variable to your Poppler bin directory"
        )
        return None
    elif platform.system() in ["Linux", "Darwin"]:
        logger.debug("Checking for pdftoppm in PATH")
        if not shutil.which("pdftoppm"):
            logger.error(
                "Error: Poppler utilities not found. Please install Poppler (Linux: sudo apt-get install poppler-utils, macOS: brew install poppler)."
            )
            return None
        logger.debug("Found pdftoppm in PATH")
        return None
    else:
        logger.error(
            "Error: Unsupported operating system. This script supports Windows, macOS, and Linux."
        )
        return None

def extract_images_from_pdf(pdf_path, dpi=DPI):
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"Input PDF file not found: {pdf_path}")
    
    try:
        poppler_path = get_poppler_path()
        images = convert_from_path(pdf_path, dpi=dpi, poppler_path=poppler_path)
        return images
    except Exception as e:
        logger.error(f"Failed to extract images from PDF: {e}")
        raise

def create_pdf_from_images(images, output_path):
    try:
        doc = fitz.open()
        
        for image in images:
            with safe_temp_file(suffix=".jpg") as temp_img_file_path:
                image.save(temp_img_file_path, format="JPEG", quality=50)
                img_pix = fitz.Pixmap(temp_img_file_path)
                page = doc.new_page(width=img_pix.width, height=img_pix.height)
                page.insert_image(page.rect, pixmap=img_pix)
                img_pix = None

        doc.save(output_path)
        doc.close()
    except Exception as e:
        logger.error(f"Failed to create PDF from images: {e}")
        raise

def compress_pdf(input_pdf, output_pdf):
    try:
        doc = fitz.open(input_pdf)
        doc.save(output_pdf, garbage=4, deflate=True, incremental=False, clean=True)
        doc.close()
    except Exception as e:
        logger.error(f"Failed to compress PDF: {e}")
        raise

def set_metadata(pdf_path, creation_date=None, modification_date=None):
    try:
        doc = fitz.open(pdf_path)

        def format_date(dt):
            return dt.strftime("D:%Y%m%d%H%M%S+00'00'")

        # Prepare the metadata dictionary
        new_metadata = {}

        if creation_date:
            try:
                creation_dt = datetime.strptime(creation_date, "%Y-%m-%d")
                creation_str = format_date(creation_dt)
                new_metadata["creationDate"] = creation_str
            except ValueError:
                logger.error("Invalid creation date format. Use YYYY-MM-DD.")
                raise ValueError("Invalid creation date format. Use YYYY-MM-DD.")

        if modification_date:
            try:
                modification_dt = datetime.strptime(modification_date, "%Y-%m-%d")
                modification_str = format_date(modification_dt)
                new_metadata["modDate"] = modification_str
            except ValueError:
                logger.error("Invalid modification date format. Use YYYY-MM-DD.")
                raise ValueError("Invalid modification date format. Use YYYY-MM-DD.")

        # Update metadata after compression
        doc.set_metadata(new_metadata)
        doc.save(pdf_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
        doc.close()
    except Exception as e:
        logger.error(f"Failed to set metadata: {e}")
        raise

def flatten_pdf(
    pdf_path, output_path, creation_date=None, modification_date=None, dpi=DPI
):
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"Input PDF file not found: {pdf_path}")
    
    try:
        with safe_temp_file(suffix=".pdf") as temp_output_path:
            images = extract_images_from_pdf(pdf_path, dpi)
            create_pdf_from_images(images, temp_output_path)

            with safe_temp_file(suffix=".pdf") as compressed_output_path:
                compress_pdf(temp_output_path, compressed_output_path)

                # Retrieve the original file's metadata dates
                doc = fitz.open(pdf_path)
                original_metadata = doc.metadata
                original_creation_date = original_metadata.get("creationDate", "")
                original_modification_date = original_metadata.get("modDate", "")
                doc.close()

                # Extract and format the original creation and modification datetime if they exist
                if original_creation_date:
                    original_creation_dt = datetime.strptime(
                        original_creation_date[2:16], "%Y%m%d%H%M%S"
                    )
                else:
                    original_creation_dt = datetime.fromtimestamp(os.path.getctime(pdf_path))

                if original_modification_date:
                    original_modification_dt = datetime.strptime(
                        original_modification_date[2:16], "%Y%m%d%H%M%S"
                    )
                else:
                    original_modification_dt = datetime.fromtimestamp(os.path.getmtime(pdf_path))

                # Use original hours and minutes if only the date part is provided
                if creation_date:
                    creation_date = datetime.strptime(creation_date, "%Y-%m-%d")
                    creation_dt = creation_date.replace(
                        hour=original_creation_dt.hour,
                        minute=original_creation_dt.minute,
                        second=original_creation_dt.second,
                    )
                else:
                    creation_dt = original_creation_dt

                if modification_date:
                    modification_date = datetime.strptime(modification_date, "%Y-%m-%d")
                    modification_dt = modification_date.replace(
                        hour=original_modification_dt.hour,
                        minute=original_modification_dt.minute,
                        second=original_modification_dt.second,
                    )
                else:
                    modification_dt = original_modification_dt

                # Ensure modification date is not earlier than the creation date
                if modification_dt < creation_dt:
                    modification_dt = creation_dt

                # Set metadata after compression
                set_metadata(
                    compressed_output_path,
                    creation_dt.strftime("%Y-%m-%d"),
                    modification_dt.strftime("%Y-%m-%d"),
                )

                # Move final compressed and metadata-set file to the intended output location
                shutil.move(compressed_output_path, output_path)

                # Set file system times if creation/modification dates are provided
                set_file_times(output_path, creation_dt, modification_dt)

    except Exception as e:
        logger.error(f"Failed to flatten PDF: {e}")
        raise

def set_file_times(file_path, creation_dt, modification_dt):
    # Apply modification and access times
    mod_timestamp = modification_dt.timestamp()
    os.utime(file_path, (mod_timestamp, mod_timestamp))

    if platform.system() == "Darwin":  # macOS only
        try:
            # Set the creation date first
            creation_str = creation_dt.strftime("%Y%m%d%H%M")
            subprocess.run(["touch", "-t", creation_str, file_path], check=True)

            # Set the modification date separately using os.utime() to prevent touch from affecting it
            os.utime(file_path, (mod_timestamp, mod_timestamp))
        except subprocess.CalledProcessError as e:
            print(f"Error setting file creation date: {e}")


def parse_arguments():
    import argparse

    parser = argparse.ArgumentParser(
        description="Flatten a PDF and optionally set metadata."
    )

    parser.add_argument("input_pdf", help="Path to the input PDF file")
    parser.add_argument("--output", "-o", help="Output PDF file name", default=None)
    parser.add_argument(
        "--dpi",
        "-d",
        type=int,
        help="DPI for image extraction (default: 200)",
        default=200,
    )
    parser.add_argument(
        "--creation-date", "-c", help="Creation date in YYYY-MM-DD format", default=None
    )
    parser.add_argument(
        "--modification-date",
        "-m",
        help="Modification date in YYYY-MM-DD format",
        default=None,
    )

    return parser.parse_args()


def main():
    args = parse_arguments()

    input_pdf = args.input_pdf
    output_pdf = args.output if args.output else f"flat-{os.path.basename(input_pdf)}"
    creation_date = args.creation_date
    modification_date = args.modification_date
    dpi = args.dpi

    flatten_pdf(input_pdf, output_pdf, creation_date, modification_date, dpi)

    print(f"File {output_pdf} saved successfully.")


if __name__ == "__main__":
    main()
