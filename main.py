import os
from google.api_core.client_options import ClientOptions
from google.cloud import documentai
from tkinter import filedialog, Tk
import mimetypes
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Get configuration from environment variables
GOOGLE_CREDENTIALS_PATH = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
PROJECT_ID = os.getenv("PROJECT_ID")
LOCATION = os.getenv("LOCATION", "eu")
PROCESSOR_ID = os.getenv("PROCESSOR_ID")
PROCESSOR_VERSION = os.getenv("PROCESSOR_VERSION", "rc")


def validate_environment():
    """Validate that all required environment variables are set"""
    if not GOOGLE_CREDENTIALS_PATH:
        print("❌ GOOGLE_APPLICATION_CREDENTIALS not set in .env file")
        return False
    if not PROJECT_ID:
        print("❌ PROJECT_ID not set in .env file")
        return False
    if not PROCESSOR_ID:
        print("❌ PROCESSOR_ID not set in .env file")
        return False
    
    if not os.path.exists(GOOGLE_CREDENTIALS_PATH):
        print(f"❌ Credentials file not found: {GOOGLE_CREDENTIALS_PATH}")
        return False
    
    # Set the environment variable for Google Cloud SDK
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = GOOGLE_CREDENTIALS_PATH
    return True


def process_document(file_path: str, mime_type: str) -> str:
    """Process document and extract text"""
    
    # Initialize Document AI client
    client = documentai.DocumentProcessorServiceClient(
        client_options=ClientOptions(
            api_endpoint=f"{LOCATION}-documentai.googleapis.com"
        )
    )

    # Get processor path
    name = client.processor_version_path(
        PROJECT_ID, LOCATION, PROCESSOR_ID, PROCESSOR_VERSION
    )

    # Read file
    with open(file_path, "rb") as image:
        image_content = image.read()

    # Configure process options for layout analysis
    process_options = documentai.ProcessOptions(
        layout_config=documentai.ProcessOptions.LayoutConfig(
            chunking_config=documentai.ProcessOptions.LayoutConfig.ChunkingConfig(
                chunk_size=1000,
                include_ancestor_headings=True,
            )
        )
    )

    # Process document
    request = documentai.ProcessRequest(
        name=name,
        raw_document=documentai.RawDocument(content=image_content, mime_type=mime_type),
        process_options=process_options,
    )

    result = client.process_document(request=request)
    document = result.document

    # Extract text using multiple methods
    # Method 1: Direct text
    if document.text and document.text.strip():
        return document.text

    # Method 2: From chunks
    if hasattr(document, 'chunked_document') and document.chunked_document:
        if document.chunked_document.chunks:
            chunk_text = ""
            for chunk in document.chunked_document.chunks:
                if hasattr(chunk, 'content'):
                    chunk_text += chunk.content + "\n"
            if chunk_text.strip():
                return chunk_text

    # Method 3: From layout blocks
    if hasattr(document, 'document_layout') and document.document_layout:
        if document.document_layout.blocks:
            layout_text = ""
            for block in document.document_layout.blocks:
                if hasattr(block, 'text_block') and block.text_block:
                    if hasattr(block.text_block, 'text'):
                        layout_text += block.text_block.text + "\n"
            if layout_text.strip():
                return layout_text

    return "No text could be extracted from the document."


def get_mime_type(file_path: str) -> str:
    """Get MIME type for file"""
    mime_type, _ = mimetypes.guess_type(file_path)
    if mime_type is None:
        extension = os.path.splitext(file_path)[1].lower()
        mime_map = {
            '.pdf': 'application/pdf',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.png': 'image/png',
            '.tiff': 'image/tiff',
            '.tif': 'image/tiff',
            '.bmp': 'image/bmp',
            '.gif': 'image/gif',
            '.webp': 'image/webp',
            '.txt': 'text/plain',
            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            '.doc': 'application/msword',
            '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            '.ppt': 'application/vnd.ms-powerpoint',
            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            '.xls': 'application/vnd.ms-excel'
        }
        mime_type = mime_map.get(extension, 'application/pdf')
    return mime_type


def main():
    print("Document AI Text Extractor")
    print("=" * 30)
    
    # Validate environment variables
    if not validate_environment():
        print("\nPlease create a .env file with the required variables:")
        print("GOOGLE_APPLICATION_CREDENTIALS=path/to/your/credentials.json")
        print("PROJECT_ID=your-project-id")
        print("PROCESSOR_ID=your-processor-id")
        print("LOCATION=eu")
        print("PROCESSOR_VERSION=rc")
        return
    
    print("✓ Environment variables validated")
    
    # Hide tkinter root window
    root = Tk()
    root.withdraw()
    
    # File selection dialog
    file_path = filedialog.askopenfilename(
        title="Select a document to process",
        filetypes=[
            ("PDF files", "*.pdf"),
            ("Word documents", "*.docx *.doc"),
            ("PowerPoint files", "*.pptx *.ppt"),
            ("Excel files", "*.xlsx *.xls"),
            ("Image files", "*.jpg *.jpeg *.png *.tiff *.tif *.bmp *.gif *.webp"),
            ("Text files", "*.txt"),
            ("All files", "*.*")
        ]
    )
    
    if not file_path:
        print("No file selected.")
        return
    
    try:
        print(f"Processing: {os.path.basename(file_path)}")
        
        # Get MIME type and process document
        mime_type = get_mime_type(file_path)
        print(f"MIME type: {mime_type}")
        
        extracted_text = process_document(file_path, mime_type)
        
        # Output extracted text
        print("\n" + "=" * 30)
        print("EXTRACTED TEXT:")
        print("=" * 30)
        print(extracted_text)
        
    except Exception as e:
        print(f"❌ Error processing document: {e}")


if __name__ == "__main__":
    main()