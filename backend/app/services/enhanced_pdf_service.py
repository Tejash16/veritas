# app/services/enhanced_pdf_service.py
import fitz  # PyMuPDF
from typing import List, Dict, Any, Tuple
import structlog
import io
import base64
from PIL import Image

logger = structlog.get_logger()

class EnhancedPDFService:
    def __init__(self):
        pass

    async def extract_with_visual_coordinates(self, file_content: bytes) -> Tuple[str, List[bytes], Dict[str, Any]]:
        """
        Extract PDF with precise visual coordinates for UI overlay
        """
        logger.info("Starting enhanced PDF extraction with coordinates")
        
        try:
            doc = fitz.open(stream=file_content, filetype="pdf")
            
            # Extract text
            full_text = ""
            for page_num in range(len(doc)):
                page = doc[page_num]
                full_text += f"\n--- Page {page_num + 1} ---\n{page.get_text()}\n"
            
            # Extract high-quality images for AI analysis
            page_images = []
            page_metadata = []
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                
                # Convert page to high-quality image
                mat = fitz.Matrix(2.0, 2.0)  # 2x zoom for better quality
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("png")
                page_images.append(img_data)
                
                # Extract page metadata for coordinate mapping
                page_metadata.append({
                    "page_number": page_num + 1,
                    "width": pix.width,
                    "height": pix.height,
                    "original_width": page.rect.width,
                    "original_height": page.rect.height,
                    "zoom_factor": 2.0
                })
                
                pix = None  # Free memory
            
            doc.close()
            
            processing_metadata = {
                "total_pages": len(page_images),
                "page_metadata": page_metadata,
                "extraction_timestamp": time.time()
            }
            
            logger.info("PDF extraction completed", 
                       pages=len(page_images),
                       text_length=len(full_text))
            
            return full_text, page_images, processing_metadata
            
        except Exception as e:
            logger.error("Enhanced PDF extraction failed", error=str(e))
            raise

# Initialize the enhanced service
enhanced_pdf_service = EnhancedPDFService()