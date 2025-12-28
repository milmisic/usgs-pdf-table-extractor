"""
PDF to Word conversion utilities.
"""

from pathlib import Path
from typing import List, Optional
import logging

try:
    from pdf2docx import Converter
    PDF2DOCX_AVAILABLE = True
except ImportError:
    PDF2DOCX_AVAILABLE = False

logger = logging.getLogger(__name__)


class PDFConverter:
    """
    Convert PDF files to Word format for table extraction.
    Uses pdf2docx library for conversion.
    
    """
    
    def __init__(self):
        """Initialize PDF converter."""
        if not PDF2DOCX_AVAILABLE:
            logger.warning(
                "pdf2docx library not installed. "
                "Install with: pip install pdf2docx"
            )
    
    def convert(
        self, 
        pdf_path: Path, 
        docx_path: Optional[Path] = None
    ) -> Path:
        """
        Convert PDF to Word document.
        
        Args:
            pdf_path: Path to input PDF file
            docx_path: Path for output Word file (optional)
                      If not provided, uses same name with .docx extension
        
        Returns:
            Path to created Word document
            
        Raises:
            ImportError: If pdf2docx is not installed
            FileNotFoundError: If PDF file doesn't exist
            
        Example:
            >>> converter = PDFConverter()
            >>> docx_file = converter.convert(Path("report.pdf"))
            >>> print(docx_file)
            PosixPath('report.docx')
        """
        if not PDF2DOCX_AVAILABLE:
            raise ImportError(
                "pdf2docx library not installed. "
                "Install with: pip install pdf2docx"
            )
        
        pdf_path = Path(pdf_path)
        
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        
        # Default output path
        if docx_path is None:
            docx_path = pdf_path.with_suffix('.docx')
        else:
            docx_path = Path(docx_path)
        
        logger.info(f"Converting {pdf_path.name} → {docx_path.name}")
        
        try:
            cv = Converter(str(pdf_path))
            cv.convert(str(docx_path))
            cv.close()
            
            logger.info(f"✅ Conversion complete: {docx_path.name}")
            return docx_path
            
        except Exception as e:
            logger.error(f"❌ Conversion failed for {pdf_path.name}: {e}")
            raise
    
    def batch_convert(
        self, 
        input_dir: Path, 
        output_dir: Optional[Path] = None,
        pattern: str = "*.pdf"
    ) -> List[Path]:
        """
        Convert multiple PDFs to Word format.
        
        Args:
            input_dir: Directory containing PDF files
            output_dir: Directory for Word files (default: same as input_dir)
            pattern: File pattern to match (default: *.pdf)
        
        Returns:
            List of paths to successfully created Word documents
            
        Example:
            >>> converter = PDFConverter()
            >>> docx_files = converter.batch_convert(
            ...     input_dir=Path("./pdfs"),
            ...     output_dir=Path("./docx")
            ... )
            >>> print(f"Converted {len(docx_files)} files")
        """
        input_dir = Path(input_dir)
        output_dir = Path(output_dir) if output_dir else input_dir
        output_dir.mkdir(parents=True, exist_ok=True)
        
        pdf_files = sorted(input_dir.glob(pattern))
        logger.info(f"Found {len(pdf_files)} PDF files")
        
        converted_files = []
        failed_files = []
        
        for pdf_path in pdf_files:
            try:
                docx_path = output_dir / pdf_path.with_suffix('.docx').name
                converted_path = self.convert(pdf_path, docx_path)
                converted_files.append(converted_path)
            except Exception as e:
                logger.warning(f"⚠️ Skipping {pdf_path.name}: {e}")
                failed_files.append(pdf_path)
                continue
        
        logger.info(
            f"✅ Converted {len(converted_files)}/{len(pdf_files)} files"
        )
        
        if failed_files:
            logger.warning(
                f"Failed conversions: {[f.name for f in failed_files]}"
            )
        
        return converted_files