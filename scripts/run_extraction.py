"""
Simple usage example for PDF/Word table extraction.

This script demonstrates basic usage of the DocxTableExtractor.
"""

from pathlib import Path
import sys

# Add parent directory to path so we can import from src
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.docx_extractor import DocxTableExtractor


def main():
    """Run basic extraction example."""
    
    # Initialize extractor
    print("Initializing extractor...")
    extractor = DocxTableExtractor(
        clean_data=True,          # Clean numeric formatting
        auto_convert_pdf=True     # Auto-convert PDFs to Word
    )
    
    # Define paths
    sample_dir = Path("examples/sample_data")
    output_dir = Path("examples/output")
    
    # Get all PDF and Word files
    pdf_files = list(sample_dir.glob("*.pdf"))
    docx_files = list(sample_dir.glob("*.docx"))
    all_files = pdf_files + docx_files
    
    if not all_files:
        print(f"\n⚠️ No PDF or Word files found in {sample_dir}")
        print("   Please add sample files to test the extractor.")
        return
    
    print(f"\n✅ Found {len(all_files)} file(s) to process:")
    for f in all_files:
        print(f"   - {f.name}")
    
    # Process each file
    print("\n" + "="*50)
    for input_file in all_files:
        print(f"\nProcessing: {input_file.name}")
        print("-" * 50)
        
        output_file = output_dir / f"{input_file.stem}_output.xlsx"
        
        try:
            extractor.process_file(input_file, output_file,    keep_intermediate=True)
            print(f"✅ Success! Output saved to: {output_file}")
        except Exception as e:
            print(f"❌ Error: {e}")
    
    print("\n" + "="*50)
    print(f"\n✅ All files processed!")
    print(f"📁 Check output folder: {output_dir}")
    print("\nOutput files:")
    for excel_file in output_dir.glob("*.xlsx"):
        print(f"   - {excel_file.name}")


if __name__ == "__main__":
    main()