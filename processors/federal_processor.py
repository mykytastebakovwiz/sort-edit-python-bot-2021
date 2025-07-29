from PyPDF2 import PdfReader, PdfWriter
from pathlib import Path

def process_federal_files(company_path: Path, federal_path: Path):
    federal_path.mkdir(exist_ok=True)
    print("Starting federal processing")

    for subfolder in sorted(company_path.iterdir()):
        if subfolder.is_dir():
            for file in sorted(subfolder.glob('FTFCS*.pdf')):
                # Read and strip first 2 pages
                reader = PdfReader(str(file))
                writer = PdfWriter()
                for page in reader.pages[2:]:  # Skip first 2 pages
                    writer.add_page(page)

                # Write to federal directory
                output_path = federal_path / file.name
                with open(output_path, 'wb') as f:
                    writer.write(f)

                # Remove original file from company folder
                file.unlink()

    print("Federal processing complete")
    
