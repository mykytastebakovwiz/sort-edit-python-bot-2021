import re
import random
import string
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.errors import PdfReadError
from pathlib import Path

def extract_name_and_zip_from_second_page(pdf_path: Path):
    """Extract full name and ZIP code from the second page of an STFCS file."""
    try:
        reader = PdfReader(str(pdf_path))
        if len(reader.pages) < 2:
            print(f"⚠️ Skipping {pdf_path.name}: Fewer than 2 pages")
            return None

        text = reader.pages[1].extract_text() or ""
        lines = text.splitlines()

        first = last = zip_code = None

        # Try to extract name from known patterns
        for i, line in enumerate(lines):
            # Pattern: line after "Instructions to Mail..." has the name
            if 'Instructions to Mail' in line and i + 1 < len(lines):
                name_line = lines[i + 1].strip()
                name_parts = name_line.split()
                if len(name_parts) == 2:
                    first, last = name_parts[0].title(), name_parts[1].title()
                    break

            # Fallback: standalone uppercase name on a line
            name_match = re.match(r'^([A-Z][A-Z\'\-]+)\s+([A-Z][A-Z\'\-]+)$', line.strip())
            if name_match and not (first and last):
                first, last = name_match.group(1).title(), name_match.group(2).title()

        # Extract ZIP code from lines like: CITY, ST 12345
        for line in lines:
            zip_match = re.search(r'\b([A-Z]{2})\s+(\d{5})(?:-\d{4})?\b', line)
            if zip_match:
                zip_code = zip_match.group(2)
                break

        if not (first and last and zip_code):
            print(f"⚠️ Skipping {pdf_path.name}: Name or ZIP not found (got: {first} {last}, ZIP: {zip_code})")
            return None

        print(f"✅ Extracted: {first} {last}, ZIP: {zip_code} from {pdf_path.name}")
        return first, last, zip_code

    except PdfReadError as e:
        print(f"❌ Skipping {pdf_path.name}: PDF read error - {e}")
        return None
    except Exception as e:
        print(f"❌ Skipping {pdf_path.name}: Unexpected error - {e}")
        return None

def generate_random_filename(first: str, last: str) -> str:
    """Generate a random 6-digit filename based on the name."""
    rand_digits = ''.join(random.choices(string.digits, k=6))
    return f"{last}_{first}_{rand_digits}.pdf"

def attach_w2_to_stfcs(company_path: Path, state_path: Path):
    """Legacy name kept. Now just processes STFCS files (no W2)."""
    state_path.mkdir(exist_ok=True)

    for subfolder in sorted(company_path.iterdir()):
        if not subfolder.is_dir():
            continue
        print(f"📁 Processing subfolder: {subfolder.name}")

        for file in sorted(subfolder.glob('STFCS*.pdf')):
            print(f"📄 Processing STFCS file: {file.name}")
            info = extract_name_and_zip_from_second_page(file)
            if not info:
                continue
            first, last, zip_code = info

            try:
                reader = PdfReader(str(file))
                if len(reader.pages) < 2:
                    print(f"⚠️ Skipping {file.name}: Fewer than 2 pages")
                    continue

                writer = PdfWriter()
                for page in reader.pages[2:]:
                    writer.add_page(page)

                output_filename = generate_random_filename(first, last)
                output_path = state_path / output_filename
                with open(output_path, 'wb') as f:
                    writer.write(f)
                print(f"✅ Saved: {output_filename} to state directory")

            except Exception as e:
                print(f"❌ Error processing {file.name}: {e}")
