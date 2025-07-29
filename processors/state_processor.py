import re
import random
import string
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.errors import PdfReadError
from pathlib import Path

def extract_name_from_second_page(pdf_path: Path):
    """Extract full name from the second page of an STFCS file, with fallback to other pages."""
    try:
        reader = PdfReader(str(pdf_path))
        if len(reader.pages) < 2:
            print(f"⚠️ Skipping {pdf_path.name}: Fewer than 2 pages")
            return None
        text = reader.pages[1].extract_text() or ""
        # Robust regex to handle names with hyphens, apostrophes, etc.
        match = re.search(r'for:\s+([A-Z][A-Za-z\'-]+)\s+([A-Z][A-Za-z\'-]+)', text, re.IGNORECASE)
        if match:
            first, last = match.group(1).strip(), match.group(2).strip()
            print(f"✅ Extracted name from {pdf_path.name} (page 2): {first} {last}")
            return first, last
        # Fallback: Check other pages
        for i, page in enumerate(reader.pages):
            if i == 1:  # Skip second page
                continue
            text = page.extract_text() or ""
            match = re.search(r'for:\s+([A-Z][A-Za-z\'-]+)\s+([A-Z][A-Za-z\'-]+)', text, re.IGNORECASE)
            if match:
                first, last = match.group(1).strip(), match.group(2).strip()
                print(f"✅ Extracted name from {pdf_path.name} (page {i+1}): {first} {last}")
                return first, last
        print(f"⚠️ Skipping {pdf_path.name}: No name found in any page")
        return None
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

def attach_w2_to_stfcs(company_path: Path, w2_path: Path, state_path: Path):
    """Attach the W2 page containing the matching name to STFCS file after stripping first 2 pages."""
    state_path.mkdir(exist_ok=True)
    
    # Get all W2 files in alphabetical order
    w2_files = sorted(w2_path.glob('*.pdf'))
    if not w2_files:
        print("⚠️ No W2 files found in W2 directory. Skipping all processing.")
        return
    print(f"📄 Found {len(w2_files)} W2 files")

    # Process subfolders in alphabetical order
    for subfolder in sorted(company_path.iterdir()):
        if not subfolder.is_dir():
            continue
        print(f"📁 Processing subfolder: {subfolder.name}")

        # Process STFCS files in alphabetical order
        for file in sorted(subfolder.glob('STFCS*.pdf')):
            print(f"📄 Processing STFCS file: {file.name}")
            
            # Extract name from STFCS file
            name = extract_name_from_second_page(file)
            if not name:
                continue
            first, last = name
            first_lower = first.lower()
            last_lower = last.lower()

            # Search all W2 files for the first page containing the name
            matched_w2 = None
            matched_page = None
            matched_page_num = None
            for w2_file in w2_files:
                try:
                    w2_reader = PdfReader(str(w2_file))
                    if len(w2_reader.pages) < 1:
                        print(f"⚠️ Skipping W2 {w2_file.name}: No pages found")
                        continue
                    for i, page in enumerate(w2_reader.pages):
                        text = (page.extract_text() or "").lower()
                        if first_lower in text and last_lower in text:
                            matched_w2 = w2_file
                            matched_page = page
                            matched_page_num = i  # Track page number (0-based)
                            print(f"✅ Found matching W2 for {first} {last} in {w2_file.name} (page {i+1})")
                            break
                    if matched_w2:
                        break
                except PdfReadError as e:
                    print(f"⚠️ Skipping W2 {w2_file.name}: PDF read error - {e}")
                    continue
                except Exception as e:
                    print(f"❌ Skipping W2 {w2_file.name}: Unexpected error - {e}")
                    continue

            if not matched_w2 or not matched_page:
                print(f"⚠️ No matching W2 page found for {first} {last}. Skipping.")
                continue

            try:
                writer = PdfWriter()
                
                # Remove first 2 pages of STFCS
                stfcs_reader = PdfReader(str(file))
                if len(stfcs_reader.pages) < 2:
                    print(f"⚠️ Skipping {file.name}: Fewer than 2 pages")
                    continue
                for page in stfcs_reader.pages[2:]:
                    writer.add_page(page)

                # Append the specific W2 page with the matching name
                writer.add_page(matched_page)

                # Save final PDF
                output_filename = generate_random_filename(first, last)
                output_path = state_path / output_filename
                with open(output_path, 'wb') as f:
                    writer.write(f)
                print(f"✅ Saved: {output_filename} with W2 from {matched_w2.name} (page {matched_page_num+1}) to state directory")

            except PdfReadError as e:
                print(f"❌ Error processing {file.name} with W2 {matched_w2.name}: PDF error - {e}")
            except Exception as e:
                print(f"❌ Error processing {file.name} with W2 {matched_w2.name}: Unexpected error - {e}")