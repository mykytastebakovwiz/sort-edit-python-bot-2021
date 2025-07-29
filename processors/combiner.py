import re
from datetime import datetime
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.errors import PdfReadError

def combine_state_files(state_path: Path, combined_path: Path):
    """Combine up to 30 PDF files from state directory into combined PDFs, returning name lists."""
    combined_path.mkdir(exist_ok=True)
    print(f"📁 Ensured combined directory exists: {combined_path}")

    processed_files = set()
    combined_info = []  # List of dicts: [{'pdf': Path, 'names': [(Last, First), ...]}, ...]

    while True:
        all_files = sorted(state_path.glob('*.pdf'))
        unprocessed_files = [f for f in all_files if f.name not in processed_files]

        if not unprocessed_files:
            print("✅ No unprocessed files remain in state directory. Combining complete.")
            break

        file_info = []
        for file in unprocessed_files:
            match = re.match(r'([A-Za-z\-]+)_([A-Za-z]+)_(\d{6})\.pdf$', file.name)
            if match:
                last, first, digits = match.groups()
                file_info.append((file, last, first, int(digits)))
            else:
                print(f"⚠️ Skipping {file.name}: Invalid filename format")
                processed_files.add(file.name)

        if not file_info:
            break

        file_info.sort(key=lambda x: x[3])  # Sort by 6-digit number
        selected_files = file_info[:30]     # Take up to 30 files

        writer = PdfWriter()
        name_list = []

        for file, last, first, _ in selected_files:
            try:
                reader = PdfReader(str(file))
                if len(reader.pages) < 1:
                    print(f"⚠️ Skipping {file.name}: No pages found")
                    processed_files.add(file.name)
                    continue

                for page in reader.pages:
                    writer.add_page(page)

                name_list.append((last, first))
                print(f"✅ Added {file.name} to combined PDF")

            except PdfReadError as e:
                print(f"❌ {file.name}: PDF read error - {e}")
            except Exception as e:
                print(f"❌ {file.name}: Unexpected error - {e}")
            finally:
                processed_files.add(file.name)

        if not name_list:
            print("⚠️ No valid files in this batch")
            continue

        # ✅ Ensure filename uniqueness with microseconds
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        output_filename = f"combined_{timestamp}.pdf"
        output_path = combined_path / output_filename

        try:
            with open(output_path, 'wb') as f:
                writer.write(f)
            print(f"✅ Saved combined PDF: {output_filename}")

            # ✅ Only append if file save is successful
            combined_info.append({'pdf': output_path, 'names': name_list})

        except Exception as e:
            print(f"❌ Failed to save {output_filename}: {e}")
            continue

    return combined_info
