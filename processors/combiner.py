import pandas as pd
import re
from datetime import datetime
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter

def combine_state_files(state_path: Path, combined_path: Path):
    """Combine up to 30 PDF files from state directory into combined PDFs, using ignore/order rules."""
    root_path = state_path.parent
    ignore_file = root_path / "ignore.xlsx"
    order_file = root_path / "order.xlsx"
    combined_path.mkdir(exist_ok=True)
    print(f"📁 Ensured combined directory exists: {combined_path}")

    def name_key(first, last, zip_code=None):
        return (first.strip().title(), last.strip().title(), str(zip_code).strip() if zip_code else None)

    # Load ignore list
    ignore_set = set()
    if ignore_file.exists():
        ignore_df = pd.read_excel(ignore_file)
        for _, row in ignore_df.iterrows():
            ignore_set.add(name_key(row['FIRST NAME'], row['LAST NAME']))
        print(f"🚫 Loaded {len(ignore_set)} ignore entries")

    # Load order list
    order_list = []
    order_set = set()
    if order_file.exists():
        order_df = pd.read_excel(order_file)
        for _, row in order_df.iterrows():
            key = name_key(row['FIRST NAME'], row['LAST NAME'], row['ZIP CODE'])
            order_list.append(key)
            order_set.add(key)
        print(f"📋 Loaded {len(order_list)} order entries")

    def extract_info(file):
        match = re.match(r'([A-Za-z\-]+)_([A-Za-z]+)_(\d{6})\.pdf$', file.name)
        if match:
            last, first, digits = match.groups()
            return (file, last.title(), first.title(), int(digits))
        return None

    def get_zip_from_pdf(pdf_path):
        try:
            reader = PdfReader(str(pdf_path))
            text = reader.pages[1].extract_text() if len(reader.pages) >= 2 else ""
            zip_match = re.search(r'\b([A-Z]{2})\s+(\d{5})(?:-\d{4})?\b', text)
            return zip_match.group(2) if zip_match else None
        except:
            return None

    all_files = sorted([f for f in state_path.glob("*.pdf")])
    file_map = {}
    used_files = set()
    for file in all_files:
        info = extract_info(file)
        if info:
            file_map[file.name] = info

    def create_batch(pdf_entries):
        writer = PdfWriter()
        name_list = []
        for file, last, first, _ in pdf_entries:
            try:
                reader = PdfReader(str(file))
                for page in reader.pages:
                    writer.add_page(page)
                name_list.append((last, first))
                print(f"✅ Added {file.name} to combined PDF")
                used_files.add(file.name)
            except Exception as e:
                print(f"❌ Failed to add {file.name}: {e}")
        if name_list:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            output_filename = f"combined_{timestamp}.pdf"
            output_path = combined_path / output_filename
            with open(output_path, 'wb') as f:
                writer.write(f)
            print(f"✅ Saved combined PDF: {output_filename}")
            return {'pdf': output_path, 'names': name_list}
        return None

    combined_info = []

    # Step 1: Process ordered files in batches
    order_batch = []
    for first, last, zip_code in order_list:
        for file, l, f, _ in file_map.values():
            if (f, l, zip_code) == (first, last, zip_code):
                if file.name not in used_files and (f, l) not in ignore_set:
                    order_batch.append((file, l, f, 0))
                    used_files.add(file.name)
                    break
        if len(order_batch) == 30:
            batch = create_batch(order_batch)
            if batch:
                combined_info.append(batch)
            order_batch = []
    if order_batch:
        batch = create_batch(order_batch)
        if batch:
            combined_info.append(batch)

    # Step 2: Fill remaining files in batches of 30
    remaining_files = [info for fname, info in file_map.items()
                       if fname not in used_files and
                          name_key(info[2], info[1]) not in ignore_set and
                          name_key(info[2], info[1], get_zip_from_pdf(info[0])) not in order_set]

    remaining_files.sort(key=lambda x: x[3])  # by 6-digit number

    while remaining_files:
        batch_files = remaining_files[:30]
        batch = create_batch(batch_files)
        if batch:
            combined_info.append(batch)
        remaining_files = remaining_files[30:]

    print("✅ All files combined successfully")
    return combined_info
