import fitz  # PyMuPDF
import xml.etree.ElementTree as ET
import os
import re
import csv
from datetime import datetime
import pytesseract
from PIL import Image
import base64

# --- CONFIGURATION ---
INPUT_PDF = "Sep.pdf"
CSV_FILE_PATH = "Sep_Onboarded.csv"  # Ensure this file exists in the same folder
OUTPUT_FOLDER = "Sep_Stamping"
PAGES_PER_CONTRACT = 10
START_REF_NUM = 1
MAX_FILE_SIZE_BYTES = 25 * 1024 * 1024  # 25 MB Limit

# --- STATE CODE MAPPING ---
STATE_MAP = {
    range(79, 87): "1",   # Johor
    range(2, 10):  "2",   # Kedah
    range(15, 19): "3",   # Kelantan
    range(75, 79): "4",   # Melaka
    range(70, 74): "5",   # Negeri Sembilan
    range(25, 29): "6",   # Pahang
    range(39, 40): "6",
    range(69, 70): "6",
    range(30, 37): "7",   # Perak
    range(1, 2):   "8",   # Perlis
    range(10, 15): "9",   # Pulau Pinang
    range(88, 92): "10",  # Sabah
    range(93, 99): "11",  # Sarawak
    range(40, 50): "12",  # Selangor
    range(63, 69): "12",
    range(20, 25): "13",  # Terengganu
    range(50, 61): "14",  # W.P. Kuala Lumpur
    range(87, 88): "15",  # W.P. Labuan
    range(62, 63): "16",  # W.P. Putrajaya
}

# --- CONSTANTS ---
TRANSFEROR_DETAILS = {
    "type": "1",
    "name": "Convergence GBS Sdn Bhd",
    "rocNo": "201601021567 (11292506-K)",
    "busType": "1",
    "street1": "19 Jalan USJ Heights 1/1B",
    "street2": "USJ Heights",
    "street3": "",
    "postcode": "47500",
    "city": "Subang Jaya",
    "state": "12",
    "country": "146",
    "telNo": "0199006888",
    "email": "",
    "pasportNo": "",
    "pasportCountry": "",
    "incomeTaxNo": "",
    "incomeTaxBranch": ""
}

if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

# --- HELPER FUNCTIONS ---

def normalize_name_key(name_str):
    """
    Removes spaces, casing, and special chars to create a reliable lookup key.
    Ex: "Adriana  Anuar" -> "adrianaannuar"
    """
    if not name_str:
        return ""
    # Only keep alphanumeric characters, lowercase
    return re.sub(r'[^a-z0-9]', '', name_str.lower())

def clean_phone_number(phone_str):
    """Removes spaces, dashes, + signs. Returns only digits."""
    if not phone_str:
        return ""
    return re.sub(r'\D', '', phone_str)

def load_csv_database(csv_path):
    """
    Reads the CSV and builds a dictionary where the Key is the normalized name.
    """
    db = {}
    print(f"Loading contact details from {csv_path}...")
    try:
        with open(csv_path, mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                # Adjust these keys if your CSV headers are different
                raw_name = row.get("Name", "")
                raw_phone = row.get("Primary Phone", "")
                raw_email = row.get("Email", "")
                
                key = normalize_name_key(raw_name)
                if key:
                    db[key] = {
                        "telNo": clean_phone_number(raw_phone),
                        "email": raw_email.strip()
                    }
        print(f"  -> Loaded {len(db)} contacts into memory.")
    except FileNotFoundError:
        print("  [!] ERROR: CSV file not found. Phone/Email will be blank.")
    except Exception as e:
        print(f"  [!] ERROR reading CSV: {e}")
    return db

def get_state_code_from_postcode(postcode):
    try:
        prefix = int(postcode[:2])
        for num_range, state_code in STATE_MAP.items():
            if prefix in num_range:
                return state_code
    except:
        pass
    return ""

def format_date(date_str):
    try:
        clean_str = date_str.strip().replace("\n", "")
        clean_str = re.sub(r'[^\w\s]', '', clean_str)
        dt = datetime.strptime(clean_str, "%d %b %Y")
        return dt.strftime("%d/%m/%Y")
    except ValueError:
        return date_str

def create_node(parent, tag, text=None, attrib=None):
    elem = ET.SubElement(parent, tag, attrib if attrib else {})
    elem.text = text if text is not None else ""
    return elem

def extract_from_image_ocr(image_path):
    img = Image.open(image_path)
    text = pytesseract.image_to_string(img, lang='eng')
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    data = {
        "instrumentDate": "", "name": "", "icNo": "",
        "street1": "", "street2": "", "postcode": ""
    }

    ic_pattern_standard = re.compile(r'(\d{6})[\s\-_]+(\d{2})[\s\-_]+(\d{4})')
    ic_pattern_raw = re.compile(r'\b(\d{6})(\d{2})(\d{4})\b')

    for i, line in enumerate(lines):
        clean_line = line.replace('O', '0').replace('|', '1').replace('l', '1')
        match = ic_pattern_standard.search(clean_line)
        if not match:
            match = ic_pattern_raw.search(clean_line)
        
        if match:
            data["icNo"] = f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
            if i > 0:
                raw_name = lines[i-1]
                data["name"] = re.sub(r'(?i)^(Name|Employee|Name of Employee)[:\-\s]*', '', raw_name).strip()
            if i + 1 < len(lines): data["street1"] = lines[i+1]
            if i + 2 < len(lines): data["street2"] = lines[i+2]
            for j in range(1, 6):
                if i + j < len(lines):
                    pc_match = re.search(r'\b\d{5}\b', lines[i+j])
                    if pc_match:
                        data["postcode"] = pc_match.group(0)
                        break
            break 

    if lines:
        data["instrumentDate"] = format_date(lines[0])
    return data

def save_batch_file(root_element, batch_number):
    tree = ET.ElementTree(root_element)
    ET.indent(tree, space="  ", level=0)
    filename = f"Submission_Batch_{batch_number}.xml"
    out_path = os.path.join(OUTPUT_FOLDER, filename)
    try:
        tree.write(out_path, encoding="utf-8", xml_declaration=True, short_empty_elements=False)
    except TypeError:
        tree.write(out_path, encoding="utf-8", xml_declaration=True)
    print(f"  [SAVED] {filename} (Batch limit reached)")

def init_new_xml_root():
    root = ET.Element("bulkstamping")
    create_node(root, "applicationType", "44")
    return root

def main():
    # 1. LOAD CSV DATA BEFORE STARTING
    contact_db = load_csv_database(CSV_FILE_PATH)
    
    print(f"Processing {INPUT_PDF} with {MAX_FILE_SIZE_BYTES / (1024*1024)}MB split limit...")
    
    doc = fitz.open(INPUT_PDF)
    total_pages = len(doc)
    
    current_batch_num = 1
    current_root = init_new_xml_root()
    current_xml_size = 500
    
    for i in range(0, total_pages, PAGES_PER_CONTRACT):
        contract_num = (i // PAGES_PER_CONTRACT) + 1
        current_ref_num = START_REF_NUM + (contract_num - 1)
        ref_no = f"CON/2025/05/{current_ref_num}"
        type_desc = f"Employment Contract {ref_no}"
        
        print(f"  Processing Contract #{contract_num}...", end=" ")

        # A. IMAGE & BASE64
        page = doc.load_page(i)
        pix = page.get_pixmap(dpi=300)
        image_filename = f"DOC_{current_ref_num}.jpeg"
        image_path = os.path.join(OUTPUT_FOLDER, image_filename)
        pix.save(image_path)
        
        with open(image_path, "rb") as image_file:
            base64_string = base64.b64encode(image_file.read()).decode('utf-8')
        
        # B. SIZE CHECK & SPLIT
        estimated_entry_size = len(base64_string) + 5000 
        if (current_xml_size + estimated_entry_size) > MAX_FILE_SIZE_BYTES:
            save_batch_file(current_root, current_batch_num)
            current_batch_num += 1
            current_root = init_new_xml_root()
            current_xml_size = 500
            print(f"    -> Starting Batch #{current_batch_num}...")
        
        # C. EXTRACT DATA & CSV LOOKUP
        extracted = extract_from_image_ocr(image_path)
        detected_state_code = get_state_code_from_postcode(extracted["postcode"])
        
        # --- LOOKUP LOGIC ---
        found_tel = ""
        found_email = ""
        
        normalized_key = normalize_name_key(extracted["name"])
        
        if normalized_key in contact_db:
            found_tel = contact_db[normalized_key]["telNo"]
            found_email = contact_db[normalized_key]["email"]
            # Debug print to confirm match
            # print(f"    [MATCH] {extracted['name']} -> {found_tel} / {found_email}")
        else:
            print(f"    [!] No CSV match for: '{extracted['name']}'")

        # D. BUILD XML
        instrument = ET.SubElement(current_root, "instrument")
        create_node(instrument, "refNo", ref_no)
        create_node(instrument, "instrumentDate", extracted["instrumentDate"])
        create_node(instrument, "instrumentDateReceive", "")
        create_node(instrument, "typeOfInstrument", "")
        create_node(instrument, "typeOfInstrumentOthers", type_desc)
        
        # Transferor
        transferor = ET.SubElement(instrument, "transferor")
        for key, val in TRANSFEROR_DETAILS.items():
            create_node(transferor, key, val)
        
        # Transferee
        transferee = ET.SubElement(instrument, "transferee")
        create_node(transferee, "type", "0")
        create_node(transferee, "name", extracted["name"])
        create_node(transferee, "nationality", "1")
        create_node(transferee, "icNo", extracted["icNo"])
        create_node(transferee, "pasportNo", "")
        create_node(transferee, "pasportCountry", "")
        create_node(transferee, "rocNo", "")
        create_node(transferee, "busType", "")
        create_node(transferee, "incomeTaxNo", "")
        create_node(transferee, "incomeTaxBranch", "")
        create_node(transferee, "street1", extracted["street1"])
        create_node(transferee, "street2", extracted["street2"])
        create_node(transferee, "street3", "")
        create_node(transferee, "postcode", extracted["postcode"])
        create_node(transferee, "city", "")
        create_node(transferee, "state", detected_state_code)
        create_node(transferee, "country", "146")
        
        # INJECT CSV DATA HERE
        create_node(transferee, "telNo", found_tel)
        create_node(transferee, "email", found_email)

        # Footer
        create_node(instrument, "noOfCopy", "")
        create_node(instrument, "remessionOrExemption", "")
        create_node(instrument, "payment", "")
        create_node(instrument, "aggrementInfo", "")
        create_node(instrument, "attachment", base64_string, attrib={"name": image_filename})

        current_xml_size += estimated_entry_size
        print("Done.")

    doc.close()
    save_batch_file(current_root, current_batch_num)
    print(f"\nAll Done! Batches saved in '{OUTPUT_FOLDER}'")

if __name__ == "__main__":
    main()
