import os
import sys
import re
import glob
import subprocess
import json
import sqlite3
import shutil
from datetime import datetime

MAPPINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mappings.json")

def load_mappings():
    if os.path.exists(MAPPINGS_FILE):
        try:
            with open(MAPPINGS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if not isinstance(data, dict):
                    data = {}
                if "signatures" not in data:
                    data["signatures"] = {}
                if "settings" not in data:
                    data["settings"] = {}
                return data
        except Exception as e:
            print(f"Error loading mappings file: {e}")
    return {"signatures": {}, "settings": {}}

def save_mappings(mappings):
    try:
        with open(MAPPINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(mappings, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Error saving mappings file: {e}")

CONFIG = load_mappings()

# Dependency check & installation
def install_and_import(package, import_name=None):
    if import_name is None:
        import_name = package
    try:
        __import__(import_name)
    except ImportError:
        print(f"Installing missing dependency: {package}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        except Exception as e:
            print(f"Failed to install {package}: {e}")
            sys.exit(1)

install_and_import("pandas")
install_and_import("openpyxl")
install_and_import("pdfplumber")
install_and_import("pypdf")
install_and_import("xlrd")

import pandas as pd
import openpyxl
import pdfplumber

# Path Configuration
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(SCRIPT_DIR, "storico_prezzi.db")

def get_input_dir():
    path = CONFIG["settings"].get("input_dir", os.path.abspath(os.path.join(SCRIPT_DIR, "..", "File da lavorare")))
    os.makedirs(path, exist_ok=True)
    return path

def get_output_dir():
    path = CONFIG["settings"].get("output_dir", os.path.abspath(os.path.join(SCRIPT_DIR, "..", "File Lavorati")))
    os.makedirs(path, exist_ok=True)
    return path

def get_filename_pattern():
    return CONFIG["settings"].get("filename_pattern", "estrazione articoli {supplier_name} del {date}")

# Auto-create default directories initially
get_input_dir()
get_output_dir()

# Try to find the template path dynamically
home_dir = os.path.expanduser("~")
possible_template_paths = [
    os.path.join(SCRIPT_DIR, "Modello importazione.xlsx"),
    os.path.join(home_dir, "OneDrive", "Cartella di lavoro", "Importazioni in Danea", "Modello importazione.xlsx"),
    os.path.join(home_dir, "OneDrive - Personal", "Cartella di lavoro", "Importazioni in Danea", "Modello importazione.xlsx"),
    os.path.join(home_dir, "OneDrive - Personal", "Documenti", "Importazioni in Danea", "Modello importazione.xlsx")
]

TEMPLATE_PATH = None
for path in possible_template_paths:
    if os.path.exists(path):
        TEMPLATE_PATH = path
        break

if TEMPLATE_PATH is None:
    TEMPLATE_PATH = os.path.join(SCRIPT_DIR, "Modello importazione.xlsx")

# --- DATABASE AND ARCHIVING LOGIC ---

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS storico_prezzi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier TEXT NOT NULL,
            product_code TEXT NOT NULL,
            description TEXT,
            price_net REAL NOT NULL,
            date_processed TEXT NOT NULL
        )
    """)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_supp_code ON storico_prezzi (supplier, product_code)")
    conn.commit()
    conn.close()

# Auto-initialize DB on module import
init_db()

def get_last_price(supplier, product_code):
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT price_net FROM storico_prezzi 
            WHERE supplier = ? AND product_code = ? 
            ORDER BY id DESC LIMIT 1
        """, (supplier, product_code))
        row = cursor.fetchone()
        conn.close()
        return row[0] if row else None
    except Exception as e:
        print(f"Error querying last price: {e}")
        return None

def update_price(supplier, product_code, description, price_net):
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO storico_prezzi (supplier, product_code, description, price_net, date_processed)
            VALUES (?, ?, ?, ?, ?)
        """, (supplier, product_code, description, price_net, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Error updating price history: {e}")



def parse_float(val):
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace('.', '').replace(',', '.')
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def clean_vat(vat_val):
    if vat_val is None or pd.isna(vat_val):
        return ''
    vat_str = str(vat_val).strip()
    digits = re.findall(r'\d+', vat_str)
    if digits:
        return str(int("".join(digits)))
    return vat_str

def create_danea_import(rows, supplier_name):
    """
    Creates an Excel file starting from the Danea template and populates it.
    """
    if not os.path.exists(TEMPLATE_PATH):
        print(f"Error: Template file not found at {TEMPLATE_PATH}")
        return None
    
    # Load template
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active
    
    # Clear headers that cause warnings or unwanted mapping in Danea
    ws.cell(row=1, column=1, value='')   # Clear "Cod."
    ws.cell(row=1, column=4, value='')   # Clear "Taglia/Colore"
    ws.cell(row=1, column=13, value='')  # Clear "Note"
    
    # Find start row (usually 2, after headers)
    # The columns are: Cod., Cod. prod. forn., Descrizione, Taglia/Colore, Q.tà, ...
    # We will append the rows starting at row 2
    
    # Clear any dummy data starting from row 2
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)
        
    for r_idx, data in enumerate(rows, start=2):
        code = data.get('Cod. prod. forn.', '')
        desc = data.get('Descrizione', '')
        new_price = parse_float(data.get('Prezzo netto', 0.0))
        note = data.get('Note', '')
        
        # Price check logic
        if code and new_price > 0:
            last_price = get_last_price(supplier_name, code)
            if last_price is not None and new_price > last_price:
                diff_pct = ((new_price - last_price) / last_price) * 100
                # Only warn if increase is significant (e.g. > 0.1%)
                if diff_pct > 0.1:
                    price_alert = f"AUMENTATO: +{diff_pct:.1f}% (da {last_price:.2f} a {new_price:.2f})"
                    if note:
                        note = f"{price_alert} | {note}"
                    else:
                        note = price_alert
                    print(f"Price Warning for {supplier_name} Code {code}: {price_alert}")
            
            # Update DB with the new price
            update_price(supplier_name, code, desc, new_price)

        ws.cell(row=r_idx, column=1, value='') # Always keep Cod. empty as requested
        ws.cell(row=r_idx, column=2, value=code)
        ws.cell(row=r_idx, column=3, value=desc)
        ws.cell(row=r_idx, column=4, value=data.get('Taglia/Colore', ''))
        ws.cell(row=r_idx, column=5, value=data.get('Q.tà', ''))
        ws.cell(row=r_idx, column=6, value='') # Q.tà disponibile
        ws.cell(row=r_idx, column=7, value=new_price)
        ws.cell(row=r_idx, column=8, value=data.get('U.m.', ''))
        ws.cell(row=r_idx, column=9, value=data.get('Sconti', ''))
        ws.cell(row=r_idx, column=10, value=clean_vat(data.get('Iva', '')))
        ws.cell(row=r_idx, column=11, value='') # Mag.
        ws.cell(row=r_idx, column=12, value='') # Importo
        ws.cell(row=r_idx, column=13, value=note)

    # Delete empty columns (from right to left)
    max_col = ws.max_column
    for col_idx in range(max_col, 0, -1):
        is_empty = True
        for row_idx in range(2, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is not None and str(val).strip() != "":
                is_empty = False
                break
        if is_empty:
            ws.delete_cols(col_idx)

    out_dir = get_output_dir()
    os.makedirs(out_dir, exist_ok=True)
    today_str = datetime.now().strftime("%d-%m-%Y")
    
    pattern = get_filename_pattern()
    filename_base = pattern.replace("{supplier_name}", supplier_name).replace("{date}", today_str)
    
    output_filename = f"{filename_base}.xlsx"
    output_path = os.path.join(out_dir, output_filename)
    counter = 1
    while os.path.exists(output_path):
        output_filename = f"{filename_base} ({counter}).xlsx"
        output_path = os.path.join(out_dir, output_filename)
        counter += 1
        
    wb.save(output_path)
    print(f"Successfully generated Danea import file: {output_path}")
    return output_path

# --- PARSING FUNCTIONS ---

def parse_ferrari(file_path):
    print(f"Parsing Ferrari file: {file_path}")
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return []
        
    rows = []
    # Check if necessary columns exist
    required = ['Barcode pezzo', 'Prezzo cessione']
    if not all(col in df.columns for col in required):
        print("Missing required columns for Ferrari template.")
        return []
        
    for _, row in df.iterrows():
        # Code: Barcode pezzo (EAN) or Codice prodotto / Codice CDF
        barcode = str(row.get('Barcode pezzo', '')).strip()
        if not barcode or barcode.lower() == 'nan':
            barcode = str(row.get('Codice prodotto', '')).strip()
            
        cod_prod_forn = str(row.get('Codice CDF', '')).strip()
        desc = str(row.get('Descrizione prodotto ', '')).strip()
        qty = parse_float(row.get('Quantità', 1))
        um = str(row.get('Unità di misura', '')).strip()
        
        # Net price calculation
        prezzo_cessione = parse_float(row.get('Prezzo cessione', 0.0))
        sconto1 = parse_float(row.get('1° Sconto', 0.0))
        sconto2 = parse_float(row.get('2° Sconto', 0.0))
        
        net_price = prezzo_cessione * (1 - sconto1/100) * (1 - sconto2/100)
        
        sconti_str = ""
        if sconto1 > 0:
            sconti_str += f"{sconto1:.0f}"
        if sconto2 > 0:
            sconti_str += f"+{sconto2:.0f}"
            
        rows.append({
            'Cod.': barcode,
            'Cod. prod. forn.': cod_prod_forn,
            'Descrizione': desc,
            'Q.tà': qty,
            'Prezzo netto': round(net_price, 4),
            'U.m.': um,
            'Sconti': sconti_str,
            'Iva': row.get('Iva', 22),
            'Note': 'Ferrari Import'
        })
    return rows

def parse_viridex(file_path):
    print(f"Parsing Viridex file: {file_path}")
    try:
        # Load raw sheet
        df_raw = pd.read_excel(file_path, header=None)
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return []
        
    # Find the header row (contains 'Codice')
    header_row_idx = None
    for idx, row in df_raw.iterrows():
        row_vals = [str(x).strip().lower() for x in row.tolist()]
        if 'codice' in row_vals and 'prezzo' in row_vals:
            header_row_idx = idx
            break
            
    if header_row_idx is None:
        print("Could not find Viridex header row containing 'Codice' and 'Prezzo'.")
        return []
        
    df = pd.read_excel(file_path, skiprows=header_row_idx)
    # Standardize column names (remove whitespaces)
    df.columns = [str(c).strip() for c in df.columns]
    
    rows = []
    for _, row in df.iterrows():
        # Skip total rows or empty rows
        code = str(row.get('Codice', '')).strip()
        if not code or code.lower() == 'nan' or 'ordine' in code.lower() or 'contributo' in str(row.get('Descrizione', '')).lower():
            continue
            
        barcode = str(row.get('Codice a Barre', '')).strip()
        if not barcode or barcode.lower() == 'nan':
            barcode = code
            
        desc = str(row.get('Descrizione', '')).strip()
        qty = parse_float(row.get('Qtà', 1))
        um = str(row.get('Um', '')).strip()
        
        prezzo = parse_float(row.get('Prezzo', 0.0))
        sconto = parse_float(row.get('Sconto%', 0.0))
        
        net_price = prezzo * (1 - sconto/100)
        sconti_str = f"{sconto:.0f}" if sconto > 0 else ""
        
        rows.append({
            'Cod.': barcode,
            'Cod. prod. forn.': code,
            'Descrizione': desc,
            'Q.tà': qty,
            'Prezzo netto': round(net_price, 4),
            'U.m.': um,
            'Sconti': sconti_str,
            'Iva': str(row.get('IVA', '22%')).replace('%', '').strip(),
            'Note': 'Viridex Import'
        })
    return rows

def parse_pilato(file_path):
    print(f"Parsing Pilato file: {file_path}")
    rows = []
    
    if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
        try:
            df = pd.read_excel(file_path)
            df.columns = [str(c).strip() for c in df.columns]
            for _, row in df.iterrows():
                code = str(row.get('Cod. prod. forn.', '')).strip()
                if not code or code.lower() == 'nan':
                    continue
                qty = parse_float(row.get('Q.tà', 1))
                net_price = parse_float(row.get('Prezzo netto', 0.0))
                rows.append({
                    'Cod.': code,
                    'Cod. prod. forn.': code,
                    'Q.tà': qty,
                    'Prezzo netto': net_price,
                    'Note': 'Pilato Excel Import'
                })
        except Exception as e:
            print(f"Error reading Pilato Excel: {e}")
            
    elif file_path.endswith('.txt') or file_path.endswith('.csv'):
        # Semi-colon separated file
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                for line in f:
                    parts = [p.strip() for p in line.split(';')]
                    if len(parts) >= 10:
                        code_prod = parts[3]
                        ean = parts[4]
                        desc = parts[5]
                        qty = parse_float(parts[7]) / 100 # e.g. 00000000600 is 6.00
                        price_raw = parse_float(parts[8]) / 10000 # e.g. 00000023200 is 2.32
                        vat = parts[9]
                        um = parts[16]
                        
                        rows.append({
                            'Cod.': '',
                            'Cod. prod. forn.': ean if ean else code_prod,
                            'Descrizione': desc,
                            'Q.tà': qty,
                            'Prezzo netto': round(price_raw, 4),
                            'U.m.': '',
                            'Iva': vat,
                            'Note': 'Pilato TXT Import'
                        })
        except Exception as e:
            print(f"Error reading Pilato TXT: {e}")
            
    return rows

def parse_monge(file_path):
    print(f"Parsing Monge PDF: {file_path}")
    rows = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                for line in text.split('\n'):
                    cleaned_line = line.strip()
                    if re.match(r'^\d{12,14}\b', cleaned_line):
                        # Split by NR (allowing spaces between N and R)
                        nr_match = re.search(r'\bN\s*R\b', cleaned_line, re.IGNORECASE)
                        if nr_match:
                            left_part = cleaned_line[:nr_match.start()].strip()
                            right_part = cleaned_line[nr_match.end():].strip()
                            
                            # Clean left part description & packaging
                            left_cleaned = re.sub(r'\bK\s+g\b', 'Kg', left_part, flags=re.IGNORECASE)
                            left_cleaned = re.sub(r'\bg\s+r\b', 'gr', left_cleaned, flags=re.IGNORECASE)
                            
                            def clean_pkg(match):
                                return match.group(0).replace(" ", "")
                            left_cleaned = re.sub(r'\b(?:\d\s*)+x\s*(?:\d\s*)+\b', clean_pkg, left_cleaned)
                            left_cleaned = re.sub(r'Kg\s*\.\s*', 'Kg.', left_cleaned, flags=re.IGNORECASE)
                            left_cleaned = re.sub(r'gr\s*\.\s*', 'gr.', left_cleaned, flags=re.IGNORECASE)
                            
                            # Reconstruct tokens
                            left_parts = left_cleaned.split()
                            ean = left_parts[0]
                            desc = " ".join(left_parts[1:])
                            
                            right_parts = right_part.split()
                            if not right_parts:
                                continue
                            
                            try:
                                vat = right_parts[-1]
                                total_amount = parse_float(right_parts[-2])
                                
                                discount_chain = ""
                                unit_price = 0.0
                                
                                # Determine if discount chain is present
                                if '+' in right_parts[-3]:
                                    discount_chain = right_parts[-3]
                                    unit_price = parse_float(right_parts[-4])
                                    qty_tokens = right_parts[:-4]
                                else:
                                    unit_price = parse_float(right_parts[-3])
                                    qty_tokens = right_parts[:-3]
                                    
                                qty = parse_float("".join(qty_tokens))
                                
                                # Calculate Net Price
                                net_price = unit_price
                                if discount_chain:
                                    for d in [parse_float(x) for x in discount_chain.split('+')]:
                                        net_price *= (1 - d/100)
                                        
                                # Adjust for Packaging Multiplier (e.g. Kg.3x4)
                                mult_match = re.search(r'\b(?:Kg|gr)\.\d+(?:,\d+)?x(\d+)\b', desc, re.IGNORECASE)
                                multiplier = 1
                                if mult_match:
                                    multiplier = int(mult_match.group(1))
                                    if multiplier > 1:
                                        qty = qty * multiplier
                                        net_price = net_price / multiplier
                                        
                                rows.append({
                                    'Cod.': '',
                                    'Cod. prod. forn.': ean,
                                    'Descrizione': desc,
                                    'Q.tà': qty,
                                    'Prezzo netto': round(net_price, 4),
                                    'U.m.': 'NR',
                                    'Sconti': discount_chain,
                                    'Iva': vat,
                                    'Note': 'Monge Import'
                                })
                            except Exception as inner_e:
                                print(f"Error parsing Monge line data: {inner_e}")
    except Exception as e:
        print(f"Error parsing Monge PDF: {e}")
    return rows

def parse_capaldo(file_path):
    print(f"Parsing Capaldo file: {file_path}")
    rows = []
    
    if file_path.endswith('.pdf') or file_path.endswith('.PDF'):
        try:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if not text:
                        continue
                    for line in text.split('\n'):
                        cleaned = line.strip()
                        if cleaned.startswith('*'):
                            cleaned = cleaned[1:].strip()
                            
                        # Starts with 5-7 digit code
                        match = re.match(r'^(\d{5,7})\b', cleaned)
                        if match:
                            code = match.group(1)
                            parts = cleaned.split()
                            if 'OMAGGIO' in parts or 'GRATIS' in cleaned.upper():
                                rows.append({
                                    'Cod.': code,
                                    'Cod. prod. forn.': code,
                                    'Prezzo netto': 0.0,
                                    'Note': 'Capaldo Omaggio'
                                })
                            else:
                                try:
                                    qty = parse_float(parts[-4])
                                    unit_price = parse_float(parts[-3])
                                    total_net = parse_float(parts[-2])
                                    vat = parts[-1]
                                    
                                    unit_net = total_net / qty if qty > 0 else 0.0
                                    rows.append({
                                        'Cod.': code,
                                        'Cod. prod. forn.': code,
                                        'Q.tà': qty,
                                        'Prezzo netto': round(unit_net, 4),
                                        'Iva': vat,
                                        'Note': 'Capaldo PDF Import'
                                    })
                                except:
                                    pass
        except Exception as e:
            print(f"Error parsing Capaldo PDF: {e}")
            
    elif file_path.endswith('.csv') or file_path.endswith('.CSV'):
        try:
            # Semicolon separated Capaldo CSV
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                df = pd.read_csv(f, sep=';')
            df.columns = [str(c).strip() for c in df.columns]
            for _, row in df.iterrows():
                # 'Nr. Articolo' is the code
                code = str(row.get('Nr. Articolo', '')).strip()
                if not code or code.lower() == 'nan':
                    continue
                qty = parse_float(row.get('Quantita', 1))
                rows.append({
                    'Cod.': code,
                    'Cod. prod. forn.': code,
                    'Q.tà': qty,
                    'Prezzo netto': 0.0,  # CSV has no price info, set to 0 or leave for manual update
                    'Note': 'Capaldo CSV Import (No Price)'
                })
        except Exception as e:
            print(f"Error parsing Capaldo CSV: {e}")
            
    return rows

# --- NEW PARSING FUNCTIONS ---

def parse_blumen(file_path):
    print(f"Parsing Blumen PDF: {file_path}")
    rows = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                for line in text.split('\n'):
                    cleaned = line.strip()
                    # Matches items starting with code and a 13-digit EAN
                    match = re.match(r'^([A-Z0-9]+)\s+(\d{13})', cleaned)
                    if match:
                        code, ean = match.group(1), match.group(2)
                        parts = cleaned.split()
                        try:
                            # Blumen rightmost pricing: VAT, Net Amount, Price, Quantity
                            qty = parse_float(parts[-4])
                            amount = parse_float(parts[-2])
                            net_price = amount / qty if qty > 0 else 0.0
                            
                            rows.append({
                                'Cod.': ean,
                                'Cod. prod. forn.': code,
                                'Q.tà': qty,
                                'Prezzo netto': round(net_price, 4),
                                'Iva': parts[-1],
                                'Note': 'Blumen PDF Import'
                            })
                        except:
                            pass
    except Exception as e:
        print(f"Error parsing Blumen PDF: {e}")
    return rows

def parse_verisem(file_path):
    print(f"Parsing Verisem PDF: {file_path}")
    rows = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                for line in text.split('\n'):
                    cleaned = line.strip()
                    # Matches Verisem items (e.g. N08155 8009197081559)
                    match = re.match(r'^([A-Z0-9]+)\s+(\d{13})', cleaned)
                    if match:
                        code, ean = match.group(1), match.group(2)
                        parts = cleaned.split()
                        try:
                            # Check if discount exists (5 numeric tokens at end vs 4)
                            # Let's verify number of tokens or match prices
                            # If there's discount: [qty, price, discount, amount, vat]
                            # If no discount: [qty, price, amount, vat]
                            # We can also compute net price directly as amount / qty if qty > 0
                            qty = parse_float(parts[-5])
                            price = parse_float(parts[-4])
                            discount = parse_float(parts[-3])
                            amount = parse_float(parts[-2])
                            
                            # Double check if parts[-3] is indeed discount
                            # by verifying if qty * price * (1 - discount/100) approx amount
                            if abs(qty * price * (1 - discount/100) - amount) < 1.0:
                                net_price = price * (1 - discount/100)
                            else:
                                # Fallback: No discount, parts[-3] is price, parts[-4] is qty
                                qty = parse_float(parts[-4])
                                price = parse_float(parts[-3])
                                net_price = price
                                qty = parse_float(parts[-4])
                            
                            rows.append({
                                'Cod.': ean,
                                'Cod. prod. forn.': code,
                                'Q.tà': qty,
                                'Prezzo netto': round(net_price, 4),
                                'Iva': parts[-1],
                                'Note': 'Verisem PDF Import'
                            })
                        except:
                            pass
    except Exception as e:
        print(f"Error parsing Verisem PDF: {e}")
    return rows

def parse_giardini(file_path):
    print(f"Parsing Giardini PDF: {file_path}")
    rows = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                for line in text.split('\n'):
                    cleaned = line.strip()
                    if cleaned.startswith("CHICCOMIO"):
                        parts = cleaned.split()
                        code = parts[0]
                        try:
                            # Format: Code, Desc..., KG, Weight, Bags, total_qty_kg, unit_price, total_amount, VAT
                            # e.g.: CHICCOMIO.25 ... KG 25 144 25 3.600,00 0,3500 1.260,00 4
                            qty_kg = parse_float(parts[-4])
                            unit_price = parse_float(parts[-3])
                            amount = parse_float(parts[-2])
                            vat = parts[-1]
                            
                            rows.append({
                                'Cod.': code,
                                'Cod. prod. forn.': code,
                                'Q.tà': qty_kg,
                                'Prezzo netto': unit_price,
                                'Iva': vat,
                                'Note': 'Giardini PDF Import'
                            })
                        except:
                            pass
    except Exception as e:
        print(f"Error parsing Giardini PDF: {e}")
    return rows

def parse_garver(file_path):
    print(f"Parsing Garver PDF: {file_path}")
    rows = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                for line in text.split('\n'):
                    cleaned = line.strip()
                    parts = cleaned.split()
                    # Match code ending with a hyphen specifically to avoid zip codes
                    if parts and re.match(r'^\d{5,8}-$', parts[0]):
                        code = parts[0].replace('-', '')
                        try:
                            # Format: Code Desc... UM Qty Price Sconto NetPrice Amount VAT
                            # e.g. with discount: 255526- ... PZ. 10,00 4,08000 -55,00 1,83600 18,36 22
                            # e.g. no discount: 702400- ... PZ. 25,00 6,70000 6,70000 167,50 22
                            if '-' in parts[-4]:
                                qty = parse_float(parts[-6])
                                list_price = parse_float(parts[-5])
                                discount = parse_float(parts[-4])
                                net_price = parse_float(parts[-3])
                            else:
                                qty = parse_float(parts[-5])
                                list_price = parse_float(parts[-4])
                                net_price = parse_float(parts[-3])
                            vat = parts[-1]
                            
                            rows.append({
                                'Cod.': '',
                                'Cod. prod. forn.': code,
                                'Q.tà': qty,
                                'Prezzo netto': net_price,
                                'Iva': vat,
                                'Note': 'Garver PDF Import'
                            })
                        except:
                            pass
    except Exception as e:
        print(f"Error parsing Garver PDF: {e}")
    return rows

def parse_record(file_path):
    print(f"Parsing Record PDF: {file_path}")
    rows = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                for line in text.split('\n'):
                    cleaned = line.strip()
                    # Line starts with 4-5 digits product code
                    match = re.match(r'^(\d{4,5})\b', cleaned)
                    if match:
                        code = match.group(1)
                        # Find 13 digit EAN
                        ean_match = re.search(r'\b(\d{13})\b', cleaned)
                        if ean_match:
                            ean = ean_match.group(1)
                            parts = cleaned.split()
                            try:
                                # Rightmost fields: VAT, Amount, Price/Discount, Qty
                                # Format can contain discount: [Qty, Price, Discount, Amount, VAT]
                                # or no discount: [Qty, Price, Amount, VAT]
                                qty = parse_float(parts[-4])
                                price = parse_float(parts[-3])
                                amount = parse_float(parts[-2])
                                vat = parts[-1]
                                
                                # Verify if parts[-3] is discount or unit price
                                # If there's 5 numeric fields at the end:
                                # parts[-5] is Qty, parts[-4] is Price, parts[-3] is Discount, parts[-2] is Amount, parts[-1] is VAT
                                try_qty = parse_float(parts[-5])
                                try_price = parse_float(parts[-4])
                                try_disc = parse_float(parts[-3])
                                if abs(try_qty * try_price * (1 - try_disc/100) - amount) < 1.0:
                                    net_price = try_price * (1 - try_disc/100)
                                    qty = try_qty
                                else:
                                    net_price = price
                                    
                                rows.append({
                                    'Cod.': ean,
                                    'Cod. prod. forn.': code,
                                    'Q.tà': qty,
                                    'Prezzo netto': round(net_price, 4),
                                    'Iva': vat,
                                    'Note': 'Record PDF Import'
                                })
                            except:
                                pass
    except Exception as e:
        print(f"Error parsing Record PDF: {e}")
    return rows

def parse_ricci(file_path):
    print(f"Parsing Ricci PDF: {file_path}")
    rows = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                for line in text.split('\n'):
                    cleaned = line.strip()
                    if "PLE" in cleaned:
                        parts = cleaned.split()
                        code = ""
                        for p in parts:
                            if p.startswith("PLE"):
                                code = p
                                break
                        if code:
                            try:
                                # Format: Date Code Desc... UM Qty Price Amount VAT
                                # e.g.: 20/07/2026 PLE1657S15LZ26 ... TON. 23,760 399,00 9.480,24 22
                                qty_tons = parse_float(parts[-4])
                                price_per_ton = parse_float(parts[-3])
                                vat = parts[-1]
                                
                                # Convert to 15kg bags
                                qty_bags = (qty_tons * 1000) / 15
                                price_per_bag = price_per_ton * 15 / 1000
                                
                                rows.append({
                                    'Cod.': '',
                                    'Cod. prod. forn.': code,
                                    'Q.tà': round(qty_bags, 2),
                                    'Prezzo netto': round(price_per_bag, 4),
                                    'Iva': vat,
                                    'Note': 'Ricci PDF Import'
                                })
                            except:
                                pass
    except Exception as e:
        print(f"Error parsing Ricci PDF: {e}")
    return rows

def parse_generic_excel_csv(file_path):
    print(f"Attempting generic Excel/CSV parsing: {file_path}")
    rows = []
    try:
        if file_path.endswith('.csv') or file_path.endswith('.CSV'):
            # Detect separator
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                first_line = f.readline()
                sep = ';' if ';' in first_line else ','
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                df_raw = pd.read_csv(f, sep=sep, header=None)
        else:
            df_raw = pd.read_excel(file_path, header=None)
            
        header_row_idx = None
        col_mapping = None
        
        # Load mappings database
        mappings_db = load_mappings()
        signatures = mappings_db.get("signatures", {})
        
        # Check if any row matches a stored signature
        for idx, row in df_raw.iterrows():
            row_str = [str(x).strip().lower() for x in row.tolist()]
            signature = "|".join(row_str)
            if signature in signatures:
                header_row_idx = idx
                col_mapping = signatures[signature]["col_mapping"]
                print(f"Matched stored signature for supplier {signatures[signature]['supplier']}: {signature}")
                break
                
        # If no stored signature, run discovery
        if col_mapping is None:
            # Extended keywords
            code_keywords = ['codice', 'ean', 'barcode', 'art.', 'articolo', 'cod. prod', 'cod_prod', 'cod.art']
            desc_keywords = ['descrizione', 'descr', 'articolo', 'denominazione', 'prodotto', 'nome']
            qty_keywords = ['quantita', 'quantità', 'q.ta', 'q.tà', 'qta', 'qty', 'volume', 'pezzi']
            price_keywords = ['prezzo', 'listino', 'netto', 'cessione', 'unitario', 'cad.', 'importo', 'valore']
            disc_keywords = ['sconto', 'sconti', 'disc', 'sct', 'sc%']
            vat_keywords = ['iva', 'aliquota', 'cod.iva', 'iva%']
            
            for idx, row in df_raw.iterrows():
                row_str = [str(x).strip().lower() for x in row.tolist()]
                has_code = any(any(kw in val for kw in code_keywords) for val in row_str if val)
                has_price = any(any(kw in val for kw in price_keywords) for val in row_str if val)
                if has_code and has_price:
                    header_row_idx = idx
                    discovered_mapping = {}
                    for col_idx, val in enumerate(row_str):
                        if not val:
                            continue
                        if any(kw in val for kw in code_keywords) and 'code' not in discovered_mapping:
                            discovered_mapping['code'] = col_idx
                        elif any(kw in val for kw in desc_keywords) and 'desc' not in discovered_mapping:
                            discovered_mapping['desc'] = col_idx
                        elif any(kw in val for kw in qty_keywords) and 'qty' not in discovered_mapping:
                            discovered_mapping['qty'] = col_idx
                        elif any(kw in val for kw in price_keywords) and 'price' not in discovered_mapping:
                            if 'netto' in val:
                                discovered_mapping['net_price'] = col_idx
                            elif 'price' not in discovered_mapping:
                                discovered_mapping['price'] = col_idx
                        elif any(kw in val for kw in disc_keywords) and 'disc' not in discovered_mapping:
                            discovered_mapping['disc'] = col_idx
                        elif any(kw in val for kw in vat_keywords) and 'vat' not in discovered_mapping:
                            discovered_mapping['vat'] = col_idx
                    
                    if 'code' in discovered_mapping or 'desc' in discovered_mapping:
                        col_mapping = discovered_mapping
                        # Save new signature to mappings database
                        signature = "|".join(row_str)
                        filename = os.path.basename(file_path)
                        supplier_name = filename.split('.')[0].replace('estrazione', '').replace('articoli', '').strip()
                        if not supplier_name:
                            supplier_name = f"Generico_{len(signatures) + 1}"
                        
                        signatures[signature] = {
                            "supplier": supplier_name,
                            "col_mapping": col_mapping
                        }
                        mappings_db["signatures"] = signatures
                        save_mappings(mappings_db)
                        print(f"Stored new signature for {supplier_name}: {signature}")
                        break
                        
        if header_row_idx is not None and col_mapping is not None:
            if file_path.endswith('.csv') or file_path.endswith('.CSV'):
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    df = pd.read_csv(f, sep=sep, skiprows=header_row_idx)
            else:
                df = pd.read_excel(file_path, skiprows=header_row_idx)
                
            cols = df.columns.tolist()
            # Convert col_mapping keys from str to int because JSON loads keys as strings!
            cleaned_mapping = {}
            for key, col_idx in col_mapping.items():
                col_idx_int = int(col_idx)
                if col_idx_int < len(cols):
                    df.rename(columns={cols[col_idx_int]: key}, inplace=True)
                    cleaned_mapping[key] = col_idx_int
                    
            for _, row in df.iterrows():
                code = str(row.get('code', '')).strip()
                if not code or code.lower() == 'nan':
                    continue
                desc = str(row.get('desc', '')).strip()
                qty = parse_float(row.get('qty', 1))
                
                price_val = parse_float(row.get('net_price', row.get('price', 0.0)))
                disc_val = parse_float(row.get('disc', 0.0))
                net_price = price_val * (1 - disc_val/100) if 'net_price' not in cleaned_mapping else price_val
                
                vat = str(row.get('vat', '22')).strip()
                
                rows.append({
                    'Cod.': '',
                    'Cod. prod. forn.': code,
                    'Descrizione': desc,
                    'Q.tà': qty,
                    'Prezzo netto': round(net_price, 4),
                    'Iva': vat,
                    'Note': 'Generic Excel/CSV Import'
                })
    except Exception as e:
        print(f"Generic Excel/CSV parse failed: {e}")
    return rows

def parse_generic_pdf(file_path):
    print(f"Attempting generic PDF parsing: {file_path}")
    rows = []
    try:
        with pdfplumber.open(file_path) as pdf:
            # 1. Try table extraction first
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if len(table) < 2:
                        continue
                    code_idx = desc_idx = qty_idx = price_idx = vat_idx = disc_idx = None
                    for col_idx, col_name in enumerate(table[0]):
                        if not col_name:
                            continue
                        col_name_lower = str(col_name).lower()
                        if any(x in col_name_lower for x in ['codice', 'ean', 'art', 'barcode']) and code_idx is None:
                            code_idx = col_idx
                        elif any(x in col_name_lower for x in ['descrizione', 'descr', 'articolo', 'nome']) and desc_idx is None:
                            desc_idx = col_idx
                        elif any(x in col_name_lower for x in ['quant', 'q.ta', 'q.tà', 'qty', 'quantita']) and qty_idx is None:
                            qty_idx = col_idx
                        elif any(x in col_name_lower for x in ['prezzo', 'listino', 'netto', 'unitario']) and price_idx is None:
                            price_idx = col_idx
                        elif any(x in col_name_lower for x in ['sconto', 'disc', 'sc%']) and disc_idx is None:
                            disc_idx = col_idx
                        elif any(x in col_name_lower for x in ['iva', 'aliquota']) and vat_idx is None:
                            vat_idx = col_idx
                            
                    if code_idx is not None or desc_idx is not None:
                        for row in table[1:]:
                            if len(row) <= max(x for x in [code_idx, desc_idx, qty_idx, price_idx, vat_idx, disc_idx] if x is not None):
                                continue
                            code = str(row[code_idx]).strip() if code_idx is not None else ''
                            if not code or code.lower() in ['nan', 'codice', 'ean', 'cod']:
                                continue
                            desc = str(row[desc_idx]).strip() if desc_idx is not None else ''
                            qty = parse_float(row[qty_idx]) if qty_idx is not None else 1.0
                            price_val = parse_float(row[price_idx]) if price_idx is not None else 0.0
                            disc_val = parse_float(row[disc_idx]) if disc_idx is not None else 0.0
                            vat = str(row[vat_idx]).strip() if vat_idx is not None else '22'
                            
                            net_price = price_val * (1 - disc_val/100)
                            
                            rows.append({
                                'Cod.': '',
                                'Cod. prod. forn.': code,
                                'Descrizione': desc,
                                'Q.tà': qty,
                                'Prezzo netto': round(net_price, 4),
                                'Iva': vat,
                                'Note': 'Generic PDF Table Import'
                            })
            
            # 2. Heuristics fallback if no items from tables
            if not rows:
                for page in pdf.pages:
                    text = page.extract_text()
                    if not text:
                        continue
                    for line in text.split('\n'):
                        cleaned = line.strip()
                        parts = cleaned.split()
                        if len(parts) >= 4:
                            code = parts[0]
                            if re.match(r'^[A-Z0-9-]{4,15}$', code, re.IGNORECASE):
                                try:
                                    vat = parts[-1]
                                    amount = parse_float(parts[-2])
                                    price = parse_float(parts[-3])
                                    qty = parse_float(parts[-4])
                                    if qty > 0 and price > 0:
                                        rows.append({
                                            'Cod.': '',
                                            'Cod. prod. forn.': code,
                                            'Q.tà': qty,
                                            'Prezzo netto': price,
                                            'Iva': vat,
                                            'Note': 'Generic PDF Heuristics'
                                        })
                                except:
                                    pass
    except Exception as e:
        print(f"Generic PDF parsing failed: {e}")
    return rows

# --- ROUTER AND RUNNER ---

def identify_and_process_file(file_path):
    filename = os.path.basename(file_path).lower()
    
    # 1. Ferrari
    if 'ferrari' in filename:
        return parse_ferrari(file_path), "Ferrari"
        
    # 2. Viridex
    if 'viridex' in filename or '107682' in filename:
        return parse_viridex(file_path), "Viridex"
        
    # 3. Pilato
    if 'pilato' in filename:
        return parse_pilato(file_path), "Pilato"
        
    # 4. Monge
    if 'monge' in filename:
        return parse_monge(file_path), "Monge"
        
    # 5. Capaldo (handles Capaldo and Calaldo)
    if 'capaldo' in filename or 'calaldo' in filename:
        return parse_capaldo(file_path), "Capaldo"
        
    # 6. Blumen
    if 'blumen' in filename:
        return parse_blumen(file_path), "Blumen"
        
    # 7. Verisem
    if 'verisem' in filename:
        return parse_verisem(file_path), "Verisem"
        
    # 8. Giardini
    if 'giardini' in filename:
        return parse_giardini(file_path), "Giardini"
        
    # 9. Garver / Riccini
    if 'garver' in filename or 'riccini' in filename:
        return parse_garver(file_path), "Garver"
        
    # 10. Record
    if 'record' in filename:
        return parse_record(file_path), "Record"
        
    # 11. Ricci
    if 'ricci' in filename:
        return parse_ricci(file_path), "Ricci"
        
    # Content-based fallbacks
    if file_path.endswith('.pdf') or file_path.endswith('.PDF'):
        try:
            with pdfplumber.open(file_path) as pdf:
                sample_text = pdf.pages[0].extract_text() or ""
                sample_text_lower = sample_text.lower()
                if "monge" in sample_text_lower:
                    return parse_monge(file_path), "Monge"
                if "capaldo" in sample_text_lower:
                    return parse_capaldo(file_path), "Capaldo"
                if "blumen" in sample_text_lower:
                    return parse_blumen(file_path), "Blumen"
                if "verisem" in sample_text_lower:
                    return parse_verisem(file_path), "Verisem"
                if "giardini" in sample_text_lower:
                    return parse_giardini(file_path), "Giardini"
                if "irrigazionericcini" in sample_text_lower or "garver" in sample_text_lower:
                    return parse_garver(file_path), "Garver"
                if "rinaldo franco" in sample_text_lower or "recordit" in sample_text_lower:
                    return parse_record(file_path), "Record"
                if "ricci pietro" in sample_text_lower or "woodtech" in sample_text_lower:
                    return parse_ricci(file_path), "Ricci"
        except:
            pass
            
    # Content-based check for CSV/TXT
    if file_path.endswith('.csv') or file_path.endswith('.CSV') or file_path.endswith('.txt'):
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                first_line = f.readline()
                if 'Nr. Articolo' in first_line:
                    return parse_capaldo(file_path), "Capaldo"
        except:
            pass
            
    # --- GENERIC FALLBACKS FOR NEW/UNKNOWN SUPPLIERS ---
    if file_path.endswith('.pdf') or file_path.endswith('.PDF'):
        gen_rows = parse_generic_pdf(file_path)
        if gen_rows:
            return gen_rows, "Generico"
            
    if file_path.endswith('.xlsx') or file_path.endswith('.xls') or file_path.endswith('.csv') or file_path.endswith('.CSV') or file_path.endswith('.txt'):
        gen_rows = parse_generic_excel_csv(file_path)
        if gen_rows:
            return gen_rows, "Generico"
            
    print(f"Warning: Could not identify supplier template for file: {file_path}")
    return None, None

def generate_report(processed_files_dict):
    """
    Writes a summary report file in the Lavorati directory.
    """
    out_dir = get_output_dir()
    os.makedirs(out_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    report_path = os.path.join(out_dir, f"report_{timestamp}.txt")
    try:
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write("==================================================\n")
            f.write("           REPORT DI ELABORAZIONE DANEA           \n")
            f.write(f"Data/Ora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write("==================================================\n\n")
            
            if not processed_files_dict:
                f.write("Nessun file elaborato in questa sessione.\n")
            else:
                f.write(f"File elaborati con successo: {len(processed_files_dict)}\n\n")
                for original_file, info in processed_files_dict.items():
                    f.write(f"File: {os.path.basename(original_file)}\n")
                    f.write(f"  - Fornitore: {info['supplier']}\n")
                    f.write(f"  - Righe estratte: {info['rows_count']}\n")
                    f.write(f"  - Destinazione Excel: {os.path.basename(info['out_file'])}\n")
                    f.write(f"  - Stato: Elaborato ed eliminato per risparmiare spazio\n\n")
            
            f.write("Fine report.\n")
        print(f"Report summary generated at: {report_path}")
        return report_path
    except Exception as e:
        print(f"Failed to generate report: {e}")
        return None

def run_import(input_dir=None):
    if input_dir is None:
        input_dir = get_input_dir()
    print(f"Scanning directory: {input_dir}")
    # Supported extensions
    patterns = ["*.xlsx", "*.xls", "*.pdf", "*.PDF", "*.csv", "*.CSV", "*.txt"]
    files = []
    for pat in patterns:
        files.extend(glob.glob(os.path.join(input_dir, pat)))
        
    # Deduplicate file paths
    files = sorted(list(set(os.path.abspath(f) for f in files)))
        
    processed_files = {}
    for f in files:
        # Skip template itself, output files, and documentation/reports
        fname_lower = os.path.basename(f).lower()
        if ("modello importazione" in fname_lower or 
            "importazione_" in fname_lower or 
            "estrazione articoli" in fname_lower or
            "guida" in fname_lower or
            "report_" in fname_lower or
            "mappings" in fname_lower or
            "storico_prezzi" in fname_lower):
            continue
            
        rows, supplier = identify_and_process_file(f)
        if rows:
            print(f"Extracted {len(rows)} items from {supplier} file.")
            out_file = create_danea_import(rows, supplier)
            if out_file:
                # Delete original file directly to save space
                try:
                    os.remove(f)
                    print(f"Original file deleted: {f}")
                    status = "Eliminato"
                except Exception as e:
                    print(f"Error deleting original file {f}: {e}")
                    status = "Errore eliminazione"
                    
                processed_files[f] = {
                    "supplier": supplier,
                    "rows_count": len(rows),
                    "out_file": out_file,
                    "status": status
                }
                
    if processed_files:
        generate_report(processed_files)
    else:
        print("No files were processed.")
    return processed_files

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading

class DaneaImportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("danea-import-automation")
        self.root.geometry("700x500")
        self.root.minsize(600, 400)
        
        # Set icon
        try:
            icon_path = os.path.join(SCRIPT_DIR, "icon.png")
            if os.path.exists(icon_path):
                self.icon_img = tk.PhotoImage(file=icon_path)
                self.root.iconphoto(False, self.icon_img)
        except Exception as e:
            print(f"Failed to load icon: {e}")
            
        # Configure styles
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TButton', font=('Helvetica', 10), padding=6)
        self.style.configure('Main.TButton', font=('Helvetica', 10, 'bold'), foreground='white', background='#2a6f97')
        self.style.map('Main.TButton', background=[('active', '#014f86')])
        
        # Directory variable
        self.input_dir_var = tk.StringVar(value=get_input_dir())
        
        # Notebook for Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.tab_process = ttk.Frame(self.notebook)
        self.tab_settings = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab_process, text=" Elaborazione ")
        self.notebook.add(self.tab_settings, text=" Impostazioni ")
        
        self.create_widgets()
        self.create_settings_widgets()
        self.scan_files()
        
    def create_widgets(self):
        # Top Frame: Path Selection
        top_frame = ttk.LabelFrame(self.tab_process, text=" Cartella di Lavoro ", padding=10)
        top_frame.pack(fill="x", padx=15, pady=10)
        
        dir_entry = ttk.Entry(top_frame, textvariable=self.input_dir_var, font=('Helvetica', 10), state="readonly")
        dir_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        # Middle Frame: Files List & Log Console
        middle_frame = ttk.PanedWindow(self.tab_process, orient="horizontal")
        middle_frame.pack(fill="both", expand=True, padx=15, pady=5)
        
        # Left Panel: Files List
        files_frame = ttk.LabelFrame(middle_frame, text=" File da elaborare ", padding=5)
        self.files_listbox = tk.Listbox(files_frame, font=('Consolas', 9), selectmode="extended")
        self.files_listbox.pack(fill="both", expand=True)
        
        list_scroll = ttk.Scrollbar(self.files_listbox, orient="vertical", command=self.files_listbox.yview)
        list_scroll.pack(side="right", fill="y")
        self.files_listbox.config(yscrollcommand=list_scroll.set)
        
        middle_frame.add(files_frame, weight=1)
        
        # Right Panel: Console Log
        console_frame = ttk.LabelFrame(middle_frame, text=" Console Log ", padding=5)
        self.console_text = tk.Text(console_frame, wrap="word", state="disabled", font=('Consolas', 9), background="#f8f9fa", foreground="#212529")
        self.console_text.pack(fill="both", expand=True)
        
        console_scroll = ttk.Scrollbar(self.console_text, orient="vertical", command=self.console_text.yview)
        console_scroll.pack(side="right", fill="y")
        self.console_text.config(yscrollcommand=console_scroll.set)
        
        middle_frame.add(console_frame, weight=2)
        
        # Bottom Frame: Action Buttons
        btn_frame = ttk.Frame(self.tab_process, padding=10)
        btn_frame.pack(fill="x", padx=15, pady=5)
        
        self.refresh_btn = ttk.Button(btn_frame, text="Aggiorna Lista", command=self.scan_files)
        self.refresh_btn.pack(side="left", padx=5)
        
        self.coffee_btn = ttk.Button(btn_frame, text="Offrimi un caffè ☕", command=self.open_coffee)
        self.coffee_btn.pack(side="left", padx=5)
        
        self.process_btn = ttk.Button(btn_frame, text="Avvia Elaborazione", style='Main.TButton', command=self.start_processing)
        self.process_btn.pack(side="right", padx=5)
        
    def create_settings_widgets(self):
        settings_container = ttk.Frame(self.tab_settings, padding=20)
        settings_container.pack(fill="both", expand=True)
        
        # 1. Directories Section
        dir_frame = ttk.LabelFrame(settings_container, text=" Cartelle di Lavoro ", padding=10)
        dir_frame.pack(fill="x", pady=(0, 15))
        
        # Input Dir
        ttk.Label(dir_frame, text="Cartella Input:", font=('Helvetica', 9, 'bold')).grid(row=0, column=0, sticky="w", pady=5)
        self.input_dir_setting_var = tk.StringVar(value=get_input_dir())
        input_entry = ttk.Entry(dir_frame, textvariable=self.input_dir_setting_var, font=('Helvetica', 9))
        input_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=5)
        browse_input_btn = ttk.Button(dir_frame, text="Sfoglia...", command=lambda: self.browse_setting_dir(self.input_dir_setting_var))
        browse_input_btn.grid(row=0, column=2, pady=5)
        
        # Output Dir
        ttk.Label(dir_frame, text="Cartella Output:", font=('Helvetica', 9, 'bold')).grid(row=1, column=0, sticky="w", pady=5)
        self.output_dir_setting_var = tk.StringVar(value=get_output_dir())
        output_entry = ttk.Entry(dir_frame, textvariable=self.output_dir_setting_var, font=('Helvetica', 9))
        output_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=5)
        browse_output_btn = ttk.Button(dir_frame, text="Sfoglia...", command=lambda: self.browse_setting_dir(self.output_dir_setting_var))
        browse_output_btn.grid(row=1, column=2, pady=5)
        
        dir_frame.columnconfigure(1, weight=1)
        
        # 2. Output Filename Pattern Section
        filename_frame = ttk.LabelFrame(settings_container, text=" Formato Nome File in Uscita ", padding=10)
        filename_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(filename_frame, text="Pattern nome file:", font=('Helvetica', 9, 'bold')).pack(anchor="w", pady=2)
        self.pattern_setting_var = tk.StringVar(value=get_filename_pattern())
        pattern_entry = ttk.Entry(filename_frame, textvariable=self.pattern_setting_var, font=('Helvetica', 10))
        pattern_entry.pack(fill="x", pady=5)
        
        help_label = ttk.Label(filename_frame, text="Segnaposto supportati:\n- {supplier_name} : Nome del fornitore (es. Monge)\n- {date} : Data corrente (es. 18-07-2026)", font=('Helvetica', 8), foreground="#555555")
        help_label.pack(anchor="w", pady=5)
        
        # 3. Operations Section (Reset)
        reset_frame = ttk.LabelFrame(settings_container, text=" Gestione Dati e Ripristino ", padding=10)
        reset_frame.pack(fill="x", pady=(0, 15))
        
        reset_btn = ttk.Button(reset_frame, text="Azzera Dati (Storico Prezzi e Mappature)", command=self.reset_script_data)
        reset_btn.pack(side="left", pady=5)
        
        # Save Buttons Frame
        save_btn_frame = ttk.Frame(settings_container)
        save_btn_frame.pack(fill="x", side="bottom", pady=10)
        
        save_btn = ttk.Button(save_btn_frame, text="Salva Impostazioni", style='Main.TButton', command=self.save_settings)
        save_btn.pack(side="right")
        
    def open_coffee(self):
        import webbrowser
        webbrowser.open("https://www.paypal.me/gates538")
        
    def browse_setting_dir(self, var):
        selected = filedialog.askdirectory(initialdir=var.get())
        if selected:
            var.set(os.path.abspath(selected))
            
    def save_settings(self):
        CONFIG["settings"]["input_dir"] = self.input_dir_setting_var.get()
        CONFIG["settings"]["output_dir"] = self.output_dir_setting_var.get()
        CONFIG["settings"]["filename_pattern"] = self.pattern_setting_var.get()
        save_mappings(CONFIG)
        
        # Update active variables
        self.input_dir_var.set(self.input_dir_setting_var.get())
        self.scan_files()
        
        messagebox.showinfo("Impostazioni", "Impostazioni salvate con successo!")
        
    def reset_script_data(self):
        confirm = messagebox.askyesno("Conferma Reset", "Sei sicuro di voler azzerare tutti i dati memorizzati?\n\nQuesto eliminerà:\n1. Lo storico dei prezzi di tutti i fornitori\n2. Le mappature dei nuovi fornitori apprese automaticamente\n\nQuesta operazione non può essere annullata.")
        if confirm:
            try:
                # 1. Reset mappings signatures
                CONFIG["signatures"] = {}
                save_mappings(CONFIG)
                
                # 2. Reset SQLite price history DB
                if os.path.exists(DB_PATH):
                    # Close connection by letting it scope out, remove DB file, re-init
                    conn = sqlite3.connect(DB_PATH)
                    conn.close()
                    os.remove(DB_PATH)
                    init_db()
                    
                messagebox.showinfo("Reset Completato", "Tutti i dati storici e di apprendimento sono stati azzerati!")
            except Exception as e:
                messagebox.showerror("Errore Reset", f"Si è verificato un errore durante il reset: {e}")
                
    def scan_files(self):
        self.files_listbox.delete(0, 'end')
        dir_to_scan = self.input_dir_var.get()
        if not os.path.exists(dir_to_scan):
            return
            
        patterns = ["*.xlsx", "*.xls", "*.pdf", "*.PDF", "*.csv", "*.CSV", "*.txt"]
        files = []
        for pat in patterns:
            files.extend(glob.glob(os.path.join(dir_to_scan, pat)))
            
        files = sorted(list(set(os.path.abspath(f) for f in files)))
        
        for f in files:
            fname = os.path.basename(f)
            fname_lower = fname.lower()
            if ("modello importazione" in fname_lower or 
                "importazione_" in fname_lower or 
                "estrazione articoli" in fname_lower or
                "guida" in fname_lower or
                "report_" in fname_lower or
                "mappings" in fname_lower or
                "storico_prezzi" in fname_lower):
                continue
            self.files_listbox.insert('end', f"  {fname}")
            
    def log(self, text):
        self.console_text.config(state="normal")
        self.console_text.insert('end', text + "\n")
        self.console_text.see('end')
        self.console_text.config(state="disabled")
        
    def start_processing(self):
        self.process_btn.config(state="disabled")
        self.refresh_btn.config(state="disabled")
        self.console_text.config(state="normal")
        self.console_text.delete('1.0', 'end')
        self.console_text.config(state="disabled")
        
        t = threading.Thread(target=self.run_import_thread)
        t.start()
        
    def run_import_thread(self):
        class RedirectText:
            def __init__(self, log_func):
                self.log_func = log_func
            def write(self, string):
                if string.strip():
                    self.log_func(string.strip())
            def flush(self):
                pass
                
        old_stdout = sys.stdout
        sys.stdout = RedirectText(self.log)
        
        try:
            input_dir = self.input_dir_var.get()
            processed = run_import(input_dir)
            self.root.after(0, lambda: self.finish_processing(processed))
        except Exception as e:
            self.log(f"Error during execution: {e}")
            self.root.after(0, lambda: self.finish_processing({}))
        finally:
            sys.stdout = old_stdout
            
    def finish_processing(self, processed):
        self.process_btn.config(state="normal")
        self.refresh_btn.config(state="normal")
        self.scan_files()
        
        if processed:
            messagebox.showinfo("Successo", f"Elaborati con successo {len(processed)} file!\nControlla la cartella degli output ed i report.")
        else:
            messagebox.showwarning("Nessun File", "Nessun file elaborato. Controlla che ci siano file da elaborare nella cartella selezionata.")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Danea Import Script")
    parser.add_argument('--dir', type=str, default=None, help="Directory to scan for supplier invoice/DDT files")
    parser.add_argument('--cli', action='store_true', help="Run in CLI mode instead of GUI")
    args = parser.parse_args()
    
    if args.cli:
        run_import(args.dir)
    else:
        root = tk.Tk()
        app = DaneaImportApp(root)
        root.mainloop()
