import streamlit as st
import pandas as pd
import re
import io
import datetime
import pytz

st.set_page_config(page_title="Medical Data Converter", page_icon="📋", layout="centered")

def parse_date(d_str):
    try: return pd.to_datetime(d_str, format='%d/%m/%Y')
    except: pass
    try: return pd.to_datetime(d_str, format='%m/%y')
    except: pass
    try: return pd.to_datetime(d_str, format='%m/%Y')
    except: pass
    try: return pd.to_datetime(d_str, format='%d/%m/%y')
    except: pass
    try: return pd.to_datetime(d_str, dayfirst=True)
    except: return pd.NaT

# -------------------------------------------------------------
# 1. ALL INVENTORY / ZERO STOCK FILE PROCESS CHEYYANULLA FUNCTION
# -------------------------------------------------------------
def process_inv_file(uploaded_file):
    try:
        df = pd.read_csv(uploaded_file, on_bad_lines='skip')
    except:
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file)

    header_idx = None
    for i, row in df.iterrows():
        if row.astype(str).str.contains('iname', case=False, na=False).any():
            header_idx = i
            break

    if header_idx is not None:
        df.columns = df.iloc[header_idx]
        df = df.iloc[header_idx+1:].reset_index(drop=True)

    df.columns = [str(c).strip().lower() for c in df.columns]

    inv_data = []
    # 🛠️ കമ്പനിയും റാക്കും ലാസ്റ്റ് സപ്ലയറും സേവ് ചെയ്യുന്ന ഡിക്ഷണറി 
    xls_lookup = {} 

    for _, row in df.iterrows():
        iname = str(row.get('iname', '')).strip()
        if not iname or iname.lower() == 'nan':
            continue

        imanf = str(row.get('imanf', '')).strip()
        rack = str(row.get('rack', '')).strip()
        sup = str(row.get('bigsup', '')).strip() 

        mfg_clean = imanf.upper() if imanf and imanf.lower() != 'nan' else "MISC."
        rack_clean = rack if rack and rack.lower() != 'nan' else ""
        last_sup_clean = sup if sup and sup.lower() != 'nan' else ""
        
        # ഡിക്ഷണറിയിലേക്ക് ഐറ്റത്തിന്റെ പേര് വെച്ച് സേവ് ചെയ്യുന്നു
        xls_lookup[iname.upper()] = {
            "mfg": mfg_clean,
            "rack": rack_clean,
            "last_sup": last_sup_clean
        }

        try: 
            qty_val = float(str(row.get('iqty', '0')).strip())
            qty = int(qty_val) if qty_val.is_integer() else qty_val
        except: 
            qty = 0

        inv_data.append({
            "Item Name": iname,
            "Manufacturer": mfg_clean,
            "Supplier": "", # സീറോ സ്റ്റോക്കിൽ ബാച്ച് സപ്ലയർ ഇല്ല
            "Last Supplier": last_sup_clean,
            "Rack ID": rack_clean,
            "Packing": "",
            "Batch": "",
            "Expiry Date": pd.NaT,
            "MRP": 0.0,
            "Quantity": qty,  
            "Invoice Date": pd.NaT,
            "Invoice Number": ""
        })
        
    return pd.DataFrame(inv_data), xls_lookup


# -------------------------------------------------------------
# 2. EXPIRY TXT FILE PROCESS CHEYYANULLA FUNCTION
# -------------------------------------------------------------
def process_txt_file(uploaded_file, xls_lookup):
    data_rows = []
    current_supplier = "UNKNOWN SUPPLIER"  
    start_parsing = False
    
    date_pattern = r'\b(?:0?[1-9]|[12][0-9]|3[01])/(?:0?[1-9]|1[012])/\d{2,4}\b|\b(?:0?[1-9]|1[012])/\d{2,4}\b|(?:(?:0?[1-9]|[12][0-9]|3[01])/(?:0?[1-9]|1[012])/\d{2,4}|(?:0?[1-9]|1[012])/\d{2,4})(?=\s)'
    
    known_mfgs = [
        "MICRO GEN", "DR. REDDY", "DR.REDDY", "BOEHRING", "CHETHANA", "LA RENON", "GLENMARK", "BLUECROS", 
        "MACLEODS", "SYSTOPIC", "BLUECOSS", "DA RENON", "RELIANCE", "ISIS HEA", "CU-CARD", 
        "CU CARD", "ALEMBIC", "CURATIO", "AKESISS", "LEEFORD", "MANKIND", "LIVIDUS", "PANACEA", 
        "WALLACE", "RENAUXE", "ARISTO", "ZEYYER", "AVELOR", "REDDYS", "KISWAR", "AKESIS", 
        "BIOCON", "SANOFI", "ABBOTT", "GERMAN", "LUPIN", "ALKEM", "EYSYS", "MISC.", "PIRCA", 
        "INTAS", "AUREL", "CIPLA", "MICRO", "LLOYD", "ZYDUS", "ELITE", "IPCA", "ICON", "H&H", 
        "SUN", "ZEY", "USV"
    ]

    stringio = io.StringIO(uploaded_file.getvalue().decode("utf-8"))
    raw_lines = stringio.readlines()
    cleaned_lines = [line.rstrip('\r\n') for line in raw_lines if line.strip()]

    merged_lines = []
    i = 0
    while i < len(cleaned_lines):
        line = cleaned_lines[i]
        
        if '\\' not in line and "======" not in line and "----" not in line and "EXPIRED" not in line.upper() and "DATE" not in line.upper():
            if i + 1 < len(cleaned_lines) and '\\' in cleaned_lines[i+1] and "Item Name" not in cleaned_lines[i+1]:
                line = line.strip() + " " + cleaned_lines[i+1].strip()
                i += 1  
                
        if '\\' in line and not list(re.finditer(date_pattern, line)) and "Item Name" not in line:
            if i + 1 < len(cleaned_lines):
                line = line.strip() + " " + cleaned_lines[i+1].strip()
                i += 1  

        merged_lines.append(line)
        i += 1

    for line_raw in merged_lines:
        line_stripped = line_raw.strip()
        
        if "======" in line_stripped:
            start_parsing = True
            continue
        if not start_parsing and "----" in line_stripped:
            start_parsing = True
            continue
        if not start_parsing or not line_stripped:
            continue
            
        if '\\' not in line_stripped and '/' not in line_stripped and '-' not in line_stripped and not any(char.isdigit() for char in line_stripped[:12]):
            if "EXPIRED ITEMS" not in line_stripped.upper() and "ITEM NAME" not in line_stripped.upper() and "DATE :" not in line_stripped.upper():
                current_supplier = line_stripped
                continue
            
        try:
            if '\\' in line_raw:
                slash_pos = line_raw.find('\\')
                before_slash = line_raw[:slash_pos].strip()
                after_slash = line_raw[slash_pos + 1:].strip()
                
                if "Item Name" in before_slash or "Manf" in after_slash:
                    continue
            else:
                continue
                
            if '-' in before_slash:
                item_name, packing = before_slash.rsplit('-', 1)
                item_name = item_name.strip()
                packing = packing.strip()
            else:
                item_name = before_slash
                packing = ""
                
            all_date_matches = list(re.finditer(date_pattern, after_slash))
            if not all_date_matches:
                continue
                
            expiry_date_str = all_date_matches[0].group(0)  
            expiry_idx = after_slash.find(expiry_date_str)
            left_part = after_slash[:expiry_idx].strip()   
            right_part = after_slash[expiry_idx + len(expiry_date_str):].strip() 
            
            mfg = ""
            batch = ""
            xls_rack = ""
            xls_last_sup = ""
            
            # 🛠️ XLS ഫയലിലെ ഡേറ്റ വെച്ച് 100% കൃത്യമായി കമ്പനിയും ബാച്ചും എടുക്കുന്നു
            item_key = item_name.upper()
            if item_key in xls_lookup:
                mfg = xls_lookup[item_key]["mfg"]
                xls_rack = xls_lookup[item_key]["rack"]
                xls_last_sup = xls_lookup[item_key]["last_sup"] # Last Supplier From XLS
                
                # XLS-ൽ നിന്നുള്ള കമ്പനിയുടെ പേര് വെച്ച് യഥാർത്ഥ ബാച്ച് നമ്പർ മാത്രം മുറിച്ചെടുക്കുന്നു
                if mfg and mfg != "MISC.":
                    pattern = re.compile('^' + re.escape(mfg) + r'\s*', re.IGNORECASE)
                    if pattern.search(left_part):
                        batch = pattern.sub('', left_part).strip()

            # XLS-ൽ ഐറ്റം ഇല്ലെങ്കിൽ പഴയ രീതിയിൽ കണ്ടുപിടിക്കുന്നു (Fallback)
            if not batch and not mfg:
                for k_mfg in known_mfgs:
                    if left_part.upper().startswith(k_mfg.upper()):
                        mfg = k_mfg
                        batch = left_part[len(k_mfg):].strip()
                        break
                
                if not mfg:
                    mfg_batch_tokens = re.split(r'\s{2,}', left_part)
                    if len(mfg_batch_tokens) >= 2:
                        mfg = mfg_batch_tokens[0].strip()
                        batch = mfg_batch_tokens[1].strip()
                    else:
                        combined = mfg_batch_tokens[0]
                        words = combined.split()
                        if len(words) > 1 and any(char.isdigit() for char in words[-1]):
                            mfg = " ".join(words[:-1]).strip()
                            batch = words[-1].strip()
                        else:
                            match = re.match(r'^([a-zA-Z\s\-\.\*]+?)([0-9].*)$', combined)
                            if match:
                                mfg = match.group(1).strip()
                                batch = match.group(2).strip()
                            else:
                                mfg = combined
                                batch = ""

            right_tokens = right_part.split()
            quantity_str = right_tokens[0] if len(right_tokens) > 0 else "0"
            mrp_str = right_tokens[1] if len(right_tokens) > 1 else "0.0"
            
            invoice = ""
            invoice_date_str = ""
            rack_id = xls_rack # XLS-ൽ നിന്നുള്ള റാക്ക് ഐഡി ആണ് ആദ്യം എടുക്കുക
            
            if len(right_tokens) > 2:
                invoice_section = " ".join(right_tokens[2:])
                inv_date_matches = list(re.finditer(date_pattern, invoice_section))
                if inv_date_matches:
                    invoice_date_str = inv_date_matches[0].group(0)
                    idx = invoice_section.find(invoice_date_str)
                    invoice_part = invoice_section[:idx].strip()
                    rack_part = invoice_section[idx + len(invoice_date_str):].strip()
                    invoice = invoice_part.strip('- ').strip()
                    if not rack_id: 
                        rack_id = rack_part.strip('- ').strip()
                else:
                    inv_parts = [p.strip() for p in invoice_section.split('-') if p.strip()]
                    if len(inv_parts) >= 2:
                        invoice = inv_parts[0]
                        if not rack_id: rack_id = inv_parts[1]
                        if invoice == '*': invoice = ""
                    elif len(inv_parts) == 1:
                        val = inv_parts[0]
                        if any(char.isdigit() for char in val) and len(val) <= 3: 
                            if not rack_id: rack_id = val
                        else: 
                            invoice = val if val != '*' else ""
            
            expiry_date = parse_date(expiry_date_str)
            invoice_date = parse_date(invoice_date_str) if invoice_date_str else pd.NaT
                
            try: quantity = int(quantity_str)
            except: quantity = 0
            try: mrp = float(mrp_str)
            except: mrp = 0.0
                
            data_rows.append({
                "Item Name": item_name,
                "Manufacturer": mfg.upper() if mfg else "MISC.",
                "Supplier": current_supplier,       # 🛠️ TXT ഫയലിലെ യഥാർത്ഥ ബാച്ച് സപ്ലയർ
                "Last Supplier": xls_last_sup,      # 🛠️ XLS ഫയലിൽ നിന്നുള്ള ലാസ്റ്റ് സപ്ലയർ
                "Rack ID": rack_id if rack_id else "",
                "Packing": packing,
                "Batch": batch if batch else "BN",
                "Expiry Date": expiry_date,
                "MRP": mrp,
                "Quantity": quantity,
                "Invoice Date": invoice_date,
                "Invoice Number": invoice if invoice else ""
            })
        except Exception:
            pass
            
    return pd.DataFrame(data_rows)


# -------------------------------------------------------------
# 3. STREAMLIT UI & MERGING LOGIC
# -------------------------------------------------------------
st.title("📋 Medical Data Converter")
st.write("Upload your Expiry `.TXT` file and All Inventory `.XLS` file to get a combined perfect Excel sheet.")

col1, col2 = st.columns(2)
with col1:
    txt_file = st.file_uploader("1. Upload Expiry .TXT File", type=["txt", "TXT"])
with col2:
    inv_file = st.file_uploader("2. Upload All Inventory .XLS/CSV", type=["xls", "xlsx", "csv"])

if txt_file is not None or inv_file is not None:
    df_list = []
    xls_lookup = {}
    
    # 🛠️ ആദ്യം XLS ഫയൽ പ്രോസസ്സ് ചെയ്ത് ഡിക്ഷണറി ഉണ്ടാക്കുന്നു
    if inv_file is not None:
        df_inv, xls_lookup = process_inv_file(inv_file)
        if not df_inv.empty:
            df_list.append(df_inv)
            
    # 🛠️ XLS ഡേറ്റ വെച്ച് TXT ഫയൽ പെർഫെക്റ്റ് ആയി പ്രോസസ്സ് ചെയ്യുന്നു
    if txt_file is not None:
        df_txt = process_txt_file(txt_file, xls_lookup)
        if not df_txt.empty:
            df_list.append(df_txt)

    if df_list:
        final_df = pd.concat(df_list, ignore_index=True)
        
        # 🛠️ പുതിയ "Last Supplier" കോളം കൂടി ഉൾപ്പെടുത്തിയിരിക്കുന്നു
        columns_order = [
            "Item Name", "Manufacturer", "Supplier", "Last Supplier", "Rack ID", "Packing", 
            "Batch", "Expiry Date", "MRP", "Quantity", "Invoice Date", "Invoice Number"
        ]
        final_df = final_df[columns_order]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name="Merged Data")
            worksheet = writer.sheets["Merged Data"]
            
            for col in worksheet.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value is not None:
                        if isinstance(cell.value, (pd.Timestamp, datetime.datetime, datetime.date)):
                            cell.number_format = 'yyyy-mm-dd'
                            cell_len = 10
                        else:
                            cell_len = len(str(cell.value))
                        max_len = max(max_len, cell_len)
                worksheet.column_dimensions[col_letter].width = max(max_len + 5, 12)
                
        processed_data = output.getvalue()
        
        st.success(f"🎉 Files processed successfully! Total {len(final_df)} items combined.")
        
        ist = pytz.timezone('Asia/Kolkata')
        current_time = datetime.datetime.now(ist).strftime("%d-%m-%Y %I-%M-%p")
        dynamic_filename = f"{current_time} - final medical data.xlsx"
        
        st.download_button(
            label="📥 DOWNLOAD MERGED EXCEL FILE",
            data=processed_data,
            file_name=dynamic_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Could not parse any valid rows. Please check the file formats.")
