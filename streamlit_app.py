import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

# ==========================================
# 1. CÁC HÀM HỖ TRỢ (HELPER)
# ==========================================

def clean_avolta_number(num_str):
    """
    Hàm chuyển đổi chuỗi sang số thực (float).
    """
    if not num_str: return 0.0
    s = str(num_str).strip()
    s = re.sub(r'[^\d.,-]', '', s)
    
    if ',' in s: 
        s = s.replace('.', '').replace(',', '.')
    else:
        if '.' in s:
            parts = s.split('.')
            if len(parts) > 1 and len(parts[-1]) == 3:
                 s = s.replace('.', '')
            else:
                 pass
    try:
        return float(s)
    except ValueError:
        return 0.0

def clean_product_name(name):
    """
    Hàm làm sạch tên sản phẩm theo yêu cầu đặc biệt.
    Ví dụ: "Xa Veg Lettuce, Iceberg Kg" -> "Xa Lach, Iceberg"
           "Hanh Tay - Veg Onion, Peeled Kg (BK)" -> "Hanh Tay - (BK)"
    """
    if not name: return ""
    
    # 1. Sửa các lỗi đặc thù (Hard replacement)
    # Thay "Xa Veg" thành "Xa Lach" (do PDF thường bị mất chữ Lach)
    name = name.replace("Xa Veg", "Xa Lach")
    
    # 2. Danh sách các từ cần XÓA (Tiếng Anh/Đơn vị thừa)
    remove_words = [
        "Veg", "Herb", "Fruit", "Flower", "Kg", "kg", "KG",
        "Lettuce", "Onion", "Tomato", "Peeled", "Fresh", "Sliced", "Slice",
        "Beansprouts", "Carrots", "Chillies", "Ginger", "Saw Leaves",
        "Chive", "Coriander", "Knotweed", "Lemongrass", "Mint", 
        "Morning Glory", "Basil", "Lemon Leaves", "Bok Choy", "Cabbage", 
        "Celery", "Cucumber", "Shallot", "Spring"
    ]
    
    # Xóa từng từ trong danh sách (không phân biệt hoa thường)
    for word in remove_words:
        # Dùng regex để thay thế word đứng riêng lẻ hoặc dính dấu câu
        pattern = re.compile(r'\b' + re.escape(word) + r'\b', re.IGNORECASE)
        name = pattern.sub('', name)

    # 3. Làm sạch dấu câu và khoảng trắng thừa
    # Thay thế nhiều dấu phẩy liên tiếp thành 1
    name = re.sub(r',+', ',', name)
    # Thay thế dấu gạch ngang thừa
    name = re.sub(r'-+', '-', name)
    # Xóa khoảng trắng thừa
    name = re.sub(r'\s+', ' ', name).strip()
    # Xóa dấu phẩy/gạch ngang ở đầu/cuối câu
    name = name.strip(', -')
    
    # Sửa lỗi thẩm mỹ cuối cùng: ", ," -> ","
    name = name.replace(" ,", ",")
    
    return name

# ==========================================
# 2. HÀM BÓC TÁCH 4PS
# ==========================================
def parse_4ps_po(pdf):
    st.write("  > Nhận diện: Mẫu PO của 4PS. Đang xử lý...")
    items_list = []

    page1 = pdf.pages[0]
    full_text_page1 = page1.extract_text() 

    order_num_match = re.search(r"Order Number\s*:\s*(\d+)", full_text_page1)
    delivery_date_match = re.search(r"Request Del\. Time\s*:\s*(\d{2}/\d{2}/\d{4})", full_text_page1)
    buyer_name_match = re.search(r"Buyer Name\s*:\s*([^\n]+)", full_text_page1)
    
    order_number = order_num_match.group(1).strip() if order_num_match else None
    delivery_date = delivery_date_match.group(1).strip() if delivery_date_match else None
    buyer_name = buyer_name_match.group(1).strip() if buyer_name_match else None

    for i, page in enumerate(pdf.pages):
        tables = page.extract_tables({"vertical_strategy": "lines", "horizontal_strategy": "lines"})
        if not tables: tables = page.extract_tables()
        if not tables: continue 
            
        item_table = tables[-1] 
        for row in item_table:
            if not row or len(row) < 6: continue 
            product_code = row[1]
            if product_code == "Product Code": continue
            if (row[2] or "").strip() == "Total": continue
            if not product_code or product_code.strip() == "": continue
            
            quantity_str = row[4].replace(',', '') if row[4] else '0'
            price_str = row[5].replace(',', '') if row[5] else '0'
            
            # Áp dụng làm sạch tên cho cả 4PS (nếu cần đồng bộ)
            # row[2] là Item Name
            cleaned_name = row[2].replace('\n', ' ')

            standard_item = {
                "Order_Number": order_number,    
                "Buyer_Name": buyer_name,      
                "Delivery_Date": delivery_date,
                "Item_Code": product_code,
                "Vendor No.": cleaned_name, 
                "Quantity": quantity_str,
                "Price": price_str
            }
            items_list.append(standard_item)
    
    return items_list

# ==========================================
# 3. HÀM BÓC TÁCH AVOLTA (REGEX + CLEAN NAME)
# ==========================================
def parse_avolta_po(pdf):
    st.write("  > Nhận diện: Mẫu PO Avolta (SĐT 0903613502). Đang xử lý...")
    items_list = []

    page1 = pdf.pages[0]
    page1_text = page1.extract_text() or ""
    
    order_num_match = re.search(r"PO No\.[\s\S]*?(\S+)", page1_text)
    order_number = order_num_match.group(1).strip() if order_num_match else "Unknown"
    
    delivery_date_match = re.search(r"Order Date\s*(\d{2}/\d{2}/\d{4})", page1_text)
    delivery_date = delivery_date_match.group(1).strip() if delivery_date_match else None
    
    buyer_name = "Unknown"
    if "Delivery Address" in page1_text:
        parts = page1_text.split("Delivery Address")
        if len(parts) > 1:
            lines = parts[1].strip().split('\n')
            buyer_name = " ".join(lines[:2]).strip()

    # Regex Scan
    line_start_pattern = re.compile(r"^(\d+)\s+(.+)")

    for page in pdf.pages:
        text = page.extract_text()
        if not text: continue
        
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if "PO No" in line or "Page" in line or "Total" in line or "Item No" in line:
                continue

            match = line_start_pattern.match(line)
            if match:
                potential_numbers = [
                    n for n in re.findall(r'[\d.,]+', line) 
                    if any(char.isdigit() for char in n)
                ]
                
                if len(potential_numbers) >= 3:
                    item_code = potential_numbers[0]
                    qty_raw = potential_numbers[1]
                    
                    if len(potential_numbers) >= 4:
                        price_raw = potential_numbers[-2]
                    else:
                        price_raw = potential_numbers[-1]
                    
                    try:
                        start_index = line.find(item_code) + len(item_code)
                        end_index = line.find(qty_raw, start_index)
                        if end_index != -1:
                            raw_item_name = line[start_index:end_index].strip()
                        else:
                            raw_item_name = match.group(2)
                    except:
                        raw_item_name = match.group(2)

                    # --- ÁP DỤNG LÀM SẠCH TÊN ---
                    final_name = clean_product_name(raw_item_name)

                    qty_final = clean_avolta_number(qty_raw)
                    price_final = clean_avolta_number(price_raw)
                    
                    if 0 < price_final < 1000:
                        price_final *= 1000

                    items_list.append({
                        "Order_Number": order_number,    
                        "Buyer_Name": buyer_name,      
                        "Delivery_Date": delivery_date,
                        "Item_Code": item_code,
                        "Vendor No.": final_name, # <-- Tên đã làm sạch
                        "Quantity": qty_final,
                        "Price": price_final
                    })

    return items_list

# ==========================================
# 4. HÀM TẠO EXCEL
# ==========================================
def create_hybrid_excel(standard_df, unrecognized_files_list):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        if not standard_df.empty:
            standard_df.to_excel(writer, sheet_name="TongHop_DonHang", index=False)
        else:
            pd.DataFrame(["Không có dữ liệu chuẩn hóa."]).to_excel(writer, sheet_name="TongHop_DonHang", index=False, header=False)
        
        if unrecognized_files_list:
            st.write("--- Đang xử lý các file khác (Dump Text) ---")
            for uploaded_file in unrecognized_files_list:
                safe_sheet_name = re.sub(r'[\\/*?:"<>|\[\]\s]', '_', uploaded_file.name.split('.')[0])[:30]
                try:
                    uploaded_file.seek(0)
                    with pdfplumber.open(uploaded_file) as pdf:
                        all_lines = []
                        for page in pdf.pages:
                            text = page.extract_text(layout=True, keep_blank_chars=True)
                            if text: all_lines.extend(text.split('\n'))
                            all_lines.append("--- END PAGE ---")
                    
                    if all_lines:
                        pd.DataFrame(all_lines).to_excel(writer, sheet_name=safe_sheet_name, index=False, header=False)
                except Exception as e:
                    st.error(f"Lỗi dump file {uploaded_file.name}: {e}")

    return output.getvalue()

# ==========================================
# 5. GIAO DIỆN CHÍNH
# ==========================================
st.set_page_config(page_title="Công cụ tổng hợp PO", layout="wide")
st.title("🚀 Công cụ trích xuất dữ liệu PO sang Excel")
st.markdown("Hỗ trợ: 4PS & Avolta (SĐT 0903613502). Đã tích hợp làm sạch tên sản phẩm.")

uploaded_files = st.file_uploader("Tải file PDF lên:", type="pdf", accept_multiple_files=True)

if uploaded_files and st.button("Xử lý tất cả file"):
    all_standardized_data = []
    unrecognized_files = []
    
    progress_bar = st.progress(0)
    
    with st.expander("Chi tiết quá trình xử lý:", expanded=True):
        for i, uploaded_file in enumerate(uploaded_files):
            file_name = uploaded_file.name
            st.write(f"--- Đang mở: **{file_name}** ---")
            
            try:
                uploaded_file.seek(0)
                with pdfplumber.open(uploaded_file) as pdf:
                    if not pdf.pages:
                        st.error("File lỗi hoặc không có trang.")
                        continue
                    
                    page1_text = pdf.pages[0].extract_text() or ""
                    
                    items = []
                    is_recognized = False
                    customer_name = ""

                    if "4PS CORPORATION" in page1_text or "CÔNG TY TNHH MTV KITCHEN 4PS" in page1_text:
                        customer_name = "4PS"
                        items = parse_4ps_po(pdf)
                        is_recognized = True
                    elif "0903613502" in page1_text:
                        customer_name = "Avolta"
                        items = parse_avolta_po(pdf)
                        is_recognized = True
                    
                    if is_recognized:
                        for item in items:
                            item['Customer'] = customer_name
                            item['File_Name'] = file_name
                            all_standardized_data.append(item)
                        st.success(f"  > Đã xử lý xong ({customer_name}). Lấy được {len(items)} dòng.")
                    else:
                        st.info("  > Không nhận diện được mẫu. Chuyển sang chế độ dump text.")
                        unrecognized_files.append(uploaded_file)

            except Exception as e:
                st.error(f"Lỗi khi xử lý file {file_name}: {e}")
            
            progress_bar.progress((i + 1) / len(uploaded_files))

    df_standard = pd.DataFrame(all_standardized_data)
    
    if not df_standard.empty:
        try:
            if '4PS' in df_standard['Customer'].values:
                 df_standard['Quantity'] = pd.to_numeric(df_standard['Quantity'], errors='coerce').fillna(0)
                 df_standard['Price'] = pd.to_numeric(df_standard['Price'], errors='coerce').fillna(0)
        except: pass
        
        cols = ['Customer', 'Order_Number', 'Buyer_Name', 'Delivery_Date', 'Item_Code', 'Vendor No.', 'Quantity', 'Price', 'File_Name']
        final_cols = [c for c in cols if c in df_standard.columns]
        df_standard = df_standard[final_cols]
        
        st.success(f"🎉 Hoàn tất! Tổng hợp được {len(df_standard)} dòng dữ liệu chuẩn hóa.")
        st.dataframe(df_standard)
    else:
        st.warning("Chưa tìm thấy dữ liệu chuẩn hóa nào.")

    excel_data = create_hybrid_excel(df_standard, unrecognized_files)
    
    st.download_button(
        label="📥 Tải file Excel kết quả",
        data=excel_data,
        file_name="TongHop_PO_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
