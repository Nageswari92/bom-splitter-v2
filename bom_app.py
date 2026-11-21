import io
import os
import re
import zipfile
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill

# Streamlit Page Configuration
st.set_page_config(page_title="BAVA TECH - BOM Splitter & Tree App", layout="wide", page_icon="⚙️")

# Custom CSS Styling
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 8px; font-weight: bold; }
    h1 { color: #1E3A8A; }
    </style>
    """, unsafe_allow_html=True)

# Sidebar Navigation
with st.sidebar:
    st.image("https://img.icons8.com/fluent/96/000000/tree-structure.png", width=80)
    st.title("BAVA TECH")
    st.caption("BOM Management Suite")
    st.markdown("---")
    st.subheader("💡 How to use:")
    st.write("1. Upload your Indented BOM File.")
    st.write("2. Upload your Template File.")
    st.write("3. Click Generate to build tree & split BOM files in ZIP.")
    st.markdown("---")
    st.caption("v2.2.0 | Integrated Engine")

st.title("📊 Indented BOM Splitter & Formatter")
st.write("Upload your BOM and Template files below to process and download the split Excel files along with the formatted BoM Tree inside a hierarchical ZIP package.")

# --- File Upload Section ---
col1, col2 = st.columns(2)
with col1:
    uploaded_bom = st.file_uploader(
        "1. Upload Indented BOM Excel File (.xls / .xlsx)", type=["xls", "xlsx"]
    )
with col2:
    uploaded_template = st.file_uploader(
        "2. Upload Empty Template Excel File (.xlsx)", type=["xlsx"]
    )

if st.button("Generate BOM Files & Tree Package", type="primary"):
    if uploaded_bom is not None and uploaded_template is not None:
        try:
            with st.spinner("Processing files and building tree directory... Please wait..."):
                bom_base_name = os.path.splitext(uploaded_bom.name)[0]
                bom_folder_name = bom_base_name.strip()

                # 1. Read Excel files
                df = pd.read_excel(uploaded_bom, sheet_name=None, dtype=str)

                df_bom = df["BOM"] if "BOM" in df else list(df.values())[0]
                df_bom.columns = df_bom.columns.str.strip()
                df_bom["Level"] = df_bom["Level"].astype(int)

                df_mfg = df["MFG"] if "MFG" in df else None

                # --- Convert MFG data into a Lookup Dictionary ---
                mfg_lookup = {}
                if df_mfg is not None:
                    df_mfg.columns = df_mfg.columns.str.strip()
                    for idx, row in df_mfg.iterrows():
                        part_no = str(row["Part Number"]).strip()
                        if part_no not in mfg_lookup:
                            mfg_lookup[part_no] = {
                                "mfg_name": (
                                    str(row["Manufacturer Name"]).strip()
                                    if pd.notna(row["Manufacturer Name"])
                                    else ""
                                ),
                                "mfg_pn": (
                                    str(row["Manufacturer Part Number"]).strip()
                                    if pd.notna(row["Manufacturer Part Number"])
                                    else ""
                                ),
                            }

                # Assembly tracking structures
                assemblies = {}
                current_parent_key_at_lvl = {0: None}
                current_path_at_lvl = {0: ""} 

                # 2. Hierarchy Tracking Logic
                for idx, row in df_bom.iterrows():
                    lvl = row["Level"]
                    part_no = row["Part Number"]

                    parent_key = current_parent_key_at_lvl.get(lvl - 1)
                    
                    if parent_key is not None and parent_key in assemblies:
                        assemblies[parent_key]["children"].append(row)

                    unique_key = f"{part_no}_row_{idx}"
                    current_parent_key_at_lvl[lvl] = unique_key
                    
                    if lvl == 0:
                        current_path_at_lvl[lvl] = ""
                    else:
                        upstream_path = current_path_at_lvl.get(lvl - 1, "")
                        upstream_parent_part = df_bom.iloc[idx - 1]["Part Number"] if idx > 0 else ""
                        actual_parent = assemblies[parent_key]["part_number"] if parent_key in assemblies else upstream_parent_part
                        
                        if upstream_path == "":
                            current_path_at_lvl[lvl] = str(actual_parent)
                        else:
                            current_path_at_lvl[lvl] = f"{upstream_path}/{actual_parent}"

                    if unique_key not in assemblies:
                        assemblies[unique_key] = {
                            "part_number": part_no, 
                            "info": row, 
                            "children": [],
                            "folder_path": current_path_at_lvl[lvl]
                        }

                # Memory buffer to create the ZIP file
                zip_buffer = io.BytesIO()

                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:

                    # =========================================================
                    # 3. CUSTOM BOM TREE CREATION & HIGHLIGHTING
                    # =========================================================
                    tree_data = []
                    total_rows = len(df_bom)

                    for index, row in df_bom.iterrows():
                        level = row["Level"]
                        part_number = str(row["Part Number"]).strip() 
                        
                        tree_row = {
                            "Customer": bom_folder_name if index == 0 else None,
                            "Top Level": None,
                            "Level 2": None,
                            "Level 3": None,
                            "Level 4": None,
                            "Level 5": None
                        }
                        
                        has_child = False
                        if level in [0, 1, 2, 3, 4]:
                            for j in range(index + 1, total_rows):
                                if df_bom.iloc[j]["Level"] == level + 1:
                                    has_child = True
                                    break
                                elif df_bom.iloc[j]["Level"] <= level:
                                    break

                        if level == 0:
                            tree_row["Top Level"] = part_number
                            tree_data.append(tree_row)
                        elif level == 1 and has_child:
                            tree_row["Level 2"] = part_number
                            tree_data.append(tree_row)
                        elif level == 2 and has_child:
                            tree_row["Level 3"] = part_number
                            tree_data.append(tree_row)
                        elif level == 3 and has_child:
                            tree_row["Level 4"] = part_number
                            tree_data.append(tree_row)
                        elif level == 4 and has_child:
                            tree_row["Level 5"] = part_number
                            tree_data.append(tree_row)

                    df_tree = pd.DataFrame(tree_data)
                    df_tree = df_tree[["Customer", "Top Level", "Level 2", "Level 3", "Level 4", "Level 5"]]
                    
                    bom_tree_output = io.BytesIO()
                    with pd.ExcelWriter(bom_tree_output, engine='openpyxl') as writer:
                        df_tree.to_excel(writer, index=False)
                        
                        worksheet = writer.sheets['Sheet1']
                        orange_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
                        tree_prefixes = ("0042", "0040", "0043", "0243", "0015", "0300")
                        
                        for row in worksheet.iter_rows(min_row=2, min_col=1, max_col=6):
                            for cell in row:
                                if cell.value is not None:
                                    val_str = str(cell.value).strip()
                                    if val_str.startswith(tree_prefixes):
                                        cell.fill = orange_fill

                    bom_tree_output.seek(0)
                    bom_tree_filename = f"BoM_Tree_{bom_folder_name}.xlsx"
                    bom_tree_zip_path = f"{bom_folder_name}/{bom_tree_filename}"
                    zip_file.writestr(bom_tree_zip_path, bom_tree_output.getvalue())

                    # =========================================================
                    # 4. INDIVIDUAL SPLIT BOM GENERATION LOGIC
                    # =========================================================
                    for unique_key, data in assemblies.items():
                        child_rows = data["children"]

                        if not child_rows:
                            continue

                        parent_part = str(data["part_number"]).strip()
                        
                        excluded_prefixes = ("0243", "0040", "1290", "0020", "0042", "0015", "0300", "0043")
                        if parent_part.startswith(excluded_prefixes):
                            continue

                        parent_info = data["info"]
                        parent_desc = str(parent_info["Description"]).strip()
                        inherited_folder_path = data["folder_path"]

                        parent_rev = (
                            str(parent_info["Revision"]).strip()
                            if pd.notna(parent_info["Revision"])
                            else ""
                        )
                        if parent_rev.lower() == "nan":
                            parent_rev = ""
                        elif len(parent_rev) == 1 and parent_rev.isdigit():
                            parent_rev = parent_rev.zfill(2)

                        uploaded_template.seek(0)
                        wb = load_workbook(uploaded_template)

                        if "VSE_BOM" in wb.sheetnames:
                            ws = wb["VSE_BOM"]
                        else:
                            ws = wb.active

                        # Capture Row 9 base formulas BEFORE any deletion
                        row9_formulas = {}
                        for col_idx in range(12, ws.max_column + 1):  # Column L (12) to BE (57)
                            val = ws.cell(row=9, column=col_idx).value
                            if isinstance(val, str) and val.startswith("="):
                                row9_formulas[col_idx] = val

                        # Header fields
                        ws["D2"] = str(parent_part)
                        ws["D3"] = str(parent_rev)
                        ws["D4"] = parent_desc
                        ws["D5"] = "00"
                        ws["D6"] = "APP01"
                        ws["D7"] = ""

                        ws["H2"] = "Std. Due (Days)"
                        ws["H3"] = "BU"

                        header_font = Font(name="Verdana", size=10, bold=True)
                        ws["H2"].font = header_font
                        ws["H3"].font = header_font
                        ws["H2"].alignment = Alignment(horizontal="left", vertical="center")
                        ws["H3"].alignment = Alignment(horizontal="left", vertical="center")

                        for row_idx in range(2, 8):
                            ws[f"D{row_idx}"].alignment = Alignment(horizontal="left", vertical="center")
                            ws[f"D{row_idx}"].font = Font(name="Verdana", size=10, bold=False)

                        # Row 9: Parent Component Row
                        row_num = 9
                        ws.cell(row=row_num, column=1, value=int(parent_info["Level"])).alignment = Alignment(horizontal="center")

                        seq_val = (
                            str(parent_info["Find No"]).split(".")[0].zfill(4)
                            if pd.notna(parent_info["Find No"])
                            else "0000"
                        )
                        ws.cell(row=row_num, column=2, value=seq_val).alignment = Alignment(horizontal="center")
                        ws.cell(row=row_num, column=3, value="")

                        # Store Qty as NUMERIC (float) so formulas don't break
                        qty_cell_9 = ws.cell(row=row_num, column=4)
                        if pd.notna(parent_info["Quantity"]):
                            try:
                                qty_cell_9.value = float(parent_info["Quantity"])
                            except ValueError:
                                qty_cell_9.value = 1.0
                        else:
                            qty_cell_9.value = 1.0
                        qty_cell_9.number_format = "0.000"
                        qty_cell_9.alignment = Alignment(horizontal="right")

                        ws.cell(row=row_num, column=5, value=str(parent_part))
                        ws.cell(row=row_num, column=6, value=str(parent_rev)).alignment = Alignment(horizontal="center")
                        ws.cell(row=row_num, column=7, value=parent_desc)
                        ws.cell(row=row_num, column=8, value="")
                        ws.cell(row=row_num, column=9, value=str(parent_info["UOM"])).alignment = Alignment(horizontal="center")

                        parent_part_clean = str(parent_part).strip()
                        if parent_part_clean in mfg_lookup:
                            ws.cell(row=row_num, column=10, value=mfg_lookup[parent_part_clean]["mfg_name"])
                            ws.cell(row=row_num, column=11, value=mfg_lookup[parent_part_clean]["mfg_pn"])

                        for col in range(1, 12):
                            ws.cell(row=row_num, column=col).font = Font(name="Verdana", size=10)

                        row_num += 1

                        # Row 10+: Child Items
                        for child in child_rows:
                            ws.cell(row=row_num, column=1, value=int(child["Level"])).alignment = Alignment(horizontal="center")

                            c_seq = (
                                str(child["Find No"]).split(".")[0].zfill(4)
                                if pd.notna(child["Find No"])
                                else "0000"
                            )
                            ws.cell(row=row_num, column=2, value=c_seq).alignment = Alignment(horizontal="center")
                            ws.cell(row=row_num, column=3, value="")

                            # Store Child Qty as NUMERIC (float)
                            qty_cell_child = ws.cell(row=row_num, column=4)
                            if pd.notna(child["Quantity"]):
                                try:
                                    qty_cell_child.value = float(child["Quantity"])
                                except ValueError:
                                    qty_cell_child.value = 0.0
                            else:
                                qty_cell_child.value = 0.0
                            qty_cell_child.number_format = "0.000"
                            qty_cell_child.alignment = Alignment(horizontal="right")

                            child_part = str(child["Part Number"])
                            ws.cell(row=row_num, column=5, value=child_part)

                            c_rev = (
                                str(child["Revision"]).strip()
                                if pd.notna(child["Revision"])
                                else ""
                            )
                            if c_rev.lower() == "nan":
                                c_rev = ""
                            elif len(c_rev) == 1 and c_rev.isdigit():
                                c_rev = c_rev.zfill(2)

                            ws.cell(row=row_num, column=6, value=str(c_rev)).alignment = Alignment(horizontal="center")
                            ws.cell(row=row_num, column=7, value=child["Description"])
                            ws.cell(row=row_num, column=8, value="")
                            ws.cell(row=row_num, column=9, value=str(child["UOM"])).alignment = Alignment(horizontal="center")

                            child_part_clean = child_part.strip()
                            if child_part_clean in mfg_lookup:
                                ws.cell(row=row_num, column=10, value=mfg_lookup[child_part_clean]["mfg_name"])
                                ws.cell(row=row_num, column=11, value=mfg_lookup[child_part_clean]["mfg_pn"])

                            for col in range(1, 12):
                                cell = ws.cell(row=row_num, column=col)
                                cell.font = Font(name="Verdana", size=10)

                            row_num += 1

                        # --- DELETE UNUSED TEMPLATE ROWS (From row_num up to row 1399) ---
                        if row_num <= 1399:
                            delete_count = 1399 - row_num + 1
                            ws.delete_rows(idx=row_num, amount=delete_count)

                        # --- RESTORE & UPDATE ALL FORMULAS CLEANLY ---
                        last_data_row = row_num - 1  # Data ends here (e.g., Row 35)
                        # After deleting up to 1399:
                        # Row 1400 (Black separator line) becomes -> row_num
                        # Row 1401 (Totals / Leadtime Days) becomes -> row_num + 1
                        # Row 1402 (Leadtime Weeks) becomes -> row_num + 2
                        total_row = row_num + 1
                        weeks_row = row_num + 2

                        # 1. Restore Row-by-Row Formulas (Row 9 to last_data_row, Columns M to BE)
                        for r in range(9, last_data_row + 1):
                            for col_idx, base_f in row9_formulas.items():
                                # Replace row 9 references (e.g., B9, D9, N9, O9, S9, W9...) with current row 'r'
                                # Keeps $V$7, $Z$7, $A:$AD untouched
                                updated_f = re.sub(
                                    r'(\$?[A-Z]{1,3})9\b',
                                    rf'\g<1>{r}',
                                    base_f
                                )
                                ws.cell(row=r, column=col_idx).value = updated_f

                        # 2. Update Total Row Formulas (Row 1401 -> total_row)
                        ws.cell(row=total_row, column=17).value = f"=MAX(Q9:Q{last_data_row})"  # Col Q: Leadtime Days
                        sum_cols = ["S", "T", "W", "X", "AA", "AB", "AE", "AF", "AI", "AJ", 
                                    "AM", "AN", "AQ", "AR", "AU", "AV", "AY", "AZ", "BC", "BD"]
                        for col_let in sum_cols:
                            ws[f"{col_let}{total_row}"].value = f"=SUM({col_let}9:{col_let}{last_data_row})"

                        # 3. Update Leadtime Weeks Formula (Row 1402 -> weeks_row)
                        ws.cell(row=weeks_row, column=17).value = f"=Q{total_row}/7"

                        excel_data = io.BytesIO()
                        wb.save(excel_data)
                        excel_data.seek(0)

                        output_filename = f"{parent_part}_REV_{parent_rev}_VSE_00_BOM.xlsx"
                        
                        if inherited_folder_path == "":
                            full_zip_path = f"{bom_folder_name}/{output_filename}"
                        else:
                            full_zip_path = f"{bom_folder_name}/{inherited_folder_path}/{output_filename}"
                            
                        zip_file.writestr(full_zip_path, excel_data.read())

                zip_buffer.seek(0)

                st.success("🎉 Process Completed Successfully! All BOM files and highlighted BoM_Tree created.")
                st.download_button(
                    label="📥 Download Hierarchical Tree BOM Package (ZIP)",
                    data=zip_buffer,
                    file_name=f"{bom_folder_name}.zip",
                    mime="application/zip",
                    type="primary"
                )
        except Exception as e:
            st.error(f"❌ Processing Error: {e}")
    else:
        st.warning("⚠️ Please upload BOTH the Indented BOM Excel and Template Excel files before proceeding!")
