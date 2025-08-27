import streamlit as st
import pandas as pd
from pathlib import Path
import os
from io import BytesIO

# ========= Page Setup =========
st.set_page_config(page_title="Excel & Steel Manager — Bushra Abu Hani", layout="wide", page_icon="📊")

# ========= CSS Theme =========
st.markdown("""
<style>
  body {
      background: linear-gradient(160deg, #E6E6FA 0%, #D8BFD8 100%);
  }
  .title {
      font-size: 2.6rem; 
      font-weight: bold;
      background: linear-gradient(90deg,#6a0dad,#b57edc,#dda0dd);
      background-size: 200% 100%;
      animation: gradientMove 6s ease infinite;
      -webkit-background-clip: text;
      color: transparent;
  }
  @keyframes gradientMove {
      0%{background-position:0% 50%}
      50%{background-position:100% 50%}
      100%{background-position:0% 50%}
  }
  .card {
      border-radius: 20px;
      border: 1px solid rgba(0,0,0,.08);
      padding: 1rem;
      margin-bottom: 1rem;
      background: linear-gradient(180deg, rgba(230,230,250,.95), rgba(220,220,250,.85));
      box-shadow: 0 6px 16px rgba(0,0,0,0.1);
  }
  button {
      background-color: #D8BFD8 !important;  
      color: #fff !important;
      font-weight: bold;
  }
  button:hover {
      background-color: #C8A2C8 !important;
      color: #fff !important;
  }
</style>
""", unsafe_allow_html=True)

# ========= Header =========
st.markdown('<h1 class="title">📊 Excel & Steel Manager — Bushra Abu Hani</h1>', unsafe_allow_html=True)

# ========= Data Folder =========
DATA_FOLDER = Path("All Data - Bushra Abu Hani")
DATA_FOLDER.mkdir(parents=True, exist_ok=True)

# ========= Employees File =========
EMPLOYEES_FILE = DATA_FOLDER / "Employees.xlsx"
if not EMPLOYEES_FILE.exists():
    employees_data = [
        [198,"Donald","OConnell","DOCONNEL","650.507.9833","21-Jun-07","SH_CLERK",2600,"-",124,50],
        [199,"Douglas","Grant","DGRANT","650.507.9844","13-Jan-08","SH_CLERK",2600,"-",124,50],
        [200,"Jennifer","Whalen","JWHALEN","515.123.4444","17-Sep-03","AD_ASST",4400,"-",101,10],
        [201,"Michael","Hartstein","MHARTSTE","515.123.5555","17-Feb-04","MK_MAN",13000,"-",100,20],
        [202,"Pat","Fay","PFAY","603.123.6666","17-Aug-05","MK_REP",6000,"-",201,20],
        [203,"Susan","Mavris","SMAVRIS","515.123.7777","7-Jun-02","HR_REP",6500,"-",101,40],
        [204,"Hermann","Baer","HBAER","515.123.8888","7-Jun-02","PR_REP",10000,"-",101,70],
        [205,"Shelley","Higgins","SHIGGINS","515.123.8080","7-Jun-02","AC_MGR",12008,"-",101,110],
        [206,"William","Gietz","WGIETZ","515.123.8181","7-Jun-02","AC_ACCOUNT",8300,"-",205,110],        [140,"Joshua","Patel","JPATEL","650.121.1834","6-Apr-06","ST_CLERK",2500,"-",123,50]
    ]
    employees_df = pd.DataFrame(employees_data, columns=[
        "EMPLOYEE_ID","FIRST_NAME","LAST_NAME","EMAIL","PHONE_NUMBER",
        "HIRE_DATE","JOB_ID","SALARY","COMMISSION_PCT","MANAGER_ID","DEPARTMENT_ID"
    ])
    employees_df.to_excel(EMPLOYEES_FILE, index=False)

# ========= Helper Functions =========
def list_excel_files(folder: Path):
    return sorted([f for f in os.listdir(folder) if f.lower().endswith(".xlsx")])

def safe_read_excel(path: Path) -> pd.DataFrame:
    if not path.exists() or path.stat().st_size == 0:
        return pd.DataFrame()
    try:
        return pd.read_excel(path)
    except Exception as e:
        st.warning(f"⚠️ Failed to read file: {e}")
        return pd.DataFrame()

def safe_write_excel(path: Path, df: pd.DataFrame):
    try:
        df.to_excel(path, index=False)
        st.toast("Saved successfully ✅", icon="✅")
    except Exception as e:
        st.error(f"❌ Failed to save: {e}")

def create_excel(path: Path):
    if path.exists():
        st.info("ℹ️ File already exists.")
    else:
        safe_write_excel(path, pd.DataFrame())
        st.balloons()

# ========= Sidebar =========
with st.sidebar:
    st.header("⚙️ Settings")
    mode = st.radio("Choose Mode:", ["Excel Manager 📂", "Steel Weight Calculator ⚖️"], key="mode_radio")

    st.markdown("---")
    st.header("📂 File Management")
    files = list_excel_files(DATA_FOLDER)
    current_file = st.selectbox("Select Excel file", options=["— None —"] + files, index=0)

    # Create New File
    st.subheader("➕ Create New File")
    new_name = st.text_input("File name (e.g., data.xlsx)", value="")
    if st.button("Create File", use_container_width=True):
        if not new_name.lower().endswith(".xlsx"):
            st.warning("⚠️ File name must end with .xlsx")
        else:
            create_excel(DATA_FOLDER / new_name)
            st.rerun()

    # Special File: Warehouse Number One
    st.subheader("🏭 Warehouse Number One")
    w_name = "Warehouse Number One.xlsx"
    if st.button("📦 Create Warehouse File", use_container_width=True):
        path = DATA_FOLDER / w_name
        cols = ["Product", "Quantity", "Amount", "Weight", "Product Serial Number", "Product Supplier"]
        if path.exists():
            st.warning("⚠️ File already exists.")
        else:
            df = pd.DataFrame(columns=cols)
            df.to_excel(path, index=False)
            st.success("✅ 'Warehouse Number One' created successfully!")
            st.balloons()
            st.rerun()

    # Delete File
    st.subheader("🗑️ Delete File")
    if current_file != "— None —":
        confirm_del = st.checkbox("Confirm deletion")
        if st.button("Delete Selected File", disabled=not confirm_del, use_container_width=True):
            try:
                os.remove(DATA_FOLDER / current_file)
                st.toast("File deleted 🗑️", icon="🗑️")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Failed to delete file: {e}")

# ========= Mode: Excel Manager =========
if mode == "Excel Manager 📂":
    if current_file != "— None —":
        path = DATA_FOLDER / current_file
        st.markdown(f"#### 🗃️ Current File: {current_file}")
        df = safe_read_excel(path)
        df = df[[c for c in df.columns if c.lower() not in ["name", "row name"]]]

        # --- Column & Row Management ---
        with st.container():
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.subheader("🛠️ Columns & Rows Management")
            c1, c2, c3 = st.columns([1,1,1])

            # Add Columns
            with c1:
                st.write("➕ Add Columns")
                ncols = st.number_input("Number of Columns", min_value=1, max_value=20, value=1, step=1)
                col_names = []
                for i in range(ncols):
                    col_name = st.text_input(f"Column {i+1} Name", value=f"Column_{i+1}", key=f"col_name_{i}")
                    if col_name.strip().lower() in ["name", "row name"]:
                        st.warning("⚠️ Column name 'Name' or 'Row Name' not allowed.")
                        col_name = f"Column_{len(df.columns)+i+1}"
                    col_names.append(col_name)
                if st.button("Add Columns", use_container_width=True):
                    for name in col_names:
                        new_name = name
                        k = 2
                        while new_name in df.columns:
                            new_name = f"{name}_{k}"
                            k += 1
                        df[new_name] = None
                    safe_write_excel(path, df)
                    st.balloons()
                    st.rerun()

            # Delete Columns
            with c2:
                st.write("➖ Delete Columns")
                cols_to_drop = st.multiselect("Select columns", options=list(df.columns))
                if st.button("Delete Selected Columns", use_container_width=True, disabled=len(cols_to_drop)==0):
                    df = df.drop(columns=cols_to_drop, errors="ignore")
                    safe_write_excel(path, df)
                    st.toast("Columns deleted 🗑️", icon="🗑️")
                    st.rerun()

            # Add Rows
            with c3:
                st.write("➕ Add Rows")
                nrows = st.number_input("Number of Rows", min_value=1, max_value=50, value=1, step=1)
                if st.button("Add Rows", use_container_width=True):
                    blanks = pd.DataFrame([{c: None for c in df.columns} for _ in range(nrows)])
                    df = pd.concat([df, blanks], ignore_index=True)
                    safe_write_excel(path, df)
                    st.balloons()
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        # --- Editable Table ---
        st.markdown("### ✏️ Editable Table")
        original_dtypes = df.dtypes.to_dict()
        df_editable = df.astype(str)

        edited_df = st.data_editor(
            df_editable,
            hide_index=True,
            use_container_width=True,
            num_rows="dynamic",
            key="editor",
            column_config={col: st.column_config.TextColumn(label=col) for col in df_editable.columns}
        )

        def restore_dtypes(original_dtypes, edited_df):
            for col, dtype in original_dtypes.items():
                if col in edited_df.columns and pd.api.types.is_numeric_dtype(dtype):
                    edited_df[col] = pd.to_numeric(edited_df[col], errors='coerce')
            return edited_df

        # --- Action Bar ---
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("⚡ Actions")
        s1, s2, s3 = st.columns([1,1,1])

        with s1:
            if st.button("💾 Save Table", use_container_width=True):
                edited_df_to_save = restore_dtypes(original_dtypes, edited_df)
                safe_write_excel(path, edited_df_to_save)
                st.balloons()

        with s2:
            st.write("🗑️ Delete Rows")
            if not edited_df.empty:
                rows_to_delete = st.multiselect(
                    "Select Rows",
                    options=list(range(len(edited_df))),
                    format_func=lambda x: f"Row {x+1}"
                )
                if st.button("Delete Selected Rows", disabled=len(rows_to_delete)==0, use_container_width=True):
                    new_df = edited_df.drop(index=rows_to_delete).reset_index(drop=True)
                    safe_write_excel(path, new_df)
                    st.toast("Rows deleted 🗑️", icon="🗑️")
                    st.balloons()
                    st.rerun()

        with s3:
            if st.button("⟲ Reload Table", use_container_width=True):
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    else:
        st.info("ℹ️ Select a file from the sidebar or create a new one.")

# ========= Mode: Steel Weight Calculator =========
if mode == "Steel Weight Calculator ⚖️":
    st.subheader("⚖️ Steel Weight Calculator")

    DENSITIES = {
        "Mild Steel (7850)": 7850,
        "Stainless Steel (8000)": 8000,
        "Aluminum (2700)": 2700,
    }

    if "rows" not in st.session_state:
        st.session_state.rows = []

    with st.expander("📘 Formula Explanation", expanded=False):
        st.write("Weight = Length × Width × Thickness × Density (kg/m³). Dimensions in meters, result in kg.")

    with st.form("input_form", clear_on_submit=False):
        col1, col2, col3, col4 = st.columns(4)
        with col1: length = st.number_input("Length (m)", min_value=0.0, step=0.01)
        with col2: width = st.number_input("Width (m)", min_value=0.0, step=0.01)
        with col3: thickness = st.number_input("Thickness (m)", min_value=0.0, step=0.001, format="%.4f")
        with col4: steel_type = st.selectbox("Material", list(DENSITIES.keys()))
        submitted = st.form_submit_button("Calculate")

    if submitted:
        if length <= 0 or width <= 0 or thickness <= 0:
            st.error("⚠️ All dimensions must be > 0.")
        else:
            density = DENSITIES[steel_type]
            weight = length * width * thickness * density
            result = {
                "Length (m)": length,
                "Width (m)": width,
                "Thickness (m)": thickness,
                "Material": steel_type,
                "Density (kg/m³)": density,
                "Weight (kg)": round(weight, 4),
            }
            st.session_state.rows.append(result)
            st.success("✅ Weight calculated & added to table!")

    st.subheader("📊 Results")
    if st.session_state.rows:
        df = pd.DataFrame(st.session_state.rows)
        st.table(df)

        def to_excel_bytes(dataframe: pd.DataFrame) -> bytes:
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                dataframe.to_excel(writer, index=False, sheet_name="weights")
            return output.getvalue()

        excel_bytes = to_excel_bytes(df)
        st.download_button(
            "⬇️ Download Excel",
            data=excel_bytes,
            file_name="steel_weights.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if st.button("🗑️ Clear Table"):
            st.session_state.rows = []
            st.info("Table cleared.")
    else:
        st.info("ℹ️ No results yet. Enter values above and click Calculate.")

# ========= Footer =========
st.markdown("""
<div style="margin-top: 1rem; padding: .8rem 1rem; border-radius: 20px;
            background: linear-gradient(90deg, #D8BFD8 0%, #E6E6FA 100%);
            color: #fff; font-weight: bold; text-align: center;">
✨
</div>
""", unsafe_allow_html=True)
