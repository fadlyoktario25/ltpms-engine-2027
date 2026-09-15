from collections import Counter
import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st
import xlrd

st.set_page_config(
    page_title="Portal Teknik LTPMS & STMPS", page_icon="⚙️", layout="wide"
)

st.title("⚙️ Portal Utama Maintenance (LTPMS & STMPS Generator)")
st.markdown(
    "Pusat otomatisasi dokumen perencanaan pemeliharaan jangka panjang"
    " (**LTPMS**) dan breakdown bulanan/mingguan (**STMPS**)."
)

# Sidebar Pengaturan Tahun & Bulan Target
st.sidebar.header("🔧 Pengaturan Target Dokumen")
target_year = st.sidebar.number_input(
    "Target Tahun", min_value=2025, max_value=2040, value=2027, step=1
)
selected_month = st.sidebar.selectbox(
    "Pilih Bulan untuk STMPS",
    options=list(range(1, 13)),
    format_func=lambda x: [
        "Januari",
        "Februari",
        "Maret",
        "April",
        "Mei",
        "Juni",
        "Juli",
        "Agustus",
        "September",
        "Oktober",
        "November",
        "Desember",
    ][x - 1],
    index=9,  # Default Oktober
)

y_n = target_year - 1
y_n1 = target_year - 2
y_n2 = target_year - 3
y_n3 = target_year - 4


def parse_history_ps(uploaded_file, year):
  ps_data = {}
  if not uploaded_file:
    return ps_data
  content = uploaded_file.read()
  uploaded_file.seek(0)

  if uploaded_file.name.endswith(".xlsx"):
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=False)
    for sname in wb.sheetnames:
      ws = wb[sname]
      for r in range(15, ws.max_row + 1):
        tag = str(ws.cell(r, 2).value or "").strip()
        name = str(ws.cell(r, 3).value or "").strip()
        if tag.endswith(".0"):
          tag = tag[:-2]
        clean_name = " ".join(name.upper().split())
        if not clean_name and not tag:
          continue

        months = []
        for c in range(4, 40):
          cell = ws.cell(r, c)
          if cell.fill and cell.fill.fill_type == "darkHorizontal":
            m = ((c - 4) // 3) + 1
            if m not in months:
              months.append(m)
        key = tag if tag else clean_name
        if months:
          ps_data[key] = sorted(months)
  else:
    wb = xlrd.open_workbook(file_contents=content, formatting_info=True)
    for sname in wb.sheet_names():
      if "sheet" in sname.lower() and sname.lower() != "1":
        continue
      sh = wb.sheet_by_name(sname)
      for r in range(14, sh.nrows):
        tag = str(sh.cell_value(r, 1)).strip() if sh.ncols > 1 else ""
        name = str(sh.cell_value(r, 2)).strip() if sh.ncols > 2 else ""
        if tag.endswith(".0"):
          tag = tag[:-2]
        clean_name = " ".join(name.upper().split())
        if not clean_name and not tag:
          continue

        months = []
        for c in range(3, min(39, sh.ncols)):
          xf = wb.xf_list[sh.cell_xf_index(r, c)]
          pat = xf.background.fill_pattern
          col = xf.background.pattern_colour_index
          is_ps = (
              col in (23, 24, 13)
              or pat == 5
              or (pat == 1 and col not in (8, 9, 64, 65))
          )
          if is_ps:
            m = ((c - 3) // 3) + 1
            if m not in months:
              months.append(m)
        key = tag if tag else clean_name
        if months:
          ps_data[key] = sorted(months)
  return ps_data


def process_ltpms(f_n3, f_n2, f_n1, f_n):
  ps_n3 = parse_history_ps(f_n3, y_n3) if f_n3 else {}
  ps_n2 = parse_history_ps(f_n2, y_n2) if f_n2 else {}
  ps_n1 = parse_history_ps(f_n1, y_n1) if f_n1 else {}

  content_master = f_n.read()
  f_n.seek(0)
  wb_master = openpyxl.load_workbook(
      io.BytesIO(content_master), data_only=False
  )

  if "1" in wb_master.sheetnames:
    wb_master["1"]["C7"].value = f":  {target_year}"

  ps_fill = PatternFill(
      fill_type="darkHorizontal",
      start_color="00000000",
      end_color="00000000",
  )
  pc_fill = PatternFill(
      fill_type="solid", start_color="00000000", end_color="00000000"
  )
  blank_fill = PatternFill(fill_type=None)

  scheduled_items = []

  for sname in wb_master.sheetnames:
    ws = wb_master[sname]
    for r in range(15, ws.max_row + 1):
      if r > ws.max_row - 8:
        row_txt = " ".join(
            [str(ws.cell(r, c).value or "") for c in range(1, 5)]
        ).upper()
        if "CHECK" in row_txt or "SERVICE" in row_txt or "NOTE" in row_txt:
          continue

      tag = str(ws.cell(r, 2).value or "").strip()
      name = str(ws.cell(r, 3).value or "").strip()
      rem = str(ws.cell(r, 40).value or "").strip()
      if tag.endswith(".0"):
        tag = tag[:-2]
      clean_name = " ".join(name.upper().split())
      if not clean_name and not tag:
        continue

      key = tag if tag else clean_name
      rem_upper = rem.upper()

      row_has_pc = False
      p_n = []
      for c in range(4, 40):
        cell = ws.cell(r, c)
        m = ((c - 4) // 3) + 1
        if (
            cell.fill
            and cell.fill.fill_type == "solid"
            and getattr(cell.fill.start_color, "index", None) in (0, 8)
        ):
          row_has_pc = True
        elif cell.fill and cell.fill.fill_type == "darkHorizontal":
          if m not in p_n:
            p_n.append(m)

      m_ps = re.search(r"PS\s*:\s*1\s*X\s*(\d+)\s*BLN", rem_upper)
      interval_ps = int(m_ps.group(1)) if m_ps else 12

      p_n3 = ps_n3.get(key, [])
      p_n2 = ps_n2.get(key, [])
      p_n1 = ps_n1.get(key, [])

      target_ps_months = []
      status_sched = ""

      if interval_ps == 12:
        target_ps_months = p_n if p_n else p_n1
        status_sched = "Rutin Tahunan (12 BLN)"
      elif interval_ps == 24:
        if p_n3 and p_n1:
          target_ps_months = p_n1 if p_n1 else p_n3
          status_sched = (
              f"Siklus Ganjil Asli ({y_n3} -> {y_n1} -> {target_year})"
          )
        elif p_n2:
          target_ps_months = []
          status_sched = f"Siklus Genap (Base {y_n2} -> Skip {target_year})"
        elif p_n1 or p_n3:
          target_ps_months = p_n1 if p_n1 else p_n3
          status_sched = f"Due Siklus 24 BLN dari {y_n1}"
      elif interval_ps == 36:
        if p_n2:
          target_ps_months = p_n2
          status_sched = f"Due Siklus 36 BLN dari {y_n2}"
        else:
          target_ps_months = []
          status_sched = f"Belum Jatuh Tempo ({interval_ps} BLN)"
      elif interval_ps == 48:
        if p_n3:
          target_ps_months = p_n3
          status_sched = f"Due Siklus 48 BLN dari {y_n3}"
        else:
          target_ps_months = []
          status_sched = f"Belum Jatuh Tempo ({interval_ps} BLN)"
      else:
        target_ps_months = p_n

      for m in p_n:
        if m not in target_ps_months:
          c_base = 4 + (m - 1) * 3
          fill_to_apply = pc_fill if row_has_pc else blank_fill
          for sub in range(3):
            ws.cell(r, c_base + sub).fill = fill_to_apply

      for m in target_ps_months:
        c_base = 4 + (m - 1) * 3
        for sub in range(3):
          ws.cell(r, c_base + sub).fill = ps_fill

      if target_ps_months:
        scheduled_items.append({
            "Sheet": sname,
            "Nama Mesin": name,
            "Interval": f"1x{interval_ps} BLN",
            "Keterangan": status_sched,
            "Bulan Servis": str(target_ps_months),
        })

  output_stream = io.BytesIO()
  wb_master.save(output_stream)
  output_stream.seek(0)
  return output_stream, scheduled_items


def generate_stmps(uploaded_ltpms, target_m, plant_name):
  content = uploaded_ltpms.read()
  uploaded_ltpms.seek(0)
  wb = openpyxl.load_workbook(io.BytesIO(content), data_only=False)

  month_names = [
      "JANUARI",
      "FEBRUARI",
      "MARET",
      "APRIL",
      "MEI",
      "JUNI",
      "JULI",
      "AGUSTUS",
      "SEPTEMBER",
      "OKTOBER",
      "NOVEMBER",
      "DESEMBER",
  ]
  m_name = month_names[target_m - 1]

  stmps_list = []
  c_base = 4 + (target_m - 1) * 3

  for sname in wb.sheetnames:
    ws = wb[sname]
    for r in range(15, ws.max_row + 1):
      if r > ws.max_row - 8:
        row_txt = " ".join(
            [str(ws.cell(r, c).value or "") for c in range(1, 5)]
        ).upper()
        if "CHECK" in row_txt or "SERVICE" in row_txt or "NOTE" in row_txt:
          continue

      no_val = str(ws.cell(r, 1).value or "").strip()
      tag = str(ws.cell(r, 2).value or "").strip()
      name = str(ws.cell(r, 3).value or "").strip()
      rem = str(ws.cell(r, 40).value or "").strip()
      if tag.endswith(".0"):
        tag = tag[:-2]
      clean_name = " ".join(name.upper().split())
      if not clean_name and not tag:
        continue

      sub_cells = [ws.cell(r, c_base + i) for i in range(3)]
      has_ps = any(
          c.fill and c.fill.fill_type == "darkHorizontal" for c in sub_cells
      )
      has_pc = any(
          c.fill
          and c.fill.fill_type == "solid"
          and getattr(c.fill.start_color, "index", None) in (0, 8)
          for c in sub_cells
      )

      if has_ps or has_pc:
        job_type = (
            "PERIODICAL SERVICE (PS)" if has_ps else "PERIODICAL CHECK (PC)"
        )
        stmps_list.append({
            "Sheet / Section": f"Sheet {sname}",
            "No": no_val,
            "Tag Equipment": tag,
            "Nama Mesin / Komponen": name,
            "Jenis Pekerjaan": job_type,
            "Bulan": m_name,
            "W1": "V" if has_ps or has_pc else "",
            "W2": "",
            "W3": "",
            "W4": "",
            "Instruksi Kerja / Remarks": rem,
        })

  df_stmps = pd.DataFrame(stmps_list)

  out_wb = openpyxl.Workbook()
  out_ws = out_wb.active
  out_ws.title = f"STMPS {m_name}"

  out_ws.merge_cells("A1:K1")
  out_ws["A1"] = (
      f"SHORT TERM PREVENTIVE MAINTENANCE SCHEDULE (STMPS) - {plant_name.upper()}"
  )
  out_ws["A1"].font = Font(name="Arial", size=14, bold=True)
  out_ws["A1"].alignment = Alignment(horizontal="center")

  out_ws.merge_cells("A2:K2")
  out_ws["A2"] = f"PERIODE BULAN: {m_name} {target_year}"
  out_ws["A2"].font = Font(name="Arial", size=11, bold=True)
  out_ws["A2"].alignment = Alignment(horizontal="center")

  headers = [
      "No",
      "Sheet/Section",
      "Tag Equipment",
      "Nama Mesin / Komponen",
      "Jenis Pekerjaan",
      "Bulan",
      "W1",
      "W2",
      "W3",
      "W4",
      "Standard IK / Remarks",
  ]
  out_ws.append([])
  out_ws.append(headers)

  header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
  header_fill = PatternFill(
      fill_type="solid", start_color="1F497D", end_color="1F497D"
  )
  thin_border = Border(
      left=Side(style="thin"),
      right=Side(style="thin"),
      top=Side(style="thin"),
      bottom=Side(style="thin"),
  )

  for col_num in range(1, 12):
    cell = out_ws.cell(row=4, column=col_num)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

  for r_idx, row in df_stmps.iterrows():
    r_data = [
        r_idx + 1,
        row["Sheet / Section"],
        row["Tag Equipment"],
        row["Nama Mesin / Komponen"],
        row["Jenis Pekerjaan"],
        row["Bulan"],
        row["W1"],
        row["W2"],
        row["W3"],
        row["W4"],
        row["Instruksi Kerja / Remarks"],
    ]
    out_ws.append(r_data)
    r_num = 4 + r_idx + 1
    for c_num in range(1, 12):
      c_cell = out_ws.cell(row=r_num, column=c_num)
      c_cell.border = thin_border
      if c_num in [1, 2, 3, 5, 6, 7, 8, 9, 10]:
        c_cell.alignment = Alignment(horizontal="center")

  for col in out_ws.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    col_letter = openpyxl.utils.get_column_letter(col[0].column)
    out_ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

  output_stmps = io.BytesIO()
  out_wb.save(output_stmps)
  output_stmps.seek(0)
  return output_stmps, df_stmps


# ================= TAB BAGIAN ATAS =================
tab_float, tab_rolled, tab_stmps = st.tabs([
    "🏭 LTPMS Float 1",
    "🏭 LTPMS Rolled Glass",
    "📅 STMPS Generator (Breakdown Bulanan)",
])

# --- TAB 1: FLOAT 1 ---
with tab_float:
  st.subheader("📋 Generator LTPMS Float 1")
  col1, col2 = st.columns(2)
  with col1:
    fl_f_n3 = st.file_uploader(
        f"1. File Float 1 Tahun {y_n3} (.xls / .xlsx)",
        type=["xls", "xlsx"],
        key="fl_f_n3",
    )
    fl_f_n2 = st.file_uploader(
        f"2. File Float 1 Tahun {y_n2} (.xls / .xlsx)",
        type=["xls", "xlsx"],
        key="fl_f_n2",
    )
  with col2:
    fl_f_n1 = st.file_uploader(
        f"3. File Float 1 Tahun {y_n1} (.xls / .xlsx)",
        type=["xls", "xlsx"],
        key="fl_f_n1",
    )
    fl_f_n = st.file_uploader(
        f"4. File Float 1 Tahun {y_n} (Master Basis) (.xlsx)",
        type=["xlsx"],
        key="fl_f_n",
    )

  if st.button(
      "🚀 Proses Audit & Buat LTPMS Float 1", type="primary", key="btn_fl"
  ):
    if not fl_f_n:
      st.error(f"File Master Basis Tahun {y_n} untuk Float 1 wajib diunggah!")
    else:
      with st.spinner("Sedang memproses LTPMS Float 1..."):
        out_stream, sched_items = process_ltpms(
            fl_f_n3, fl_f_n2, fl_f_n1, fl_f_n
        )
        st.success(f"✅ LTPMS Float 1 {target_year} Berhasil Disusun!")
        st.download_button(
            label=f"📥 Unduh File Excel LTPMS Float 1 {target_year}",
            data=out_stream,
            file_name=f"LTPMS_FLOAT_1_{target_year}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            key="dl_fl",
            type="primary",
        )
        st.dataframe(pd.DataFrame(sched_items), use_container_width=True)

# --- TAB 2: ROLLED GLASS ---
with tab_rolled:
  st.subheader("📋 Generator LTPMS Rolled Glass (Figur Glass)")
  col1, col2 = st.columns(2)
  with col1:
    rg_f_n3 = st.file_uploader(
        f"1. File Rolled Glass Tahun {y_n3} (.xls / .xlsx)",
        type=["xls", "xlsx"],
        key="rg_f_n3",
    )
    rg_f_n2 = st.file_uploader(
        f"2. File Rolled Glass Tahun {y_n2} (.xls / .xlsx)",
        type=["xls", "xlsx"],
        key="rg_f_n2",
    )
  with col2:
    rg_f_n1 = st.file_uploader(
        f"3. File Rolled Glass Tahun {y_n1} (.xls / .xlsx)",
        type=["xls", "xlsx"],
        key="rg_f_n1",
    )
    rg_f_n = st.file_uploader(
        f"4. File Rolled Glass Tahun {y_n} (Master Basis) (.xlsx)",
        type=["xlsx"],
        key="rg_n_basis",
    )

  if st.button(
      "🚀 Proses Audit & Buat LTPMS Rolled Glass", type="primary", key="btn_rg"
  ):
    if not rg_f_n:
      st.error(
          f"File Master Basis Tahun {y_n} untuk Rolled Glass wajib diunggah!"
      )
    else:
      with st.spinner("Sedang memproses LTPMS Rolled Glass..."):
        out_stream, sched_items = process_ltpms(
            rg_f_n3, rg_f_n2, rg_f_n1, rg_f_n
        )
        st.success(f"✅ LTPMS Rolled Glass {target_year} Berhasil Disusun!")
        st.download_button(
            label=f"📥 Unduh File Excel LTPMS Rolled Glass {target_year}",
            data=out_stream,
            file_name=f"LTPMS_ROLLED_GLASS_{target_year}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            key="dl_rg",
            type="primary",
        )
        st.dataframe(pd.DataFrame(sched_items), use_container_width=True)

# --- TAB 3: STMPS GENERATOR BULANAN ---
with tab_stmps:
  st.subheader(
      f"📅 Generator STMPS Bulan"
      f" {['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'][selected_month-1]}"
  )
  st.markdown(
      "Unggah file **LTPMS (Float 1 / Rolled Glass)** yang sudah jadi untuk"
      " mengekstrak seluruh pekerjaan pemeliharaan bulanan menjadi **Short Term"
      " Preventive Maintenance Schedule (STMPS)** siap cetak."
  )

  plant_choice = st.radio(
      "Pilih Plant:",
      ["Rolled Glass (Figur Glass)", "Float 1"],
      horizontal=True,
  )
  f_ltpms_target = st.file_uploader(
      "Unggah File Excel LTPMS Hasil Generate (.xlsx)",
      type=["xlsx"],
      key="f_stmps_inp",
  )

  if st.button(
      f"⚡ Generate STMPS"
      f" {['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'][selected_month-1]}",
      type="primary",
  ):
    if not f_ltpms_target:
      st.error("Silakan unggah berkas LTPMS terlebih dahulu!")
    else:
      with st.spinner(
          f"Sedang mengekstrak daftar pemeliharaan bulan ke-{selected_month}..."
      ):
        stmps_out, df_res = generate_stmps(
            f_ltpms_target, selected_month, plant_choice
        )
        st.success(
            f"✅ STMPS {plant_choice} Bulan {selected_month} Berhasil Disusun!"
        )
        st.download_button(
            label=(
                f"📥 Unduh File Excel STMPS {plant_choice} Bulan"
                f" {selected_month}"
            ),
            data=stmps_out,
            file_name=(
                f"STMPS_{plant_choice.replace(' ', '_')}_BULAN_{selected_month}_{target_year}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            type="primary",
        )
        st.dataframe(df_res, use_container_width=True)
