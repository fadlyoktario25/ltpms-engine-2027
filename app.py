import io
import re
import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd
import streamlit as st
import xlrd

st.set_page_config(
    page_title="Portal Teknik LTPMS & STMPS", page_icon="⚙️", layout="wide"
)

st.title("⚙️ Portal Utama Maintenance (LTPMS & STMPS Generator)")
st.markdown("""
Pusat otomatisasi dokumen perencanaan pemeliharaan jangka panjang (**LTPMS**) dan break-down bulanan/harian (**STMPS**).
""")

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
      fill_type="darkHorizontal", start_color="00000000", end_color="00000000"
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


def generate_stmps_layout(uploaded_ltpms, target_m, target_y, plant_name):
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

  content = uploaded_ltpms.read()
  uploaded_ltpms.seek(0)
  wb_lt = openpyxl.load_workbook(io.BytesIO(content), data_only=False)

  ps_fill = PatternFill(
      fill_type="darkHorizontal", start_color="00000000", end_color="00000000"
  )
  pc_fill = PatternFill(
      fill_type="solid", start_color="00000000", end_color="00000000"
  )
  blank_fill = PatternFill(fill_type=None)

  c_base = 4 + (target_m - 1) * 3  # Kolom bulan target di LTPMS

  # Pembagian Hari Kerja (Senin-Jumat) Menghindari Sabtu/Minggu (Tanggal Merah)
  week_working_days = {
      0: [4, 5, 6, 7, 8],  # Minggu 1 (Tgl 4 - 8)
      1: [11, 12, 13, 14, 15],  # Minggu 2 (Tgl 11 - 15)
      2: [18, 19, 20, 21, 22],  # Minggu 3 (Tgl 18 - 22)
      3: [25, 26, 27, 28, 29],  # Minggu 4 (Tgl 25 - 29)
  }

  plant_label = "PIGUR GLASS" if "ROLLED" in plant_name.upper() else "FLOAT 1"
  summary_rows = []

  for sname in wb_lt.sheetnames:
    ws = wb_lt[sname]

    # 1. Ekstrak pekerjaan bulan target dari LTPMS Rolled Glass
    row_jobs = {}
    for r in range(15, ws.max_row + 1):
      if r > ws.max_row - 12:
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
      key = tag if tag else clean_name
      if not key:
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
        sub_w = 0
        for idx_s, c_c in enumerate(sub_cells):
          if c_c.fill and c_c.fill.fill_type:
            sub_w = idx_s
            break
        row_jobs[r] = (
            "PS" if has_ps else "PC",
            sub_w,
            key,
            tag,
            name,
            rem,
        )

    # 2. Unmerge Header Bulan di Baris 13-14 (Kolom 4 sd 39)
    ranges_to_remove = []
    for rng in ws.merged_cells.ranges:
      if (
          rng.min_col >= 4
          and rng.max_col <= 39
          and (rng.min_row <= 14 and rng.max_row >= 13)
      ):
        ranges_to_remove.append(rng)
    for rng in ranges_to_remove:
      ws.unmerge_cells(str(rng))

    # 3. Update Header Dokumen Presisi Sesuai Layout Pabrik
    ws["A1"].value = f"SHORT TERM P/M {plant_label} --- SCHEDULE"
    ws["A2"].value = "Doc No : QR/ENG/MEIU/25, REV: 04"
    ws["C7"].value = f":   {m_name} {target_y}"  # C7 PERIODE
    ws["C13"].value = "DATE     "  # C13 DATE

    # 4. Set Header Tanggal 1..31 di Kolom 4..34
    for d in range(1, 32):
      ws.cell(13, 3 + d).value = float(d)

    # 5. Bersihkan seluruh warna di area bulan (Kolom 4 s/d 39)
    for r in range(15, ws.max_row + 1):
      for c in range(4, 40):
        ws.cell(r, c).fill = blank_fill

    # 6. Pasang arsir HANYA di hari kerja (Menghindari tanggal merah/weekend)
    for r, (j_type, sub_w, key, tag, name, rem) in row_jobs.items():
      target_fill = ps_fill if j_type == "PS" else pc_fill
      assigned_days = week_working_days.get(sub_w, [4, 5, 6, 7, 8])

      for d in assigned_days:
        ws.cell(r, 3 + d).fill = target_fill

      summary_rows.append({
          "Sheet": sname,
          "Tag Equipment": tag,
          "Nama Mesin": name,
          "Jenis Pekerjaan": j_type,
          "Hari Kerja Execusi": (
              f"Tgl {assigned_days[0]} s/d {assigned_days[-1]} {m_name}"
          ),
          "Remarks": rem,
      })

  output_stmps = io.BytesIO()
  wb_lt.save(output_stmps)
  output_stream = output_stmps
  output_stream.seek(0)
  return output_stream, pd.DataFrame(summary_rows)


# ================= TAB BAGIAN ATAS =================
tab_float, tab_rolled, tab_stmps = st.tabs([
    "🏭 LTPMS Float 1",
    "🏭 LTPMS Rolled Glass",
    "📅 STMPS Rolled Glass Oktober",
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

# --- TAB 3: STMPS ROLLED GLASS OKTOBER ---
with tab_stmps:
  st.subheader(
      f"📅 Generator STMPS Rolled Glass (Bulan"
      f" {['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'][selected_month-1]})"
  )
  st.markdown(
      "Unggah file **LTPMS Rolled Glass** untuk menghasilkan dokumen **STMPS"
      " Format Asli Pabrik** yang langsung menyesuaikan hari kerja (menghindari"
      " Sabtu/Minggu)."
  )

  f_ltpms_target = st.file_uploader(
      "Upload File LTPMS Rolled Glass Basis (.xlsx)",
      type=["xlsx"],
      key="f_stmps_ltpms",
  )

  if st.button(
      f"⚡ Generate STMPS Rolled Glass Bulan"
      f" {['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'][selected_month-1]}",
      type="primary",
  ):
    if not f_ltpms_target:
      st.error("Silakan unggah berkas LTPMS Rolled Glass terlebih dahulu!")
    else:
      with st.spinner(
          f"Sedang memproses STMPS Rolled Glass bulan ke-{selected_month}..."
      ):
        stmps_out, df_sum = generate_stmps_layout(
            f_ltpms_target, selected_month, target_year, "Rolled Glass"
        )
        m_name_up = [
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
        ][selected_month - 1]

        st.success(
            f"✅ STMPS PIGUR GLASS {m_name_up} {target_year} Berhasil Disusun!"
        )
        st.download_button(
            label=(
                f"📥 Unduh File Excel SHORT TERM PIGURE GLASS {m_name_up}"
                f" {target_year}"
            ),
            data=stmps_out,
            file_name=(
                f"SHORT_TERM_PIGURE_GLASS_{m_name_up}_{target_year}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            type="primary",
        )
        st.dataframe(df_sum, use_container_width=True)
