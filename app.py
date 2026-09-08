import io
import re
from collections import Counter
import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd
import streamlit as st
import xlrd

st.set_page_config(
    page_title="LTPMS Generator & Smart Audit", page_icon="⚙️", layout="wide"
)

st.title("⚙️ LTPMS Multi-Year Generator & Smart Audit System")
st.markdown("""
Aplikasi web ini menyusun **Long Term Preventive Maintenance Schedule (LTPMS)** secara akurat:
- **Penambal Otomatis Kotak PC yang Bolong:** Memulihkan kotak hitam PC yang hilang akibat bekas servis lama (misal Bulan 4 pada Blower, Bulan 11, Bulan 7 pada Batch Charger, dll.).
- **Preservasi Baris Bersih:** Menjaga baris komponen tetap bersih putih tanpa menimbulkan balok hitam liar.
- **Koreksi Siklus Multi-Tahun:** Mengelola ritme servis 24, 36, dan 48 bulan secara presisi.
""")

# Sidebar settings
st.sidebar.header("🔧 Pengaturan Target")
target_year = st.sidebar.number_input(
    "Target Tahun LTPMS", min_value=2025, max_value=2040, value=2027, step=1
)
y_n = target_year - 1
y_n1 = target_year - 2
y_n2 = target_year - 3
y_n3 = target_year - 4

st.sidebar.info(f"""
**Histori yang dibutuhkan:**
- Tahun N-3: **{y_n3}**
- Tahun N-2: **{y_n2}**
- Tahun N-1: **{y_n1}**
- Tahun N (Master Basis): **{y_n}**
""")

st.subheader("📁 Unggah File Riwayat LTPMS")
col1, col2 = st.columns(2)
with col1:
  f_n3 = st.file_uploader(
      f"1. File Tahun {y_n3} (.xls / .xlsx)", type=["xls", "xlsx"], key="f_n3"
  )
  f_n2 = st.file_uploader(
      f"2. File Tahun {y_n2} (.xls / .xlsx)", type=["xls", "xlsx"], key="f_n2"
  )
with col2:
  f_n1 = st.file_uploader(
      f"3. File Tahun {y_n1} (.xls / .xlsx)", type=["xls", "xlsx"], key="f_n1"
  )
  f_n = st.file_uploader(
      f"4. File Tahun {y_n} (Master Basis) (.xlsx)", type=["xlsx"], key="f_n"
  )


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
      if sname == "Sheet1":
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


if st.button("🚀 Proses Audit & Buat LTPMS", type="primary"):
  if not f_n:
    st.error(f"File Master Basis Tahun {y_n} wajib diunggah!")
  else:
    with st.spinner(
        "Sedang menyelaraskan jadwal dan menambal kotak PC yang bolong..."
    ):
      ps_n3 = parse_history_ps(f_n3, y_n3) if f_n3 else {}
      ps_n2 = parse_history_ps(f_n2, y_n2) if f_n2 else {}
      ps_n1 = parse_history_ps(f_n1, y_n1) if f_n1 else {}

      content_master = f_n.read()
      f_n.seek(0)
      wb_master = openpyxl.load_workbook(
          io.BytesIO(content_master), data_only=False
      )

      # Update Header Tahun di Sheet 1
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
      restored_pc_info = []

      # Komponen murni yang tidak pernah memiliki baris PC
      pure_sub_components = [
          "MOTOR",
          "BELT UNIT",
          "GEARBOX",
          "REDUCER",
          "SHAFT",
          "COUPLING",
      ]

      for sname in wb_master.sheetnames:
        ws = wb_master[sname]
        for r in range(15, ws.max_row + 1):
          # Abaikan area catatan / legenda bawah
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

          # Cek apakah ini sub-komponen murni (Motor, Belt Unit)
          is_pure_sub = any(
              clean_name.startswith(kw) or f" {kw} " in f" {clean_name} "
              for kw in pure_sub_components
          )

          # 1. Deteksi bulan PC asli yang sudah ada di baris ini
          existing_pc_months = set()
          p_n = []
          for c in range(4, 40):
            cell = ws.cell(r, c)
            m = ((c - 4) // 3) + 1
            if (
                cell.fill
                and cell.fill.fill_type == "solid"
                and getattr(cell.fill.start_color, "index", None) in (0, 8)
            ):
              existing_pc_months.add(m)
            elif cell.fill and cell.fill.fill_type == "darkHorizontal":
              if m not in p_n:
                p_n.append(m)

          # 2. Logika "Rhythmic Gap Healing" untuk kotak PC yang bolong
          final_pc_months = set(existing_pc_months)
          m_pc = re.search(r"PC\s*:\s*1\s*X\s*(\d+)\s*BLN", rem_upper)
          pc_int = int(m_pc.group(1)) if m_pc else None

          # Jika baris ini memiliki kotak PC dan memiliki ritme interval (misal 2, 3, 4 bulan)
          if not is_pure_sub and pc_int and pc_int > 1 and existing_pc_months:
            # Hitung offset ritme yang paling dominan
            offsets = [m % pc_int for m in existing_pc_months]
            dominant_offset = Counter(offsets).most_common(1)[0][0]

            # Bentuk ritme utuh sepanjang 12 bulan
            full_cycle = [
                m for m in range(1, 13) if (m % pc_int) == dominant_offset
            ]
            for m in full_cycle:
              if m not in final_pc_months:
                final_pc_months.add(m)
                restored_pc_info.append({
                    "Sheet": sname,
                    "Nama Mesin": name,
                    "Bulan Ditambal": f"Bulan {m}",
                    "Ritme": f"PC 1x{pc_int} BLN",
                })

          # 3. Penentuan Jadwal Servis Besar (PS)
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
              status_sched = (
                  f"Siklus Genap (Base {y_n2} -> Skip {target_year}, next"
                  f" {target_year+1})"
              )
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

          # 4. Pewarnaan Sel 1 Tahun Penuh
          for m in range(1, 13):
            c_base = 4 + (m - 1) * 3
            if m in target_ps_months:
              # Prioritas utama: Servis Besar (PS) menimpa bulan tersebut
              for sub in range(3):
                ws.cell(r, c_base + sub).fill = ps_fill
            elif m in final_pc_months and not is_pure_sub:
              # Kotak hitam PC ritme lengkap
              for sub in range(3):
                ws.cell(r, c_base + sub).fill = pc_fill
            else:
              # Bersih putih
              for sub in range(3):
                ws.cell(r, c_base + sub).fill = blank_fill

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

      st.success(
          f"✅ LTPMS {target_year} Berhasil Disempurnakan! ({len(restored_pc_info)}"
          " kotak PC yang bolong berhasil dipulihkan)"
      )

      st.download_button(
          label=f"📥 Unduh File Excel LTPMS {target_year}",
          data=output_stream,
          file_name=f"LTPMS_FLOAT_1_{target_year}.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          type="primary",
      )

      tab1, tab2 = st.tabs(
          ["📋 Daftar Servis Terjadwal", "🩹 Kotak PC yang Dipulihkan"]
      )
      with tab1:
        st.dataframe(pd.DataFrame(scheduled_items), use_container_width=True)
      with tab2:
        if restored_pc_info:
          st.info(
              f"Daftar {len(restored_pc_info)} kotak hitam PC yang berhasil"
              " ditambal di posisi ritme yang seharusnya:"
          )
          st.dataframe(pd.DataFrame(restored_pc_info), use_container_width=True)
        else:
          st.info("Semua ritme PC sudah lengkap.")
