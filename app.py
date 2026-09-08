import io
import re
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
Aplikasi web ini mengotomatisasi penyusunan **Long Term Preventive Maintenance Schedule (LTPMS)** secara akurat.
- **Audit Siklus Ganjil/Genap (24 BLN):** Menyelaraskan mesin bergantian (misal Batch Charger Left/Right).
- **Pewarnaan Penuh Satu Bulan:** Seluruh sub-kolom per bulan terisi rapi tanpa garis terpotong.
- **Pemulihan PC Presisi:** Mengisi kotak hitam PC penuh sesuai instruksi kerja (IK).
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


def parse_file(uploaded_file, year):
  data = {}
  if not uploaded_file:
    return data
  content = uploaded_file.read()
  uploaded_file.seek(0)

  if uploaded_file.name.endswith(".xlsx"):
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=False)
    for sname in wb.sheetnames:
      ws = wb[sname]
      for r in range(15, ws.max_row + 1):
        tag = str(ws.cell(r, 2).value or "").strip()
        name = str(ws.cell(r, 3).value or "").strip()
        rem = str(ws.cell(r, 40).value or "").strip()
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
        data[key] = {
            "sheet": sname,
            "name": name,
            "tag": tag,
            "rem": rem,
            "months": sorted(months),
        }
  else:
    wb = xlrd.open_workbook(file_contents=content, formatting_info=True)
    for sname in wb.sheet_names():
      if sname == "Sheet1":
        continue
      sh = wb.sheet_by_name(sname)
      for r in range(14, sh.nrows):
        tag = str(sh.cell_value(r, 1)).strip() if sh.ncols > 1 else ""
        name = str(sh.cell_value(r, 2)).strip() if sh.ncols > 2 else ""
        rem = str(sh.cell_value(r, 39)).strip() if sh.ncols > 39 else ""
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
        data[key] = {
            "sheet": sname,
            "name": name,
            "tag": tag,
            "rem": rem,
            "months": sorted(months),
        }
  return data


if st.button("🚀 Proses Audit & Buat LTPMS", type="primary"):
  if not f_n:
    st.error(f"File Master Basis Tahun {y_n} wajib diunggah!")
  else:
    with st.spinner(
        "Sedang memproses rekonsiliasi dan membuat jadwal presisi..."
    ):
      h_n3 = parse_file(f_n3, y_n3) if f_n3 else {}
      h_n2 = parse_file(f_n2, y_n2) if f_n2 else {}
      h_n1 = parse_file(f_n1, y_n1) if f_n1 else {}

      content_master = f_n.read()
      f_n.seek(0)
      wb_master = openpyxl.load_workbook(
          io.BytesIO(content_master), data_only=False
      )

      # 1. Update Header Tahun di Sheet 1
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

      anomalies = []
      scheduled_items = []
      consec_fixed = 0

      for sname in wb_master.sheetnames:
        ws = wb_master[sname]
        for r in range(15, ws.max_row + 1):
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

          # Parsing PS
          m_ps = re.search(r"PS\s*:\s*1\s*X\s*(\d+)\s*BLN", rem_upper)
          interval_ps = int(m_ps.group(1)) if m_ps else 12

          # Parsing PC
          m_pc_bln = re.search(r"PC\s*:\s*1\s*X\s*(\d+)\s*BLN", rem_upper)
          pc_interval_bln = int(m_pc_bln.group(1)) if m_pc_bln else None
          has_pc_mgg = "MGG" in rem_upper

          expected_pc_months = set()
          if pc_interval_bln:
            for m in range(1, 13):
              if m % pc_interval_bln == 0:
                expected_pc_months.add(m)
          elif has_pc_mgg or "PC" in rem_upper:
            expected_pc_months = set(range(1, 13))

          # Riwayat tahun-tahun sebelumnya
          p_n3 = h_n3.get(key, {}).get("months", [])
          p_n2 = h_n2.get(key, {}).get("months", [])
          p_n1 = h_n1.get(key, {}).get("months", [])

          p_n = []
          for c in range(4, 40):
            cell = ws.cell(r, c)
            if cell.fill and cell.fill.fill_type == "darkHorizontal":
              m = ((c - 4) // 3) + 1
              if m not in p_n:
                p_n.append(m)

          target_ps_months = []
          status_sched = ""

          if interval_ps == 12:
            target_ps_months = p_n if p_n else p_n1
            status_sched = "Rutin Tahunan (12 BLN)"
          elif interval_ps == 24:
            # Evaluasi Siklus 24 Bulan:
            # Jika ada di N-3 (2023) dan N-1 (2025) -> Siklus GANJIL (2023 -> 2025 -> 2027) => WAJIB ADA di 2027!
            # Kesalahan muncul di 2024/2026 dianulir.
            if p_n3 and p_n1:
              target_ps_months = p_n1 if p_n1 else p_n3
              status_sched = (
                  f"Due Siklus Ganjil ({y_n3} -> {y_n1} -> {target_year})"
              )
              if p_n2 or p_n:
                anomalies.append({
                    "Sheet": sname,
                    "Tag": tag,
                    "Nama Mesin": name,
                    "Instruksi Kerja": "1x24 BLN",
                    "Anomali": (
                        f"Salah input di tahun genap ({y_n2}/{y_n}), siklus asli"
                        f" ganjil ({y_n3} & {y_n1})"
                    ),
                    "Tindakan": (
                        f"Tetap dijadwalkan di {target_year} sesuai siklus"
                        " ganjil aslinya"
                    ),
                })
            elif p_n2:  # Murni siklus genap (2024 -> 2026 -> 2028)
              target_ps_months = []
              status_sched = (
                  f"Siklus Genap (Base {y_n2} -> Skip {target_year}, next"
                  f" {target_year+1})"
              )
              consec_fixed += 1
              if p_n1:
                anomalies.append({
                    "Sheet": sname,
                    "Tag": tag,
                    "Nama Mesin": name,
                    "Instruksi Kerja": "1x24 BLN",
                    "Anomali": f"Muncul berulang di {y_n2} & {y_n1}",
                    "Tindakan": (
                        f"Diistirahatkan (Skip) di {target_year}, dialihkan ke"
                        f" {target_year+1}"
                    ),
                })
            elif p_n1 or p_n3:
              target_ps_months = p_n1 if p_n1 else p_n3
              status_sched = f"Due dari Siklus {y_n1} (24 BLN)"
          elif interval_ps == 36:
            if p_n2:  # 2024 + 3 = 2027
              target_ps_months = p_n2
              status_sched = f"Due dari Siklus {y_n2} (36 BLN)"
            else:
              target_ps_months = []
              status_sched = (
                  f"Belum Jatuh Tempo di {target_year} (Siklus 36 BLN)"
              )
          elif interval_ps == 48:
            if p_n3:  # 2023 + 4 = 2027
              target_ps_months = p_n3
              status_sched = f"Due dari Siklus {y_n3} (48 BLN)"
            else:
              target_ps_months = []
              status_sched = (
                  f"Belum Jatuh Tempo di {target_year} (Siklus 48 BLN)"
              )
          else:
            target_ps_months = p_n

          # Tulis kotak penuh (3 sub-kolom per bulan)
          for m in range(1, 13):
            c_base = 4 + (m - 1) * 3
            if m in target_ps_months:
              for sub in range(3):
                ws.cell(r, c_base + sub).fill = ps_fill
            elif m in expected_pc_months:
              for sub in range(3):
                ws.cell(r, c_base + sub).fill = pc_fill
            else:
              for sub in range(3):
                ws.cell(r, c_base + sub).fill = blank_fill

          if target_ps_months:
            scheduled_items.append({
                "Sheet": sname,
                "Nama Mesin": name,
                "Interval": f"1x{interval_ps} BLN",
                "Keterangan": status_sched,
                "Bulan Pelaksanaan": str(target_ps_months),
            })

      output_stream = io.BytesIO()
      wb_master.save(output_stream)
      output_stream.seek(0)

      st.success(
          f"✅ LTPMS {target_year} Berhasil Digenerate! (Batch Charger Left &"
          " mesin ganjil berhasil dijadwalkan)"
      )

      st.download_button(
          label=f"📥 Unduh File Excel LTPMS {target_year}",
          data=output_stream,
          file_name=f"LTPMS_FLOAT_1_{target_year}.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          type="primary",
      )

      tab1, tab2 = st.tabs(
          ["⚠️ Temuan Anomali & Koreksi", "📋 Daftar Servis Terjadwal"]
      )
      with tab1:
        if anomalies:
          st.dataframe(pd.DataFrame(anomalies), use_container_width=True)
        else:
          st.info("Semua siklus berjalan normal.")
      with tab2:
        st.dataframe(pd.DataFrame(scheduled_items), use_container_width=True)
