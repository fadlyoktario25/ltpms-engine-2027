import io
import re
import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd
import streamlit as st
import xlrd

st.set_page_config(
    page_title="LTPMS Generator & Audit System", page_icon="⚙️", layout="wide"
)

st.title("⚙️ LTPMS Multi-Year Generator & Smart Audit System")
st.markdown("""
Aplikasi web ini mengotomatisasi penyusunan **Long Term Preventive Maintenance Schedule (LTPMS)** secara akurat.
Sistem menganalisis siklus riwayat hingga 4 tahun ke belakang, mendeteksi anomali (*over-maintenance* / *overdue*), serta **memulihkan kotak PC (Periodical Check)** yang sempat tertimpa atau kosong akibat pergeseran jadwal PS (Periodical Service).
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
        "Sedang memproses, memulihkan kotak PC, dan mengaudit siklus data..."
    ):
      h_n3 = parse_file(f_n3, y_n3) if f_n3 else {}
      h_n2 = parse_file(f_n2, y_n2) if f_n2 else {}
      h_n1 = parse_file(f_n1, y_n1) if f_n1 else {}

      content_master = f_n.read()
      f_n.seek(0)
      wb_master = openpyxl.load_workbook(
          io.BytesIO(content_master), data_only=False
      )

      h_n = {}
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
          months = []
          for c in range(4, 40):
            cell = ws.cell(r, c)
            if cell.fill and cell.fill.fill_type == "darkHorizontal":
              m = ((c - 4) // 3) + 1
              if m not in months:
                months.append(m)
          key = tag if tag else clean_name
          h_n[key] = {
              "sheet": sname,
              "name": name,
              "tag": tag,
              "rem": rem,
              "months": sorted(months),
          }

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
      restored_pc_count = 0

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
          m_ps = re.search(r"PS\s*:\s*1\s*X\s*(\d+)\s*BLN", rem_upper)
          interval_ps = int(m_ps.group(1)) if m_ps else 12

          m_pc = re.search(r"PC\s*:\s*([^\:\n]+)", rem_upper)
          pc_text = m_pc.group(1).strip() if m_pc else ""
          has_pc_bln = "1X1 BLN" in pc_text.replace(
              " ", ""
          ) or "1X1BLN" in pc_text.replace(" ", "")
          has_pc_mgg = "1X1 MGG" in pc_text.replace(
              " ", ""
          ) or "1X1MGG" in pc_text.replace(" ", "")

          p_n3 = h_n3.get(key, {}).get("months", [])
          p_n2 = h_n2.get(key, {}).get("months", [])
          p_n1 = h_n1.get(key, {}).get("months", [])
          p_n = h_n.get(key, {}).get("months", [])

          if interval_ps >= 24:
            active_years = []
            if p_n3:
              active_years.append(y_n3)
            if p_n2:
              active_years.append(y_n2)
            if p_n1:
              active_years.append(y_n1)
            if p_n:
              active_years.append(y_n)

            for idx in range(1, len(active_years)):
              diff = active_years[idx] - active_years[idx - 1]
              if diff < (interval_ps // 12):
                anomalies.append({
                    "Sheet": sname,
                    "Tag": tag,
                    "Nama Mesin": name,
                    "Instruksi Kerja": f"1x{interval_ps} BLN",
                    "Status Anomali": (
                        f"Muncul berdekatan/berturut di {active_years[idx-1]} &"
                        f" {active_years[idx]}"
                    ),
                    "Rekomendasi": (
                        f"Koreksi jadwal agar jeda {interval_ps // 12} tahun"
                    ),
                })
                break

          target_months = []
          status_sched = ""
          if interval_ps == 12:
            target_months = p_n if p_n else p_n1
            status_sched = "Rutin Tahunan (12 BLN)"
          elif interval_ps == 24:
            if p_n1:
              target_months = p_n1
              status_sched = f"Due dari {y_n1} (2 Tahun)"
            elif p_n:
              target_months = []
              status_sched = (
                  f"Skip (Sudah di {y_n}, next {target_year + 1})"
              )
            elif p_n3:
              target_months = p_n3
              status_sched = f"Siklus dari {y_n3}"
          elif interval_ps == 36:
            if p_n2:
              target_months = p_n2
              status_sched = f"Due dari {y_n2} (3 Tahun)"
            elif p_n or p_n1:
              target_months = []
              status_sched = f"Skip (Sudah di {y_n if p_n else y_n1})"
          elif interval_ps == 48:
            if p_n3:
              target_months = p_n3
              status_sched = f"Due dari {y_n3} (4 Tahun)"
            elif p_n or p_n1 or p_n2:
              target_months = []
              status_sched = "Skip (Belum Jatuh Tempo)"
          else:
            target_months = p_n
            status_sched = "Mengikuti Pola Master"

          for m in range(1, 13):
            c_base = 4 + (m - 1) * 3
            for c_sub in range(c_base, c_base + 3):
              cell = ws.cell(r, c_sub)
              if cell.fill and cell.fill.fill_type == "darkHorizontal":
                if m not in target_months and (has_pc_bln or has_pc_mgg):
                  cell.fill = pc_fill
                  restored_pc_count += 1
                else:
                  cell.fill = blank_fill

          if target_months:
            for m in target_months:
              c_base = 4 + (m - 1) * 3
              ws.cell(r, c_base).fill = ps_fill
            scheduled_items.append({
                "Sheet": sname,
                "Nama Mesin": name,
                "Interval": f"1x{interval_ps} BLN",
                "Keterangan": status_sched,
                "Bulan Pelaksanaan": str(target_months),
            })

      output_stream = io.BytesIO()
      wb_master.save(output_stream)
      output_stream.seek(0)

      st.success(
          f"✅ LTPMS {target_year} Berhasil Digenerate! (Total"
          f" {restored_pc_count} kotak PC berhasil dipulihkan dari bekas PS"
          " lama)"
      )

      st.download_button(
          label=f"📥 Unduh File Excel LTPMS {target_year}",
          data=output_stream,
          file_name=f"LTPMS_FLOAT_1_{target_year}.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          type="primary",
      )

      tab1, tab2 = st.tabs(
          ["⚠️ Temuan Anomali Jadwal", "📋 Daftar Servis Terjadwal"]
      )
      with tab1:
        if anomalies:
          st.warning(
              f"Ditemukan {len(anomalies)} anomali jadwal historis"
              " (Over-Maintenance / Jadwal Tumpang Tindih):"
          )
          st.dataframe(pd.DataFrame(anomalies), use_container_width=True)
        else:
          st.info("Semua siklus historis berjalan normal sesuai interval.")

      with tab2:
        st.dataframe(pd.DataFrame(scheduled_items), use_container_width=True)
