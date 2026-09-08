import os
import re
import openpyxl
from openpyxl.styles import PatternFill
import xlrd


def build_lookup_xlrd(book, year):
  """Membaca data historis dari format file lama .xls (2023, 2024, 2025)."""
  lookup = {}
  for sname in book.sheet_names():
    if sname == "Sheet1":
      continue
    sh = book.sheet_by_name(sname)
    for r in range(14, sh.nrows):
      tag = str(sh.cell_value(r, 1)).strip() if sh.ncols > 1 else ""
      name = str(sh.cell_value(r, 2)).strip() if sh.ncols > 2 else ""
      rem = str(sh.cell_value(r, 39)).strip() if sh.ncols > 39 else ""
      if tag.endswith(".0"):
        tag = tag[:-2]
      clean_name = " ".join(name.upper().split())
      if not clean_name and not tag:
        continue

      ps_months = []
      for c in range(3, min(39, sh.ncols)):
        xf = book.xf_list[sh.cell_xf_index(r, c)]
        pat = xf.background.fill_pattern
        col = xf.background.pattern_colour_index
        is_ps = (
            col in (23, 24, 13)
            or pat == 5
            or (pat == 1 and col not in (8, 9, 64, 65))
        )
        if is_ps:
          m = ((c - 3) // 3) + 1
          if m not in ps_months:
            ps_months.append(m)
      rec = {
          "tag": tag,
          "name": name,
          "rem": rem,
          "ps_months": sorted(ps_months),
      }
      if clean_name:
        lookup[clean_name] = rec
      if tag:
        lookup[tag] = rec
  return lookup


def build_lookup_openpyxl(book, year):
  """Membaca data historis dari format file .xlsx (2026 REVISI)."""
  lookup = {}
  for sname in book.sheetnames:
    ws = book[sname]
    for r in range(15, ws.max_row + 1):
      tag = str(ws.cell(r, 2).value or "").strip()
      name = str(ws.cell(r, 3).value or "").strip()
      rem = str(ws.cell(r, 40).value or "").strip()
      if tag.endswith(".0"):
        tag = tag[:-2]
      clean_name = " ".join(name.upper().split())
      if not clean_name and not tag:
        continue

      ps_months = []
      for c in range(4, 40):
        cell = ws.cell(r, c)
        if cell.fill and cell.fill.fill_type == "darkHorizontal":
          m = ((c - 4) // 3) + 1
          if m not in ps_months:
            ps_months.append(m)
      rec = {
          "tag": tag,
          "name": name,
          "rem": rem,
          "ps_months": sorted(ps_months),
      }
      if clean_name:
        lookup[clean_name] = rec
      if tag:
        lookup[tag] = rec
  return lookup


def main():
  print("1. Membaca 4 File Riwayat Historis...")
  wb_2023 = xlrd.open_workbook("LTPMS FLOAT 1 2023.xls", formatting_info=True)
  wb_2024 = xlrd.open_workbook("LTPMS FLOAT 1 2024.xls", formatting_info=True)
  wb_2025 = xlrd.open_workbook("LTPMS FLOAT 1 2025.xls", formatting_info=True)
  wb_2026 = openpyxl.load_workbook("LTPMS FLOAT 1 2026 REVISI.xlsx")

  l23 = build_lookup_xlrd(wb_2023, 2023)
  l24 = build_lookup_xlrd(wb_2024, 2024)
  l25 = build_lookup_xlrd(wb_2025, 2025)
  l26 = build_lookup_openpyxl(wb_2026, 2026)

  # Master template output menggunakan basis 2026
  wb_out = openpyxl.load_workbook("LTPMS FLOAT 1 2026 REVISI.xlsx")
  wb_out["1"]["C7"].value = ":  2027"

  # Definisi pola arsir Excel
  ps_fill = PatternFill(
      fill_type="darkHorizontal",
      start_color="00000000",
      end_color="00000000",
  )  # Arsir Garis (PS)
  pc_fill = PatternFill(
      fill_type="solid", start_color="00000000", end_color="00000000"
  )  # Kotak Hitam Solid (PC)
  blank_fill = PatternFill(fill_type=None)  # Putih Bersih

  print("2. Memproses Audit dan Sinkronisasi Siklus 2027...")
  total_ps_count = 0
  total_pc_count = 0
  consec_fixed = 0

  for sname in wb_out.sheetnames:
    ws = wb_out[sname]
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

      # A. Parsing Siklus Servis Besar (PS)
      m_ps = re.search(r"PS\s*:\s*1\s*X\s*(\d+)\s*BLN", rem_upper)
      interval_ps = int(m_ps.group(1)) if m_ps else 12

      # B. Parsing Siklus Pengecekan Rutin (PC)
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

      # C. Riwayat Pelaksanaan Nyata
      p23 = l23.get(key, {}).get("ps_months", [])
      p24 = l24.get(key, {}).get("ps_months", [])
      p25 = l25.get(key, {}).get("ps_months", [])
      p26 = l26.get(key, {}).get("ps_months", [])

      # D. Logika Rekonsiliasi & Koreksi Anomali Berturut-turut
      target_ps_months = []
      if interval_ps == 12:
        target_ps_months = p26 if p26 else p25
      elif interval_ps == 24:
        if p24:
          # Anomali: Terjadwal di 2024 dan 2025.
          # Base aslinya Genap (2024 -> 2026 -> 2028).
          # Maka di 2027 WAJIB DI-SKIP!
          target_ps_months = []
          consec_fixed += 1
        elif p23 or p25:
          # Base aslinya Ganjil (2023 -> 2025 -> 2027).
          # Maka di 2027 WAJIB ADA!
          target_ps_months = p25 if p25 else p23
      elif interval_ps == 36:
        if p24:
          # Siklus 3 tahunan dari 2024 (2024 + 3 = 2027) -> Wajib Ada di 2027!
          target_ps_months = p24
        else:
          # Jika aktif di 2023 atau 2026 -> Belum jatuh tempo di 2027 (Skip)
          target_ps_months = []
      elif interval_ps == 48:
        if p23:
          # Siklus 4 tahunan dari 2023 (2023 + 4 = 2027) -> Wajib Ada di 2027!
          target_ps_months = p23
        else:
          target_ps_months = []
      else:
        target_ps_months = p26

      # E. Pewarnaan Kotak Penuh (3 Sub-Kolom per Bulan)
      for m in range(1, 13):
        c_base = 4 + (m - 1) * 3
        if m in target_ps_months:
          total_ps_count += 1
          for sub in range(3):
            ws.cell(r, c_base + sub).fill = ps_fill
        elif m in expected_pc_months:
          total_pc_count += 1
          for sub in range(3):
            ws.cell(r, c_base + sub).fill = pc_fill
        else:
          for sub in range(3):
            ws.cell(r, c_base + sub).fill = blank_fill

  output_name = "LTPMS FLOAT 1 2027 REVISI.xlsx"
  wb_out.save(output_name)
  print(f"Selesai! Dokumen berhasil dibuat: {output_name}")
  print(f"- Total Anomali Berturut-turut Dikoreksi: {consec_fixed} peralatan")
  print(f"- Total Bulan Servis (PS Penuh): {total_ps_count // 3} bulan")
  print(
      "- Total Bulan Pengecekan (PC Penuh Dipulihkan):"
      f" {total_pc_count // 3} bulan"
  )


if __name__ == "__main__":
  main()
