# app/core/excel_writer.py

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io

class ExcelWriter:

    def __init__(self, config):
        self.config = config

    def write(self, source_file, df, slot_str):
        wb = load_workbook(source_file)

        if 'Jadwal' in wb.sheetnames:
            wb.remove(wb['Jadwal'])

        ws = wb.create_sheet("Jadwal")

        headers = ['POLI ASAL', 'JENIS POLI', 'HARI', 'DOKTER'] + slot_str
        ws.append(headers)

        for _, r in df.iterrows():
            ws.append([r[h] for h in headers])

        self.style(ws, df, slot_str)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    def style(self, ws, df, slot_str):
        fill_r = PatternFill(start_color="00FF00", fill_type="solid")
        fill_e = PatternFill(start_color="0000FF", fill_type="solid")
        fill_o = PatternFill(start_color="FF0000", fill_type="solid")

        day_slots = {h: {s: 0 for s in slot_str} for h in self.config.hari_list}

        for idx, row in df.iterrows():
            hari = row['HARI']
            for s in slot_str:
                if row[s] == 'E':
                    day_slots[hari][s] += 1

        max_row = len(df) + 1

        for r in range(2, max_row + 1):
            hari = ws.cell(r, 3).value
            for i, s in enumerate(slot_str, 5):
                cell = ws.cell(r, i)
                val = cell.value

                if val == 'R':
                    cell.fill = fill_r
                elif val == 'E':
                    cell.fill = fill_e
                    if day_slots[hari][s] > self.config.max_poleks_per_slot:
                        cell.fill = fill_o
