from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io


class ExcelWriter:

    def __init__(self, config):
        self.config = config

    def write(self, source_file, df, slot_str):
        wb = load_workbook(source_file)

        if "Jadwal" in wb.sheetnames:
            del wb["Jadwal"]

        ws = wb.create_sheet("Jadwal")

        headers = ["POLI ASAL", "JENIS POLI", "HARI", "DOKTER"] + slot_str
        ws.append(headers)

        for _, row in df.iterrows():
            ws.append([row.get(h, "") for h in headers])

        self.colorize(ws, df, slot_str)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    def colorize(self, ws, df, slot_str):
        fill_r = PatternFill(start_color="00FF00", fill_type="solid")
        fill_e = PatternFill(start_color="0000FF", fill_type="solid")

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if cell.value == "R":
                    cell.fill = fill_r
                elif cell.value == "E":
                    cell.fill = fill_e
