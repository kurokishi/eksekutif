from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side
import io


class ExcelWriter:

    def __init__(self, config):
        self.config = config
        self.fill_r = PatternFill(start_color="00FF00", fill_type="solid")  # Hijau R
        self.fill_e = PatternFill(start_color="0000FF", fill_type="solid")  # Biru E
        self.fill_over = PatternFill(start_color="FF0000", fill_type="solid")  # Merah overload

        # BORDER
        self.border_top_thick = Border(
            top=Side(border_style="thick")
        )

    def write(self, source_file, df, slot_str):
        wb = load_workbook(source_file)

        if "Jadwal" in wb.sheetnames:
            del wb["Jadwal"]

        ws = wb.create_sheet("Jadwal")

        headers = ["POLI ASAL", "JENIS POLI", "HARI", "DOKTER"] + slot_str
        ws.append(headers)

        for _, row in df.iterrows():
            ws.append([row.get(h, "") for h in headers])

        self.apply_styles(ws, df, slot_str)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    def apply_styles(self, ws, df, slot_str):
        """
        - Warnai R, E, overload
        - Hitung overload per hari per slot
        - Tambah border pemisah antar hari
        """

        # Hitungan E per hari per slot
        counter = {hari: {slot: 0 for slot in slot_str}
                   for hari in df["HARI"].unique()}

        records = df.to_dict("records")
        excel_row = 2

        last_hari = None  # untuk border pemisah antar hari

        for rowdata in records:
            hari = rowdata["HARI"]

            # === BORDER PEMISAH HARI ===
            if last_hari is not None and hari != last_hari:
                # beri border tebal di baris ini
                for col in range(1, len(rowdata) + len(slot_str)):
                    ws.cell(row=excel_row, column=col).border = self.border_top_thick

            last_hari = hari

            # warnai sel per slot
            for slot in slot_str:
                val = rowdata.get(slot, "")
                cell = ws.cell(row=excel_row, column=slot_str.index(slot) + 5)

                if val == "R":
                    cell.fill = self.fill_r

                elif val == "E":
                    counter[hari][slot] += 1

                    if counter[hari][slot] > self.config.max_poleks_per_slot:
                        cell.fill = self.fill_over
                    else:
                        cell.fill = self.fill_e

            excel_row += 1
