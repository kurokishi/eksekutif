from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io


class ExcelWriter:

    def __init__(self, config):
        self.config = config
        self.fill_r = PatternFill(start_color="00FF00", fill_type="solid")  # Hijau R
        self.fill_e = PatternFill(start_color="0000FF", fill_type="solid")  # Biru E
        self.fill_over = PatternFill(start_color="FF0000", fill_type="solid")  # Merah overload

    def write(self, source_file, df, slot_str):
        wb = load_workbook(source_file)

        # remove old sheet
        if "Jadwal" in wb.sheetnames:
            del wb["Jadwal"]

        ws = wb.create_sheet("Jadwal")

        headers = ["POLI ASAL", "JENIS POLI", "HARI", "DOKTER"] + slot_str
        ws.append(headers)

        for _, r in df.iterrows():
            ws.append([r.get(h, "") for h in headers])

        self.apply_styles(ws, df, slot_str)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    def apply_styles(self, ws, df, slot_str):
        """
        RULE:
        - Hitung jumlah E (Poleks) per hari per slot
        - Jika melebihi max_poleks_per_slot → baris berikutnya warna MERAH
        - Reset hitungan per hari
        """

        # Struktur counter: {hari: {slot: count}}
        counter = {hari: {slot: 0 for slot in slot_str}
                   for hari in df["HARI"].unique()}

        # Mapping row → dict df
        records = df.to_dict("records")

        # Excel rows start at row 2
        excel_row = 2

        for rowdata in records:
            hari = rowdata["HARI"]

            for slot in slot_str:
                value = rowdata.get(slot, "")

                excel_cell = ws.cell(row=excel_row, column=slot_str.index(slot) + 5)

                # beri warna reguler/poleks
                if value == "R":
                    excel_cell.fill = self.fill_r

                elif value == "E":
                    counter[hari][slot] += 1

                    # overload?
                    if counter[hari][slot] > self.config.max_poleks_per_slot:
                        excel_cell.fill = self.fill_over
                    else:
                        excel_cell.fill = self.fill_e

                # value kosong → tetap kosong, tidak diwarnai

            excel_row += 1
