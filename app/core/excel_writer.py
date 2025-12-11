from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
import io
from datetime import datetime, timedelta
import pandas as pd


class ExcelWriter:

    def __init__(self, config):
        self.config = config

        # warna slot
        self.fill_r = PatternFill(start_color="00FF00", fill_type="solid")
        self.fill_e = PatternFill(start_color="0000FF", fill_type="solid")
        self.fill_over = PatternFill(start_color="FF0000", fill_type="solid")

        # border tebal untuk pemisah hari
        self.border_top_thick = Border(top=Side(border_style="thick"))

        # border header
        self.border_header = Border(bottom=Side(border_style="thick"))

    # --------------------------------------------------------
    # Helper: gabung rentang waktu slot per 30 menit
    # --------------------------------------------------------
    def _combine_ranges(self, slots, interval):
        if not slots:
            return []

        ts = [datetime.strptime(t, "%H:%M") for t in slots]
        ts.sort()

        ranges = []
        start = ts[0]
        end = start + timedelta(minutes=interval)

        for t in ts[1:]:
            if t == end:
                end = t + timedelta(minutes=interval)
            else:
                ranges.append((start, end))
                start = t
                end = start + timedelta(minutes=interval)

        ranges.append((start, end))
        return ranges

    def _format_range(self, a, b):
        return f"{a.strftime('%H.%M')}–{b.strftime('%H.%M')}"

    # --------------------------------------------------------
    # SHEET: Rekap Layanan per Dokter
    # --------------------------------------------------------
    def _create_rekap_layanan(self, wb, df, slot_str):
        if "Rekap Layanan" in wb.sheetnames:
            del wb["Rekap Layanan"]

        ws = wb.create_sheet("Rekap Layanan")
        ws.append(["POLI ASAL", "HARI", "DOKTER", "JENIS", "JAM LAYANAN"])

        interval = self.config.interval_minutes

        for (poli, hari, dokter), g in df.groupby(
            ["POLI ASAL", "HARI", "DOKTER"]
        ):
            r_slots = [s for s in slot_str if g.iloc[0].get(s, "") == "R"]
            e_slots = [s for s in slot_str if g.iloc[0].get(s, "") == "E"]

            for a, b in self._combine_ranges(r_slots, interval):
                ws.append([poli, hari, dokter, "Reguler", self._format_range(a, b)])

            for a, b in self._combine_ranges(e_slots, interval):
                ws.append([poli, hari, dokter, "Poleks", self._format_range(a, b)])

    # --------------------------------------------------------
    # SHEET: Rekap Poli
    # --------------------------------------------------------
    def _create_rekap_poli(self, wb, df, slot_str):
        if "Rekap Poli" in wb.sheetnames:
            del wb["Rekap Poli"]

        ws = wb.create_sheet("Rekap Poli")
        ws.append([
            "POLI",
            "HARI",
            "TOTAL JAM REGULER",
            "TOTAL JAM POLEKS",
            "TOTAL JAM LAYANAN"
        ])

        interval = self.config.interval_minutes

        for (poli, hari), g in df.groupby(["POLI ASAL", "HARI"]):
            tot_r = 0
            tot_e = 0
            for slot in slot_str:
                v = g.iloc[0].get(slot, "")
                if v == "R":
                    tot_r += interval / 60
                elif v == "E":
                    tot_e += interval / 60

            ws.append([
                poli,
                hari,
                round(tot_r, 2),
                round(tot_e, 2),
                round(tot_r + tot_e, 2)
            ])

    # --------------------------------------------------------
    # SHEET: Rekap Dokter
    # --------------------------------------------------------
    def _create_rekap_dokter(self, wb, df, slot_str):
        if "Rekap Dokter" in wb.sheetnames:
            del wb["Rekap Dokter"]

        ws = wb.create_sheet("Rekap Dokter")
        ws.append([
            "DOKTER",
            "HARI",
            "TOTAL JAM REGULER",
            "TOTAL JAM POLEKS",
            "TOTAL JAM"
        ])

        interval = self.config.interval_minutes

        for (dokter, hari), g in df.groupby(["DOKTER", "HARI"]):
            tot_r = 0
            tot_e = 0

            for slot in slot_str:
                v = g.iloc[0].get(slot, "")
                if v == "R":
                    tot_r += interval / 60
                elif v == "E":
                    tot_e += interval / 60

            ws.append([
                dokter,
                hari,
                round(tot_r, 2),
                round(tot_e, 2),
                round(tot_r + tot_e, 2)
            ])

    # --------------------------------------------------------
    # SHEET: Grafik Beban Poli
    # --------------------------------------------------------
    def _create_grafik_poli(self, wb):
        if "Grafik Beban Poli" in wb.sheetnames:
            del wb["Grafik Beban Poli"]

        ws = wb.create_sheet("Grafik Beban Poli")
        ws["A1"] = "Grafik Beban Poli (Total Jam Layanan per Minggu)"

        rp = wb["Rekap Poli"]

        table = {}
        for row in rp.iter_rows(min_row=2, values_only=True):
            poli = row[0]
            total = row[4]
            table[poli] = table.get(poli, 0) + total

        ws.append(["POLI", "TOTAL JAM"])

        for k, v in table.items():
            ws.append([k, v])

        chart = BarChart()
        chart.title = "Beban Poli"

        data = Reference(ws, min_col=2, min_row=2, max_row=ws.max_row)
        cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)

        chart.add_data(data)
        chart.set_categories(cats)
        chart.y_axis.title = "Jam"

        ws.add_chart(chart, "E5")

    # --------------------------------------------------------
    # SHEET UTAMA: Jadwal
    # --------------------------------------------------------
    def write(self, source_file, df, slot_str):
        wb = load_workbook(source_file)

        if "Jadwal" in wb.sheetnames:
            del wb["Jadwal"]

        ws = wb.create_sheet("Jadwal")
        headers = ["POLI ASAL", "JENIS POLI", "HARI", "DOKTER"] + slot_str
        ws.append(headers)

        for _, row in df.iterrows():
            ws.append([row.get(h, "") for h in headers])

        # Warnai jadwal + border antar hari + overload
        self.apply_styles(ws, df, slot_str)

        # Sheet Rekap
        self._create_rekap_layanan(wb, df, slot_str)
        self._create_rekap_poli(wb, df, slot_str)
        self._create_rekap_dokter(wb, df, slot_str)
        self._create_grafik_poli(wb)

        # Finishing style
        self._auto_width_all_sheets(wb)
        self._style_headers_all(wb)
        self._freeze_headers_all(wb)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    # --------------------------------------------------------
    # STYLE UNTUK SHEET JADWAL
    # --------------------------------------------------------
    def apply_styles(self, ws, df, slot_str):

        counter = {hari: {s: 0 for s in slot_str} for hari in df["HARI"].unique()}
        records = df.to_dict("records")

        excel_row = 2
        last_hari = None

        # jumlah kolom dinamis
        num_cols = max(ws.max_column, 4 + len(slot_str))

        for rec in records:
            hari = rec.get("HARI")

            # border pemisah antar hari
            if last_hari is not None and hari != last_hari:
                for col in range(1, num_cols + 1):
                    try:
                        ws.cell(row=excel_row, column=col).border = self.border_top_thick
                    except:
                        pass

            last_hari = hari

            # pewarnaan slot
            for idx, slot in enumerate(slot_str):
                v = rec.get(slot, "")
                col_idx = 5 + idx

                try:
                    cell = ws.cell(row=excel_row, column=col_idx)
                except:
                    continue

                if v == "R":
                    cell.fill = self.fill_r

                elif v == "E":
                    counter[hari][slot] += 1

                    if counter[hari][slot] > self.config.max_poleks_per_slot:
                        cell.fill = self.fill_over
                    else:
                        cell.fill = self.fill_e

            excel_row += 1

    # --------------------------------------------------------
    # STYLE PROFESIONAL
    # --------------------------------------------------------
    def _auto_width_all_sheets(self, wb):
        for ws in wb.worksheets:
            for col in ws.columns:
                max_len = 0
                for cell in col:
                    try:
                        max_len = max(max_len, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[col[0].column_letter].width = max_len + 2

    def _style_headers_all(self, wb):
        for ws in wb.worksheets:
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
                cell.border = self.border_header

    def _freeze_headers_all(self, wb):
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
