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

        # border header
        self.border_header = Border(bottom=Side(border_style="thick"))

    # --------------------------------------------------------
    # Helper: gabung rentang waktu slot per interval
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
    # SHEET: PEAK HOUR ANALYSIS
    # --------------------------------------------------------
    def _create_peak_hour(self, wb, df, slot_str):
        if "Peak Hour Analysis" in wb.sheetnames:
            del wb["Peak Hour Analysis"]

        ws = wb.create_sheet("Peak Hour Analysis")
        ws.append(["HARI", "SLOT", "JUMLAH", "KATEGORI"])

        for hari, g in df.groupby("HARI"):
            slot_count = {s: 0 for s in slot_str}

            for _, row in g.iterrows():
                for s in slot_str:
                    if row.get(s) in ["R", "E"]:
                        slot_count[s] += 1

            max_val = max(slot_count.values())
            peak_slots = [s for s, v in slot_count.items() if v == max_val]

            kategori = "High Load" if max_val >= 10 else "Medium" if max_val >= 5 else "Low"

            for s in peak_slots:
                ws.append([hari, s, max_val, kategori])

    # --------------------------------------------------------
    # SHEET: CONFLICT CHECK DOCTOR
    # --------------------------------------------------------
    def _create_conflict_doctor(self, wb, df, slot_str):
        if "Conflict Dokter" in wb.sheetnames:
            del wb["Conflict Dokter"]

        ws = wb.create_sheet("Conflict Dokter")
        ws.append(["DOKTER", "HARI", "SLOT", "KONFLIK"])

        for (dokter, hari), g in df.groupby(["DOKTER", "HARI"]):
            for slot in slot_str:
                vals = g[slot].unique()

                # Konflik: dokter muncul di lebih dari 1 poli
                if len(vals) > 1 and any(v in ["R", "E"] for v in vals):
                    ws.append([dokter, hari, slot, "Dokter memiliki 2 poli berbeda pada jam yang sama"])

                # Konflik: R dan E di slot sama
                if "R" in vals and "E" in vals:
                    ws.append([dokter, hari, slot, "Bentrok Reguler & Poleks"])

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
        ws.append(["POLI", "HARI", "TOTAL JAM REG", "TOTAL JAM POLEKS", "TOTAL JAM"])

        interval = self.config.interval_minutes

        for (poli, hari), g in df.groupby(["POLI ASAL", "HARI"]):
            tot_r = sum((g.iloc[0].get(s) == "R") * interval/60 for s in slot_str)
            tot_e = sum((g.iloc[0].get(s) == "E") * interval/60 for s in slot_str)
            ws.append([poli, hari, round(tot_r, 2), round(tot_e, 2), round(tot_r+tot_e, 2)])

    # --------------------------------------------------------
    # SHEET: Rekap Dokter
    # --------------------------------------------------------
    def _create_rekap_dokter(self, wb, df, slot_str):
        if "Rekap Dokter" in wb.sheetnames:
            del wb["Rekap Dokter"]

        ws = wb.create_sheet("Rekap Dokter")
        ws.append(["DOKTER", "HARI", "TOTAL JAM REG", "TOTAL JAM POLEKS", "TOTAL JAM"])

        interval = self.config.interval_minutes

        for (dokter, hari), g in df.groupby(["DOKTER", "HARI"]):
            tot_r = sum((g.iloc[0].get(s) == "R") * interval/60 for s in slot_str)
            tot_e = sum((g.iloc[0].get(s) == "E") * interval/60 for s in slot_str)
            ws.append([dokter, hari, round(tot_r, 2), round(tot_e, 2), round(tot_r+tot_e, 2)])

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
        for p, t in table.items():
            ws.append([p, t])

        chart = BarChart()
        chart.title = "Beban Poli"
        data = Reference(ws, min_col=2, min_row=2, max_row=ws.max_row)
        cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
        chart.add_data(data)
        chart.set_categories(cats)
        ws.add_chart(chart, "E5")

    # --------------------------------------------------------
    # SHEET UTAMA: Jadwal (tanpa border antar hari)
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

        # pewarnaan slot (tanpa border antar hari)
        self.apply_styles(ws, df, slot_str)

        # Sheet Rekap
        self._create_rekap_layanan(wb, df, slot_str)
        self._create_rekap_poli(wb, df, slot_str)
        self._create_rekap_dokter(wb, df, slot_str)
        self._create_peak_hour(wb, df, slot_str)
        self._create_conflict_doctor(wb, df, slot_str)
        self._create_grafik_poli(wb)

        # finishing
        self._auto_width_all_sheets(wb)
        self._style_headers_all(wb)
        self._freeze_headers_all(wb)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    # --------------------------------------------------------
    # Pewarnaan slot (tanpa border antar hari)
    # --------------------------------------------------------
    def apply_styles(self, ws, df, slot_str):

        counter = {hari: {s: 0 for s in slot_str} for hari in df["HARI"].unique()}
        records = df.to_dict("records")

        excel_row = 2

        for rec in records:
            hari = rec.get("HARI")

            for idx, slot in enumerate(slot_str):
                v = rec.get(slot, "")
                col_idx = 5 + idx

                cell = ws.cell(row=excel_row, column=col_idx)

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
    # Styling profesional
    # --------------------------------------------------------
    def _auto_width_all_sheets(self, wb):
        for ws in wb.worksheets:
            for col in ws.columns:
                width = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                ws.column_dimensions[col[0].column_letter].width = width + 2

    def _style_headers_all(self, wb):
        for ws in wb.worksheets:
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
                cell.border = self.border_header

    def _freeze_headers_all(self, wb):
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
