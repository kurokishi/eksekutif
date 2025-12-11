from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
import io
import pandas as pd
from datetime import datetime, timedelta


class ExcelWriter:

    def __init__(self, config):
        self.config = config

        # warna slot
        self.fill_r = PatternFill(start_color="00FF00", fill_type="solid")
        self.fill_e = PatternFill(start_color="0000FF", fill_type="solid")
        self.fill_over = PatternFill(start_color="FF0000", fill_type="solid")

        # border thicc antar hari
        # self.border_top_thick = Border(top=Side(border_style="thick"))
        # self.border_header = Border(bottom=Side(border_style="thick"))

    # =====================================================
    # RANGE BUILDER UNTUK REKAP
    # =====================================================
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

    # =====================================================
    # SHEET: REKAP LAYANAN (PER DOKTER)
    # =====================================================
    def _create_rekap_layanan(self, wb, df, slot_str):
        if "Rekap Layanan" in wb.sheetnames:
            del wb["Rekap Layanan"]

        ws = wb.create_sheet("Rekap Layanan")
        ws.append(["POLI ASAL", "HARI", "DOKTER", "JENIS", "JAM LAYANAN"])

        interval = self.config.interval_minutes

        for (poli, hari, dokter), g in df.groupby(["POLI ASAL", "HARI", "DOKTER"]):
            r_slots = [s for s in slot_str if g.iloc[0].get(s, "") == "R"]
            e_slots = [s for s in slot_str if g.iloc[0].get(s, "") == "E"]

            r_ranges = self._combine_ranges(r_slots, interval)
            e_ranges = self._combine_ranges(e_slots, interval)

            for a, b in r_ranges:
                ws.append([poli, hari, dokter, "Reguler", self._format_range(a, b)])

            for a, b in e_ranges:
                ws.append([poli, hari, dokter, "Poleks", self._format_range(a, b)])

    # =====================================================
    # SHEET: REKAP PER POLI
    # =====================================================
    def _create_rekap_poli(self, wb, df, slot_str):
        if "Rekap Poli" in wb.sheetnames:
            del wb["Rekap Poli"]

        ws = wb.create_sheet("Rekap Poli")
        ws.append(["POLI", "HARI", "TOTAL JAM REGULER", "TOTAL JAM POLEKS", "TOTAL JAM LAYANAN"])

        interval = self.config.interval_minutes

        for (poli, hari), g in df.groupby(["POLI ASAL", "HARI"]):
            total_r = 0
            total_e = 0

            for slot in slot_str:
                v = g.iloc[0].get(slot, "")
                if v == "R":
                    total_r += interval / 60
                elif v == "E":
                    total_e += interval / 60

            ws.append([
                poli,
                hari,
                round(total_r, 2),
                round(total_e, 2),
                round(total_r + total_e, 2)
            ])

    # =====================================================
    # SHEET: TOTAL JAM DOKTER PER HARI
    # =====================================================
    def _create_rekap_dokter(self, wb, df, slot_str):
        if "Rekap Dokter" in wb.sheetnames:
            del wb["Rekap Dokter"]

        ws = wb.create_sheet("Rekap Dokter")
        ws.append(["DOKTER", "HARI", "TOTAL JAM REGULER", "TOTAL JAM POLEKS", "TOTAL JAM"])

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

            ws.append([dokter, hari, round(tot_r, 2), round(tot_e, 2), round(tot_r + tot_e, 2)])

    # =====================================================
    # SHEET: GRAFIK BEBAN POLI
    # =====================================================
    def _create_grafik_poli(self, wb):
        if "Grafik Beban Poli" in wb.sheetnames:
            del wb["Grafik Beban Poli"]

        ws = wb.create_sheet("Grafik Beban Poli")

        # Ambil data dari sheet Rekap Poli
        rp = wb["Rekap Poli"]
        ws["A1"] = "Grafik Beban Poli (Total Jam Layanan per Minggu)"

        # copy hasil agregasi per poli
        table = {}
        for row in rp.iter_rows(min_row=2, values_only=True):
            poli = row[0]
            total = row[4]  # total jam layanan

            table[poli] = table.get(poli, 0) + total

        ws.append(["POLI", "TOTAL JAM"])
        for k, v in table.items():
            ws.append([k, v])

        # buat chart
        chart = BarChart()
        chart.title = "Beban Poli (Jam Layanan)"

        data = Reference(ws, min_col=2, min_row=2, max_row=ws.max_row)
        cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)

        chart.add_data(data)
        chart.set_categories(cats)
        chart.y_axis.title = "Jam"

        ws.add_chart(chart, "E5")

    # =====================================================
    # SHEET: JADWAL UTAMA
    # =====================================================
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

        # === tambahkan seluruh sheet rekap ===
        self._create_rekap_layanan(wb, df, slot_str)
        self._create_rekap_poli(wb, df, slot_str)
        self._create_rekap_dokter(wb, df, slot_str)
        self._create_grafik_poli(wb)

        # FORMAT PROFESIONAL
        self._auto_width_all_sheets(wb)
        self._style_headers_all(wb)
        self._freeze_headers_all(wb)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    # =====================================================
    # STYLE PADA JADWAL
    # =====================================================
    def apply_styles(self, ws, df, slot_str):
        counter = {hari: {s: 0 for s in slot_str} for hari in df["HARI"].unique()}
        records = df.to_dict("records")

        excel_row = 2
        last_hari = None

        for rec in records:
            hari = rec["HARI"]

            # border antar hari
            if last_hari is not None and hari != last_hari:
                for col in range(1, 5 + len(slot_str)):
                    ws.cell(row=excel_row, column=col).border = self.border_top_thick

            last_hari = hari

            for slot in slot_str:
                v = rec.get(slot, "")
                cell = ws.cell(row=excel_row, column=slot_str.index(slot) + 5)

                if v == "R":
                    cell.fill = self.fill_r
                elif v == "E":
                    counter[hari][slot] += 1
                    if counter[hari][slot] > self.config.max_poleks_per_slot:
                        cell.fill = self.fill_over
                    else:
                        cell.fill = self.fill_e

            excel_row += 1

    # =====================================================
    # PROFESSIONAL GLOBAL STYLING
    # =====================================================
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
            for col in ws.iter_cols(min_row=1, max_row=1):
                for cell in col:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal="center")
                    cell.border = self.border_header

    def _freeze_headers_all(self, wb):
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
