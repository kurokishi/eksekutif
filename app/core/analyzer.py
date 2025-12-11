# app/core/analyzer.py
import pandas as pd
import re

class ErrorAnalyzer:
    """Sederhana: analisa struktur sheet Reguler/Poleks"""

    @staticmethod
    def analyze_sheet(df: pd.DataFrame, hari_list: list):
        report = {
            'is_valid': True,
            'errors': [],
            'warnings': [],
            'total_rows': len(df)
        }

        # cek kolom wajib
        required = ['Nama Dokter', 'Poli Asal', 'Jenis Poli']
        missing = [c for c in required if c not in df.columns]
        if missing:
            report['is_valid'] = False
            report['errors'].append(f"Missing required columns: {missing}")

        # cek format waktu sederhana
        pattern = re.compile(r'^\d{1,2}[:\.]\d{2}\s*-\s*\d{1,2}[:\.]\d{2}$')
        cols = [h for h in hari_list if h in df.columns]
        bad = 0
        for col in cols:
            for v in df[col].dropna().astype(str).values:
                if not pattern.search(v.strip()):
                    bad += 1
        if bad:
            report['warnings'].append(f"{bad} cell(s) with non-standard time format")
            # tidak langsung menjadikan invalid, karena ada auto-fix
        return report

    @staticmethod
    def format_report(rep: dict) -> str:
        s = f"Valid: {'✅' if rep.get('is_valid') else '❌'}\n"
        s += f"Total Rows: {rep.get('total_rows',0)}\n"
        if rep.get('errors'):
            s += "Errors:\n" + "\n".join(f"- {e}" for e in rep['errors']) + "\n"
        if rep.get('warnings'):
            s += "Warnings:\n" + "\n".join(f"- {w}" for w in rep['warnings']) + "\n"
        return s
