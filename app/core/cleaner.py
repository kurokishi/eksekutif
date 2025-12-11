# app/core/cleaner.py

import pandas as pd
import re

class DataCleaner:

    @staticmethod
    def clean(df, hari_list, jenis_poli):
        df = df.copy()

        required = ['Nama Dokter', 'Poli Asal', 'Jenis Poli']
        for col in required:
            if col not in df.columns:
                df[col] = ''

        df['Jenis Poli'] = df['Jenis Poli'].fillna(jenis_poli)

        for hari in hari_list:
            if hari in df.columns:
                df[hari] = df[hari].apply(DataCleaner.fix_time)

        if any(h in df.columns for h in hari_list):
            df = df[df[hari_list].notna().any(axis=1)]

        return df

    @staticmethod
    def fix_time(time_str):
        if pd.isna(time_str):
            return ""
        s = re.sub(r'[^\d.\-:]', '', str(time_str).strip())
        if '.' in s and ':' not in s:
            parts = s.split('-')
            if len(parts) == 2:
                return f"{parts[0].replace('.',':')}-{parts[1].replace('.',':')}"
        return s
