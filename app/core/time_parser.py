# app/core/time_parser.py

import re
import pandas as pd
from datetime import time

class TimeParser:

    @staticmethod
    def parse(time_str):
        if pd.isna(time_str) or str(time_str).strip() == "":
            return None, None

        s = str(time_str).strip().replace(' ', '').replace('.', ':')

        match = re.search(r'(\d{1,2}:\d{2})-(\d{1,2}:\d{2})', s)
        if not match:
            return None, None
        
        start_s, end_s = match.groups()
        sh, sm = map(int, start_s.split(':'))
        eh, em = map(int, end_s.split(':'))

        return time(sh, sm), time(eh, em)
