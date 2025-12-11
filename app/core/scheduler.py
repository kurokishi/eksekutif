# app/core/scheduler.py

import pandas as pd
from datetime import datetime, timedelta, time
from .time_parser import TimeParser
from .cleaner import DataCleaner

class Scheduler:

    def __init__(self, config):
        self.config = config

    def generate_time_slots(self):
        slots = []
        current = time(self.config.start_hour, self.config.start_minute)
        end = time(14, 30)

        while current <= end:
            slots.append(current)
            dt = datetime.combine(datetime.today(), current) + timedelta(minutes=self.config.interval_minutes)
            current = dt.time()

        return slots

    def process(self, df, jenis):
        df = DataCleaner.clean(df, self.config.hari_list, jenis)
        tp = TimeParser()

        results = []
        slots = self.generate_time_slots()
        slots_str = [t.strftime("%H:%M") for t in slots]

        for (dokter, poli), grp in df.groupby(['Nama Dokter', 'Poli Asal']):
            for hari in self.config.hari_list:

                if hari not in grp.columns:
                    continue

                ranges = []
                for s in grp[hari]:
                    st, en = tp.parse(s)
                    if st and en:
                        if en > time(14, 30):
                            en = time(14, 30)
                        ranges.append((st, en))

                merged = self.merge_ranges(ranges)
                row = self.create_row(poli, jenis, hari, dokter, merged, slots, slots_str)
                results.append(row)

        return pd.DataFrame(results)

    def merge_ranges(self, ranges):
        if not ranges:
            return []
        ranges.sort()
        merged = [list(ranges[0])]
        for start, end in ranges[1:]:
            ls, le = merged[-1]
            if start <= le:
                merged[-1][1] = max(le, end)
            else:
                merged.append([start, end])
        return merged

    def create_row(self, poli, jenis, hari, dokter, merged, slots, slots_str):
        row = {
            'POLI ASAL': poli,
            'JENIS POLI': jenis,
            'HARI': hari,
            'DOKTER': dokter
        }
        for i, slot in enumerate(slots):
            end = (datetime.combine(datetime.today(), slot) + timedelta(minutes=30)).time()
            overlap = any(not (end <= st or slot >= en) for st, en in merged)
            row[slots_str[i]] = 'R' if overlap and jenis == 'Reguler' else 'E' if overlap else ''
        return row
