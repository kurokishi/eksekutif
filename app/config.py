# app/config.py

from datetime import time

class Config:
    def __init__(self):
        # Waktu dasar
        self.start_hour = 7
        self.start_minute = 30
        self.interval_minutes = 30
        self.max_poleks_per_slot = 7

        # Mode Sabtu (default: off)
        self.enable_sabtu = False

        # Hari default
        self.hari_order = {
            "Senin": 1,
            "Selasa": 2,
            "Rabu": 3,
            "Kamis": 4,
            "Jum'at": 5
        }

    @property
    def hari_list(self):
        hari = list(self.hari_order.keys())
        if self.enable_sabtu and "Sabtu" not in hari:
            hari.append("Sabtu")
        return hari
