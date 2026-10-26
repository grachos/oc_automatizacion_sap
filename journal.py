"""Append-only progress log (<excel>_avance.csv), flushed to disk on every write.

Lets the app resume after a power/network failure without duplicating purchase orders.
States: EN_PROCESO (about to save in SAP), OK (PO created), ERROR (nothing saved).
"""
import csv
import os
from datetime import datetime
from pathlib import Path

FIELDS = ["fecha", "fila", "nit", "valor", "liquidacion", "estado", "oc", "mensaje"]


def key(nit, amount, liquidation) -> str:
    return f"{nit}|{float(amount):.2f}|{liquidation}"


class Journal:
    def __init__(self, excel_path: str | Path):
        p = Path(excel_path)
        self.path = p.with_name(f"{p.stem}_avance.csv")

    def last_states(self) -> dict:
        """key -> (estado, oc, mensaje) of the latest entry per row."""
        states = {}
        if not self.path.exists():
            return states
        with open(self.path, newline="", encoding="utf-8-sig") as f:
            for rec in csv.DictReader(f):
                states[key(rec["nit"], rec["valor"], rec["liquidacion"])] = (rec["estado"], rec["oc"], rec["mensaje"])
        return states

    def write(self, row, estado: str, oc: str = "", mensaje: str = ""):
        new = not self.path.exists()
        with open(self.path, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            if new:
                w.writerow(FIELDS)
            w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), row.row, row.nit, f"{row.amount:.2f}",
                        row.liquidation, estado, oc, mensaje])
            f.flush()
            os.fsync(f.fileno())
