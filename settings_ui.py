"""Settings dialog: edits the .env values (SAP login and PO defaults) from the UI."""
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import config

SAP_FIELDS = [
    ("SAP_USER", "Usuario SAP", False),
    ("SAP_PASSWORD", "Contraseña SAP", True),
    ("SAP_CONNECTION", "Conexión (como aparece en SAP Logon)", False),
    ("SAP_CLIENT", "Mandante", False),
    ("SAP_LANGUAGE", "Idioma", False),
    ("SAP_LOGON_PATH", "Ruta de saplogon.exe", False),
]
PO_FIELDS = [
    ("PO_PURCH_ORG", "Organización de compras", False),
    ("PO_PURCH_GROUP", "Grupo de compras", False),
    ("PO_COMPANY_CODE", "Sociedad", False),
    ("PO_PLANT", "Centro", False),
    ("PO_PAYMENT_DAYS", "Pago en (días)", False),
    ("PO_TAX_CODE", "Indicador de IVA", False),
    ("PO_MATERIAL_GROUP", "Grupo de artículos", False),
    ("PO_SERVICE_NUMBER", "Nº de servicio", False),
    ("PO_GL_ACCOUNT", "Cuenta de mayor", False),
]


class SettingsDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Configuración")
        self.resizable(False, False)
        self.transient(parent)
        self.vars: dict[str, tk.StringVar] = {}
        s, p = config.SapSettings(), config.PoDefaults()
        current = {"SAP_USER": s.user, "SAP_PASSWORD": s.password, "SAP_CONNECTION": s.connection,
                   "SAP_CLIENT": s.client, "SAP_LANGUAGE": s.language, "SAP_LOGON_PATH": s.logon_path,
                   "PO_PURCH_ORG": p.purch_org, "PO_PURCH_GROUP": p.purch_group, "PO_COMPANY_CODE": p.company_code,
                   "PO_PLANT": p.plant, "PO_PAYMENT_DAYS": p.payment_days, "PO_TAX_CODE": p.tax_code,
                   "PO_MATERIAL_GROUP": p.material_group, "PO_SERVICE_NUMBER": p.service_number,
                   "PO_GL_ACCOUNT": p.gl_account}
        for title, fields in (("Acceso a SAP", SAP_FIELDS), ("Parámetros de la orden de compra", PO_FIELDS)):
            box = ttk.LabelFrame(self, text=title)
            box.pack(fill="x", padx=10, pady=6)
            for i, (key, label, secret) in enumerate(fields):
                ttk.Label(box, text=label + ":").grid(row=i, column=0, sticky="w", padx=6, pady=3)
                var = tk.StringVar(value=current[key])
                self.vars[key] = var
                ttk.Entry(box, textvariable=var, width=48, show="*" if secret else "").grid(row=i, column=1, padx=6, pady=3)
        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=10, pady=8)
        ttk.Button(btns, text="Guardar", command=self._save).pack(side="right")
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side="right", padx=6)
        self.grab_set()

    def _save(self):
        values = {k: v.get().strip() for k, v in self.vars.items()}
        required = {"SAP_USER": "Usuario", "SAP_PASSWORD": "Contraseña", "SAP_CONNECTION": "Conexión",
                    "SAP_CLIENT": "Mandante", "PO_PLANT": "Centro"}
        missing = [name for k, name in required.items() if not values[k]]
        if missing:
            messagebox.showwarning("Configuración", "Complete: " + ", ".join(missing), parent=self)
            return
        if not Path(values["SAP_LOGON_PATH"]).exists():
            if not messagebox.askyesno("Configuración", "La ruta de saplogon.exe no existe.\n¿Guardar de todos modos?", parent=self):
                return
        config.save_env(values)
        messagebox.showinfo("Configuración", "Configuración guardada.", parent=self)
        self.destroy()
