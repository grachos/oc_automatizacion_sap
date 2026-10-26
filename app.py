"""Desktop UI: pick the Excel, type the header text, create POs in SAP."""
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pythoncom

import config
from excel_io import read_rows, write_results, write_template
from settings_ui import SettingsDialog
from journal import Journal, key
from po_creator import PoCreator
from sap_gui import SapError, connect


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Creación de OCs en SAP - Terceros")
        self.geometry("900x600")
        self.minsize(760, 480)
        self.q: queue.Queue = queue.Queue()
        self.rows = []
        self.file_path = tk.StringVar()
        self.header_text = tk.StringVar()
        self._build()
        self._build_menu()
        self.after(150, self._drain_queue)
        if config.validate():  # first run / incomplete settings: open the dialog right away
            self.after(400, self._open_settings)

    def _build_menu(self):
        menubar = tk.Menu(self)
        m = tk.Menu(menubar, tearoff=0)
        m.add_command(label="Descargar plantilla de Excel...", command=self._download_template)
        m.add_command(label="Configuración...", command=self._open_settings)
        m.add_separator()
        m.add_command(label="Salir", command=self.destroy)
        menubar.add_cascade(label="Menú", menu=m)
        self.config(menu=menubar)

    def _open_settings(self):
        SettingsDialog(self)

    def _download_template(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="plantilla_ocs.xlsx",
                                            filetypes=[("Excel", "*.xlsx")], title="Guardar plantilla")
        if not path:
            return
        try:
            write_template(path)
        except OSError as exc:
            messagebox.showerror("Plantilla", f"No se pudo guardar (¿está abierta en Excel?):\n{exc}")
            return
        messagebox.showinfo("Plantilla", f"Plantilla guardada en:\n{path}")

    # ---------- UI ----------
    def _build(self):
        pad = {"padx": 8, "pady": 4}
        top = ttk.Frame(self)
        top.pack(fill="x", **pad)

        ttk.Label(top, text="Archivo Excel:").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.file_path, state="readonly").grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(top, text="Buscar...", command=self._pick_file).grid(row=0, column=2)

        ttk.Label(top, text="Texto de cabecera:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(top, textvariable=self.header_text).grid(row=1, column=1, columnspan=2, sticky="ew", padx=4, pady=(6, 0))
        top.columnconfigure(1, weight=1)

        cols = ("fila", "nit", "valor", "liquidacion", "oc", "estado", "mensaje")
        widths = (50, 110, 110, 100, 110, 80, 300)
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=12)
        for c, w in zip(cols, widths):
            self.tree.heading(c, text={"oc": "ORDEN DE COMPRA", "liquidacion": "LIQUIDACIÓN"}.get(c, c.upper()))
            self.tree.column(c, width=w, anchor="w" if c == "mensaje" else "center")
        self.tree.tag_configure("OK", background="#C6EFCE")
        self.tree.tag_configure("ERROR", background="#FFC7CE")
        self.tree.tag_configure("REVISAR", background="#FFEB9C")
        self.tree.pack(fill="both", expand=True, **pad)

        bottom = ttk.Frame(self)
        bottom.pack(fill="x", **pad)
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True)
        self.run_btn = ttk.Button(bottom, text="Crear órdenes de compra", command=self._start)
        self.run_btn.pack(side="right", padx=(8, 0))

        self.log_box = tk.Text(self, height=7, state="disabled", bg="#F4F4F4")
        self.log_box.pack(fill="x", **pad)

    def _pick_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        rows, errors = read_rows(path)
        self.file_path.set(path)
        self.rows = rows
        self._apply_journal()
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("", "end", iid=str(r.row), values=(r.row, r.nit, f"{r.amount:,.0f}", r.liquidation, r.po_number, r.status, r.message),
                             tags=(r.status,))
        self._log(f"{len(rows)} filas válidas cargadas. Total: {sum(r.amount for r in rows):,.0f}")
        for e in errors:
            self._log(f"AVISO {e}")

    def _apply_journal(self):
        """Restores progress from a previous (possibly interrupted) run of the same file."""
        states = Journal(self.file_path.get()).last_states()
        for r in self.rows:
            estado, oc, msg = states.get(key(r.nit, r.amount, r.liquidation), ("", "", ""))
            if estado == "OK":
                r.status, r.po_number, r.message = "OK", oc, "Ya creada en una ejecución anterior"
            elif estado == "EN_PROCESO":
                r.status, r.message = "REVISAR", "Se interrumpió al grabar: verifique en SAP si ya existe la OC"

    def _log(self, msg: str):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    # ---------- run ----------
    def _start(self):
        missing = config.validate()
        if missing:
            messagebox.showerror("Configuración", "Faltan valores en .env:\n- " + "\n- ".join(missing))
            return
        if not self.rows:
            messagebox.showwarning("Archivo", "Seleccione un archivo con filas válidas.")
            return
        if not self.header_text.get().strip():
            messagebox.showwarning("Texto", "Digite el texto de cabecera.")
            return
        review = [r for r in self.rows if r.status == "REVISAR"]
        if review:
            lst = "\n".join(f"- Fila {r.row}: NIT {r.nit}, liquidación {r.liquidation}" for r in review)
            if messagebox.askyesno("Filas interrumpidas",
                                   "Estas filas se interrumpieron al grabar y pudieron quedar creadas en SAP:\n"
                                   f"{lst}\n\nVerifique en SAP (ME2N, texto de posición = liquidación).\n"
                                   "¿Reintentarlas de todos modos? (Si elige No, se omiten.)"):
                for r in review:
                    r.status, r.message = "PENDIENTE", ""
        pending = [r for r in self.rows if r.status in ("PENDIENTE", "ERROR")]
        if not pending:
            messagebox.showinfo("Sin pendientes", "No hay filas por crear: todas están creadas o por revisar.")
            return
        if not messagebox.askyesno("Confirmar", f"Se crearán {len(pending)} órdenes de compra en SAP.\n¿Continuar?"):
            return
        self.run_btn.configure(state="disabled")
        self.progress.configure(maximum=len(pending), value=0)
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        pythoncom.CoInitialize()
        log = lambda m: self.q.put(("log", m))
        journal = Journal(self.file_path.get())
        src = self.file_path.get()
        from pathlib import Path
        from datetime import datetime
        out = Path(src).with_name(f"{Path(src).stem}_resultado_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        try:
            session = connect(config.SapSettings(), log)
            saving = {"row": None}
            creator = PoCreator(session, config.PoDefaults(), self.header_text.get().strip(), log,
                                before_save=lambda r: (journal.write(r, "EN_PROCESO"), saving.update(row=r)))
            for r in self.rows:
                if r.status in ("OK", "REVISAR"):
                    continue  # already created / needs manual check
                saving["row"] = None
                try:
                    r.po_number = creator.create(r)
                    r.status, r.message = "OK", "Creada"
                    journal.write(r, "OK", r.po_number)
                    log(f"Fila {r.row} NIT {r.nit}: OC {r.po_number}")
                except Exception as exc:  # keep going with next row
                    if not creator.is_alive():
                        if saving["row"] is r:
                            r.status, r.message = "REVISAR", "Se perdió la conexión al grabar: verifique en SAP si la OC existe"
                        else:
                            r.status, r.message = "ERROR", f"Se perdió la conexión con SAP: {exc}"
                            journal.write(r, "ERROR", "", r.message)
                        self.q.put(("row", r))
                        write_results(self.rows, src, out)
                        log(f"Fila {r.row}: {r.message}")
                        self.q.put(("fatal", "Se perdió la conexión con SAP. Proceso detenido.\n\n"
                                    f"{self._progress_summary()}\n\nAvance guardado en:\n{out}\n{journal.path}\n"
                                    "Al volver la conexión, cargue el mismo Excel y continúe."))
                        return
                    r.status, r.message = "ERROR", str(exc)
                    journal.write(r, "ERROR", "", r.message)
                    log(f"Fila {r.row} NIT {r.nit}: ERROR {exc}")
                self.q.put(("row", r))
                write_results(self.rows, src, out)  # partial result after every row
            log(f"Resultado guardado en: {out}")
            self.q.put(("done", out))
        except SapError as exc:
            self.q.put(("fatal", str(exc)))
        except Exception as exc:
            self.q.put(("fatal", f"Error inesperado: {exc}"))
        finally:
            pythoncom.CoUninitialize()

    def _progress_summary(self) -> str:
        ok = [r for r in self.rows if r.status == "OK"]
        rev = [r for r in self.rows if r.status == "REVISAR"]
        txt = f"Creadas: {len(ok)} de {len(self.rows)}."
        if rev:
            txt += " Revisar en SAP: " + ", ".join(f"fila {r.row} (liquidación {r.liquidation})" for r in rev) + "."
        return txt

    def _value_check(self) -> str:
        """Compares each created PO value (read from SAP) with the Excel amount."""
        created = [r for r in self.rows if r.status == "OK"]
        if not created:
            return "No se crearon OCs: no hay valores para comparar."
        diffs = [r for r in created if r.net_value is not None and abs(r.net_value - r.amount) > 0.5]
        unread = [r for r in created if r.net_value is None]
        sap_total = sum(r.net_value for r in created if r.net_value is not None)
        xls_total = sum(r.amount for r in created if r.net_value is not None)
        lines = [f"Total Excel: {xls_total:,.0f} | Total OCs SAP: {sap_total:,.0f}"]
        if not diffs and not unread:
            lines.insert(0, "✔ El valor de las OCs coincide con el Excel.")
        else:
            lines.insert(0, "⚠ El valor de las OCs NO coincide por completo con el Excel.")
            lines += [f"- Fila {r.row} (OC {r.po_number}): Excel {r.amount:,.0f} vs SAP {r.net_value:,.0f}" for r in diffs]
            lines += [f"- Fila {r.row} (OC {r.po_number}): no se pudo leer el valor en SAP" for r in unread]
        self._log("\n".join(lines))
        return "\n".join(lines)

    def _drain_queue(self):
        try:
            while True:
                kind, data = self.q.get_nowait()
                if kind == "log":
                    self._log(data)
                elif kind == "row":
                    self.tree.item(str(data.row), values=(data.row, data.nit, f"{data.amount:,.0f}", data.liquidation,
                                                          data.po_number, data.status, data.message), tags=(data.status,))
                    self.progress.step(1)
                elif kind == "done":
                    ok = sum(r.status == "OK" for r in self.rows)
                    messagebox.showinfo("Terminado", f"{ok} de {len(self.rows)} OCs creadas.\n\n{self._value_check()}\n\n{data}")
                    self.run_btn.configure(state="normal")
                elif kind == "fatal":
                    self._log(f"ERROR: {data}")
                    messagebox.showerror("SAP", data)
                    self.run_btn.configure(state="normal")
        except queue.Empty:
            pass
        self.after(150, self._drain_queue)


if __name__ == "__main__":
    App().mainloop()
