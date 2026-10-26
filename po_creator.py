"""Creates one purchase order in ME21N per Excel row.

Field names/types come from a real SAP GUI recording (grabacion_me21n.vbs, QAS).
Dynpro numbers (SAPLMEGUI:00NN) change between screens, so controls are located with
SAP's native FindByName filtered by input type (labels share the same technical name).
"""
import re
from datetime import date, timedelta

import pythoncom

from config import PoDefaults
from excel_io import PoRow
from sap_gui import SapError, call, find_all, find_by_name, find_editor, find_input, find_tab, put_text, rel_id, status, typed

HDR_DELIVERY, HDR_ORG, HDR_TEXTS = "TABHDT1", "TABHDT9", "TABHDT3"
ITEM_SERVICES, ITEM_TEXTS = "TABIDT1", "TABIDT15"
HDR_EDITOR, ITEM_EDITOR = "TEXT_EDITOR_0101", "TEXT_EDITOR_0201"
TOGGLE_BUTTON = "DYN_4000-BUTTON"
SRV_TABLE = "SAPLMLSPTC_VIEW"
DELIVERY_DAYS = 15  # Fe.entrega = today + 15 days


def _com_detail(exc: pythoncom.com_error) -> str:
    return exc.excepinfo[2] if exc.excepinfo and exc.excepinfo[2] else str(exc.strerror)


class PoCreator:
    def __init__(self, session, defaults: PoDefaults, header_text: str, log=print, before_save=None):
        self.s = session
        self.before_save = before_save  # called right before pressing Save
        self.d = defaults
        self.header_text = header_text
        self.log = log

    # ---------- helpers ----------
    @property
    def usr(self):
        return self.s.findById("wnd[0]/usr")

    def _enter(self):
        self.s.findById("wnd[0]").sendVKey(0)  # synchronous: returns after SAP round trip

    def _set(self, name: str, value: str, container=None):
        if not value:
            return
        try:
            self._set_inner(name, value, container)
        except pythoncom.com_error as exc:
            raise SapError(f"Campo {name}: {_com_detail(exc)}") from exc

    def _clear(self, name: str):
        try:
            ctrl = find_input(self.usr, name)
            if ctrl is not None and (ctrl.Text or "").strip() and ctrl.Changeable:
                put_text(ctrl, "")
        except pythoncom.com_error as exc:
            raise SapError(f"Campo {name} (limpiar): {_com_detail(exc)}") from exc

    def _set_inner(self, name: str, value: str, container=None):
        ctrl = find_input(container or self.usr, name)
        if ctrl is None:
            raise SapError(f"Campo '{name}' no encontrado en pantalla.")
        if (ctrl.Text or "").strip() == value:
            return
        if not ctrl.Changeable:
            self.log(f"  aviso: {name} no editable (valor actual '{ctrl.Text.strip()}'), se omite")
            return
        try:
            put_text(ctrl, value)
        except pythoncom.com_error as exc:
            raise SapError(f"Campo {name} = '{value}' rechazado por SAP: {_com_detail(exc)}") from exc

    def _set_cell(self, table: str, cell: str, value: str):
        """Service table cell by exact recorded id, e.g. table SAPLMLSPTC_VIEW, cell 'txtESLL-MENGE[4,0]'."""
        name = cell.split("ESLL-")[0] and "ESLL-" + cell.split("ESLL-")[1].split("[")[0]
        try:
            tbl = find_by_name(self.usr, table, "GuiTableControl")
            if tbl is None:
                raise SapError(f"Tabla {table} no encontrada.")
            try:
                ctrl = self.s.findById(rel_id(tbl) + "/" + cell)
            except pythoncom.com_error:  # column scrolled out of view -> scroll right and retry
                col = int(cell.split("[")[1].split(",")[0])
                tbl.HorizontalScrollbar.Position = max(0, col - 4)
                tbl = find_by_name(self.usr, table, "GuiTableControl")
                ctrl = self.s.findById(rel_id(tbl) + "/" + cell)
            if (ctrl.Text or "").strip() == value:
                return
            put_text(ctrl, value)
        except pythoncom.com_error as exc:
            raise SapError(f"Campo {name}: {_com_detail(exc)}") from exc

    def _expand(self, area: str):
        """Presses the expand button of one area only. Ids from the recording:
        header = .../subSUB1:SAPLMEVIEWS:1100/..., item detail = .../subSUB3:SAPLMEVIEWS:1100/...
        Buttons are re-queried after every round trip (old references go stale)."""
        for btn in find_all(self.usr, TOGGLE_BUTTON, "GuiButton"):
            if f"/{area}:SAPLMEVIEWS:1100/" in btn.Id:
                call(btn, "Press")
                return

    def _select_tab(self, tab_name: str):
        tab = find_tab(self.s, tab_name)
        if tab is None:
            self._expand("subSUB1" if tab_name.startswith("TABHDT") else "subSUB3")
            tab = find_tab(self.s, tab_name)
        if tab is None:
            raise SapError(f"Pestaña '{tab_name}' no encontrada.")
        call(tab, "Select")
        return find_tab(self.s, tab_name)

    def _set_editor(self, tab_name: str, editor_name: str, value: str, label: str):
        tab = self._select_tab(tab_name)
        editor = find_editor(tab, editor_name)
        if editor is None:
            raise SapError(f"Editor de {label} no encontrado.")
        try:
            put_text(editor, value + "\r")
        except pythoncom.com_error as exc:
            raise SapError(f"{label} rechazado por SAP: {_com_detail(exc)}") from exc

    def _handle_popups(self):
        """Service account-assignment popup (ESKN-SAKTO) and warnings -> Enter."""
        for _ in range(6):
            count = self.s.Children.Count
            if count < 2:
                return
            popup = self.s.findById(f"wnd[{count - 1}]")
            gl = find_input(popup, "ESKN-SAKTO")
            if gl is not None and gl.Changeable and not (gl.Text or "").strip():
                put_text(gl, self.d.gl_account)
            popup.sendVKey(0)

    def _raise_if_error(self, step: str):
        mtype, text = status(self.s)
        if mtype in ("E", "A"):
            raise SapError(f"{step}: {text}")

    # ---------- flow ----------
    def create(self, row: PoRow) -> str:
        self.s.findById("wnd[0]/tbar[0]/okcd").text = "/nME21N"
        self._enter()
        self._handle_popups()

        self._set("MEPO_TOPLINE-SUPERFIELD", row.nit)
        self._enter()  # SAP replies "Introduzca Org.compras": expected, org data is filled next

        self._select_tab(HDR_ORG)
        self._set("MEPO1222-EKORG", self.d.purch_org)
        self._set("MEPO1222-EKGRP", self.d.purch_group)
        self._set("MEPO1222-BUKRS", self.d.company_code)
        self._enter()

        # Payment: "Cond.pago" must stay BLANK and "Pago en" = 7 days. Done AFTER org data,
        # because the org Enter reloads the vendor's default payment terms (e.g. V030).
        self._select_tab(HDR_DELIVERY)
        self._clear("MEPO1226-ZTERM")
        self._set("MEPO1226-ZBD1T", self.d.payment_days)
        self._enter()

        self._set("MEPO1211-KNTTP", self.d.acct_assignment)
        self._set("MEPO1211-EPSTP", self.d.item_category)
        self._set("MEPO1211-TXZ01", self.d.short_text)
        self._set("MEPO1211-MENGE", self.d.quantity)
        self._set("MEPO1211-MEINS", self.d.unit)
        self._set("MEPO1211-EEIND", (date.today() + timedelta(days=DELIVERY_DAYS)).strftime("%d.%m.%Y"))  # Fe.entrega
        self._enter()  # SAP asks for the plant here; expected
        self._set("MEPO1211-NAME1", self.d.plant)
        self._set("MEPO1211-WGBEZ", self.d.material_group)
        self._enter()

        self._select_tab(ITEM_SERVICES)
        self._set("ESLL-SRVPOS", self.d.service_number)
        self._enter()
        self._set_cell(SRV_TABLE, "txtESLL-MENGE[4,0]", self.d.service_qty)
        amount = f"{row.amount:.0f}" if row.amount.is_integer() else f"{row.amount:.2f}".replace(".", ",")
        self._set_cell(SRV_TABLE, "txtESLL-TBTWR[6,0]", amount)
        self._enter()
        self._set_cell(SRV_TABLE, "ctxtESLL-MWSKZ[16,0]", self.d.tax_code)
        self._enter()          # opens "Imputación del servicio" popup
        self._handle_popups()  # G/L account + warning

        self._set_editor(HDR_TEXTS, HDR_EDITOR, self.header_text, "texto de cabecera")
        self._set_editor(ITEM_TEXTS, ITEM_EDITOR, row.liquidation, "texto de posición")
        row.net_value = self._read_net_value()
        if self.before_save:
            self.before_save(row)
        return self._save()

    def is_alive(self) -> bool:
        """False when the SAP session/connection is gone (network or SAP down)."""
        try:
            self.s.findById("wnd[0]")
            return True
        except Exception:
            return False

    def _read_net_value(self):
        """Net price of the item as shown in SAP (None if it cannot be read)."""
        try:
            ctrl = find_input(self.usr, "MEPO1211-NETPR")
            txt = (ctrl.Text or "").strip() if ctrl is not None else ""
            if not txt:
                return None
            return float(txt.replace(".", "").replace(",", ".")) if "," in txt else float(txt.replace(".", ""))
        except (pythoncom.com_error, ValueError):
            return None

    def _save(self) -> str:
        call(self.s.findById("wnd[0]/tbar[0]/btn[11]"), "Press")
        for _ in range(3):  # "save?" / warnings
            count = self.s.Children.Count
            if count < 2:
                break
            popup = self.s.findById(f"wnd[{count - 1}]")
            yes = find_by_name(popup, "SPOP-OPTION1", "GuiButton")
            if yes is not None:
                call(typed(yes), "Press")
            else:
                popup.sendVKey(0)
        mtype, text = status(self.s)
        match = re.search(r"\b(\d{8,10})\b", text)
        if mtype in ("E", "A") or not match:
            detail = text or "SAP no mostró mensaje"
            count = self.s.Children.Count
            if count > 1:  # a popup is still open: report its title
                detail += f" | ventana abierta: {self.s.findById(f'wnd[{count - 1}]').Text}"
            raise SapError(f"No se pudo grabar: {detail}")
        return match.group(1)
