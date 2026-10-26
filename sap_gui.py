"""SAP GUI Scripting helpers: launch SAP Logon, log in, find controls by name."""
import subprocess
import time

import pythoncom
import win32com.client

from config import SapSettings


class SapError(Exception):
    pass


def _get_scripting_engine(logon_path: str, timeout: int = 60):
    """Returns the SAPGUI scripting engine, launching saplogon.exe if needed."""
    deadline = time.time() + timeout
    launched = False
    while time.time() < deadline:
        try:
            return win32com.client.GetObject("SAPGUI").GetScriptingEngine
        except pythoncom.com_error:
            if not launched:
                subprocess.Popen([logon_path])
                launched = True
            time.sleep(2)
    raise SapError("No se pudo iniciar SAP Logon o el scripting está deshabilitado.")


def connect(settings: SapSettings, log=print):
    """Opens the connection, logs in and returns a GuiSession."""
    app = _get_scripting_engine(settings.logon_path)
    log(f"Abriendo conexión '{settings.connection}'...")
    try:
        conn = app.OpenConnection(settings.connection, True)
    except pythoncom.com_error as exc:
        raise SapError(f"No existe la conexión '{settings.connection}' en SAP Logon.") from exc
    session = _fresh_session(conn)
    wnd = session.findById("wnd[0]")
    wnd.findById("usr/txtRSYST-MANDT").text = settings.client
    wnd.findById("usr/txtRSYST-BNAME").text = settings.user
    wnd.findById("usr/pwdRSYST-BCODE").text = settings.password
    wnd.findById("usr/txtRSYST-LANGU").text = settings.language
    wnd.sendVKey(0)
    # After logon SAP may rebuild the session object (old references "disconnect"),
    # so always re-acquire it from the connection before using it again.
    session = _handle_logon_popups(conn)
    sbar = session.findById("wnd[0]/sbar")
    if sbar.MessageType in ("E", "A"):
        raise SapError(f"Login fallido: {sbar.Text}")
    log(f"Sesión iniciada como {settings.user}.")
    return session


def _fresh_session(conn, timeout: float = 30):
    """Returns a live session of the connection, waiting while SAP rebuilds it."""
    global _session
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            session = conn.Children(0)
            session.findById("wnd[0]")
            _session = session
            return session
        except pythoncom.com_error:
            time.sleep(0.5)
    raise SapError("La sesión de SAP no respondió después del login.")


def _handle_logon_popups(conn):
    """Closes post-logon popups: multiple logon, 'documentos no leídos', system messages...
    Returns a fresh session reference."""
    for _ in range(8):
        session = _fresh_session(conn)
        try:
            if session.Children.Count < 2:
                return session
            popup = session.findById(f"wnd[{session.Children.Count - 1}]")
            multi = find_by_name(popup, "MULTI_LOGON_OPT2", "GuiRadioButton")
            if multi is not None:  # "Continue with this logon, without ending others"
                session.findById(multi.Id.split("/ses[", 1)[1].split("/", 1)[1]).select()
            popup.sendVKey(0)  # Enter = green check (e.g. "Tiene documentos no leídos")
        except pythoncom.com_error:
            pass  # session rebuilt mid-step: loop re-acquires it
        time.sleep(0.5)
    return _fresh_session(conn)


def find_by_name(container, name: str, type_: str = None):
    try:
        return container.FindByName(name, type_) if type_ else container.FindByName(name)
    except pythoncom.com_error:
        return None


_session = None


def rel_id(ctrl) -> str:
    """'/app/con[0]/ses[0]/wnd[0]/usr/...' -> 'wnd[0]/usr/...'"""
    return ctrl.Id.split("/ses[", 1)[1].split("/", 1)[1]


def typed(ctrl):
    """Re-fetches a control through session.findById so pywin32 sees its full interface."""
    return _session.findById(rel_id(ctrl))


def call(ctrl, method: str, *args):
    """Calls a SAP GUI method with its exact name (Press, Select...) via raw COM."""
    dispid = ctrl._oleobj_.GetIDsOfNames(method)
    return ctrl._oleobj_.Invoke(dispid, 0, pythoncom.INVOKE_FUNC, 1, *args)


def put_text(ctrl, value: str):
    """Sets the Text property via raw COM so SAP's real error message is surfaced."""
    dispid = ctrl._oleobj_.GetIDsOfNames("Text")
    ctrl._oleobj_.Invoke(dispid, 0, pythoncom.INVOKE_PROPERTYPUT, 0, value)


def find_all(container, name: str, type_: str) -> list:
    """All controls with that name AND type (native SAP search, no tree walking)."""
    try:
        coll = container.FindAllByName(name, type_)
    except pythoncom.com_error:
        return []
    found = []
    for i in range(coll.Count):
        try:
            found.append(typed(coll.ElementAt(i)))
        except pythoncom.com_error:
            continue
    return found


INPUT_TYPES = ("GuiCTextField", "GuiTextField")


def find_input(container, name: str):
    """Returns the editable input field (never its label). Table cells: prefers row 0."""
    candidates = []
    for type_ in INPUT_TYPES:
        candidates += find_all(container, name, type_)
    if not candidates:
        return None
    row0 = [c for c in candidates if "[" not in c.Id or c.Id.endswith(",0]")] or candidates
    for c in row0:
        try:
            if c.Changeable:
                return c
        except pythoncom.com_error:
            continue
    return row0[0]


def find_tab(session, name: str):
    tab = find_by_name(session.findById("wnd[0]/usr"), name, "GuiTab")
    return typed(tab) if tab is not None else None


def find_editor(container, control_name: str):
    """Text editor shell inside custom control TEXT_EDITOR_0101 / TEXT_EDITOR_0201."""
    cc = find_by_name(container, control_name, "GuiCustomControl")
    if cc is None:
        return None
    return _session.findById(rel_id(cc) + "/shellcont/shell")


def status(session) -> tuple[str, str]:
    sbar = session.findById("wnd[0]/sbar")
    return sbar.MessageType, sbar.Text
