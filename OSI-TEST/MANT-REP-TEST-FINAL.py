# =============================================================================
# OSI Arecibo — Inventario, Préstamos y Mantenimientos
# Autores originales: Jorge, Janiel
# Autor de la versión 2.0: Saúl Medina
# =============================================================================

# =============================================================================
# - Inventario desde: Registro Laptops.xlsx
# - Préstamos: Registro_Prestamos_Laptop.xlsx
# - Mantenimientos/Reparaciones: Registro_Mantenimiento_Reparacion_Laptop.xlsx
# - Decomisados: Registro_Decomisados.xlsx
# - Estadísticas clicables (Total/Prestadas/Disponibles)
# - Botón "Decomisadas" para ver Registro_Decomisados.xlsx
# - Búsqueda/acciones por Num_Propiedad (con AUTOCOMPLETADO y atajos)
# - Autenticación por archivo (SHA-256) con TIMER visible
# - Validación: NO permite mantenimiento/reparación si no existe en inventario
# =============================================================================

import os
import re
import hashlib
import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# ------------------------------- Rutas --------------------------------

# Modo PORTABLE: todos los datos viven junto al ejecutable, en ./data/

def app_base_dir() -> str:
    if getattr(sys, 'frozen', False):  # PyInstaller
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = app_base_dir()
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

PATH_INV  = os.path.join(DATA_DIR, "Registro Laptops.xlsx")
PATH_MANT = os.path.join(DATA_DIR, "Registro_Mantenimiento_Reparacion_Laptop.xlsx")
PATH_PREST= os.path.join(DATA_DIR, "Registro_Prestamos_Laptop.xlsx")
PATH_DEC  = os.path.join(DATA_DIR, "Registro_Decomisados.xlsx")

# Encabezados exactos (NO cambiar)
INV_COLS  = [
    "Num_Propiedad","ID_Laptop","Service_Tag","Modelo","Disponible","Garantía","Fecha_Compra"
]
MANT_COLS = [
    "Num_Propiedad","Dia","tecnico","Tipo","Desc_Reparacion","Nombre","Descripcion","Dominio",
    "Check Update","Dell Command Updates","Bios Update","Upgrade Windows 10 - 11",
    "Office 2019 Installed","PatchMyPC Installed","Dell Support Assist Installed"
]
PREST_COLS= ["Num_Propiedad","Nombre","Identificador","Num_Tele","Dia_Pres","Dia_Entr"]
DEC_COLS  = ["Num_Propiedad","ID_Laptop","Service_Tag","Modelo",
             "Num_Mantenimiento","Num_Reparaciones","Num_Prestamos","Fecha_Dec"]

# ------------------------ Autenticación (SHA-256) ----------------------

# Hash provisto por ti (del contenido del archivo de autenticación)
AUTH_HASH = "1c0bcfd0a5eccdb952a74d0570e759d079a54940953470a3d42aa390ed476ff4"
AUTH_WINDOW_SECS = 15 * 60  # 5 minutos

def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()

# ------------------------------ Utils ---------------------------------
# ------------------------------ Utils ---------------------------------
# (pega esto aquí, a nivel de módulo, NO dentro de ninguna clase)

def _to_iso_date(value) -> str:
    import datetime as _dt
    import numpy as _np
    import pandas as pd
    if value is None:
        raise ValueError("Fecha vacía")
    # Timestamp / date
    if isinstance(value, (_dt.date, _dt.datetime, pd.Timestamp)):
        return pd.to_datetime(value).strftime("%Y-%m-%d")
    # Serial Excel (número)
    if isinstance(value, (int, float, _np.integer, _np.floating)):
        dt = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%Y-%m-%d")
    # Cadenas
    s = str(value).strip()
    if not s:
        raise ValueError("Fecha vacía")
    dt = pd.to_datetime(s, errors="coerce")
    if not pd.isna(dt):
        return dt.strftime("%Y-%m-%d")
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return pd.to_datetime(s, format=fmt, errors="raise").strftime("%Y-%m-%d")
        except Exception:
            pass
    raise ValueError(f"No se pudo parsear fecha: {value!r}")

def _read_xlsx(path, expected_cols=None, sheet_name=0):
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
        if expected_cols:
            for c in expected_cols:
                if c not in df.columns:
                    df[c] = ""
            df = df[expected_cols]
        return df
    except FileNotFoundError:
        return pd.DataFrame(columns=expected_cols or [])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer:\n{path}\n\n{e}")
        return pd.DataFrame(columns=expected_cols or [])

def _write_xlsx_exact(df, path, header_order):
    out = df.copy()
    for c in header_order:
        if c not in out.columns:
            out[c] = ""
    out = out[header_order]
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            out.to_excel(w, index=False)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar:\n{path}\n\n{e}")

def _fmt_date_only(v) -> str:
    try:
        dt = pd.to_datetime(v, errors="coerce")
        if pd.isna(dt):
            return "" if (v is None or (isinstance(v,str) and not v.strip())) else str(v)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return str(v) if v is not None else ""

def _now_full():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def _exists_decomisada(num: str) -> bool:
    dec = _read_xlsx(PATH_DEC, DEC_COLS)
    if dec.empty: return False
    return (dec["Num_Propiedad"].astype(str).str.upper()==num.upper()).any()

def _inv_has(num: str) -> bool:
    inv = _read_xlsx(PATH_INV, INV_COLS)
    if inv.empty: return False
    return (inv["Num_Propiedad"].astype(str).str.upper()==num.upper()).any()

def _normkey(s: str) -> str:
    import unicodedata as _ud
    s = "" if s is None else str(s)
    s = _ud.normalize("NFKD", s)
    s = "".join(ch for ch in s if not _ud.combining(ch))
    s = s.lower()
    for ch in (" ", "_", "-", ".", "/"):
        s = s.replace(ch, "")
    return s

def _find_pending_flag_col(df: pd.DataFrame) -> str | None:
    """
    Detecta (si existe) una columna de 'reparación pendiente por pieza' sin renombrarla ni crearla.
    Acepta variantes comunes: 'Esperando_Pieza', 'Pendiente', 'En_Espera', etc.
    Devuelve el nombre EXACTO de la columna si la encuentra; si no, None.
    """
    candidates = {
        "esperandopieza",
        "pendientepieza",
        "pendiente",
        "enespera",
        "enesperapieza",
        "pieza_pendiente",
        "piezaespera"
    }
    for col in df.columns:
        if _normkey(col) in candidates:
            return col
    return None

# ------------------------- Ventanas auxiliares ------------------------

class VentanaPrestamo(ttk.Toplevel):
    def __init__(self, master, num_prop):
        super().__init__(master)
        self.title(f"Préstamos — {num_prop}")
        self.resizable(False, False); self.grab_set()
        self.num_prop = num_prop
        pad = {"padx":10,"pady":6}

        if _exists_decomisada(num_prop):
            ttk.Label(self, text="Esta máquina está DECOMISADA.", foreground="red").grid(row=0, column=0, columnspan=2, **pad); return
        if not _inv_has(num_prop):
            ttk.Label(self, text="Esta máquina NO existe en el inventario.", foreground="red").grid(row=0, column=0, columnspan=2, **pad); return

        inv = _read_xlsx(PATH_INV, INV_COLS)
        row = inv[inv["Num_Propiedad"].astype(str).str.upper()==num_prop.upper()]
        disponible = str(row.iloc[0]["Disponible"]).strip().upper()=="X"
        ttk.Label(self, text=f"Estado actual: {'DISPONIBLE' if disponible else 'PRESTADA'}").grid(row=0, column=0, columnspan=2, **pad)

        if disponible:
            ttk.Label(self, text="Identificador:").grid(row=1, column=0, sticky=E, **pad)
            self.e_ident = ttk.Entry(self, width=28); self.e_ident.grid(row=1, column=1, sticky=EW, **pad)
            ttk.Label(self, text="Nombre:").grid(row=2, column=0, sticky=E, **pad)
            self.e_nombre = ttk.Entry(self, width=28); self.e_nombre.grid(row=2, column=1, sticky=EW, **pad)
            ttk.Label(self, text="Teléfono:").grid(row=3, column=0, sticky=E, **pad)
            self.e_tel = ttk.Entry(self, width=28); self.e_tel.grid(row=3, column=1, sticky=EW, **pad)
            ttk.Button(self, text="Registrar préstamo", bootstyle="success",
                       command=self._prestar).grid(row=4, column=0, columnspan=2, pady=(6,12))
        else:
            ttk.Label(self, text="¿Deseas registrar devolución ahora?").grid(row=1, column=0, columnspan=2, **pad)
            ttk.Button(self, text="Registrar devolución", bootstyle="warning",
                       command=self._devolver).grid(row=2, column=0, columnspan=2, pady=(6,12))

    def _prestar(self):
        ident = self.e_ident.get().strip(); nombre = self.e_nombre.get().strip(); tel = self.e_tel.get().strip()
        if not all([ident, nombre, tel]):
            messagebox.showwarning("Atención","Debes completar Identificador, Nombre y Teléfono."); return
        prest = _read_xlsx(PATH_PREST, PREST_COLS)
        nueva = pd.DataFrame([{
            "Num_Propiedad": self.num_prop, "Nombre": nombre, "Identificador": ident,
            "Num_Tele": tel, "Dia_Pres": _now_full(), "Dia_Entr": ""
        }])
        prest = pd.concat([prest, nueva], ignore_index=True)
        _write_xlsx_exact(prest, PATH_PREST, PREST_COLS)
        inv = _read_xlsx(PATH_INV, INV_COLS)
        m = inv["Num_Propiedad"].astype(str).str.upper()==self.num_prop.upper()
        if m.any():
            inv.loc[m, "Disponible"] = ""
            _write_xlsx_exact(inv, PATH_INV, INV_COLS)
        messagebox.showinfo("Éxito","Préstamo registrado."); self.destroy()

    def _devolver(self):
        prest = _read_xlsx(PATH_PREST, PREST_COLS)
        m = prest["Num_Propiedad"].astype(str).str.upper() == self.num_prop.upper()

        # --- FIX: detectar préstamos sin fecha de devolución ---
        dia_entr = prest["Dia_Entr"]

        is_empty = (
            dia_entr.isna() |
            dia_entr.astype(str).str.strip().isin(["", "NaT", "nan"])
        )

        abiertos = prest[m & is_empty]

        if abiertos.empty:
            messagebox.showwarning("Atención", "No se encontró préstamo pendiente para esta máquina.")
            return

        # --- FIX: actualizar correctamente la fecha ---
        idx = abiertos.tail(1).index[0]
        prest.at[idx, "Dia_Entr"] = _now_full()

        # Asegurar que se escriba correctamente en el Excel
        _write_xlsx_exact(prest, PATH_PREST, PREST_COLS)

        # Marcar máquina como disponible en inventario
        inv = _read_xlsx(PATH_INV, INV_COLS)
        m2 = inv["Num_Propiedad"].astype(str).str.upper() == self.num_prop.upper()
        if m2.any():
            inv.loc[m2, "Disponible"] = "X"
            _write_xlsx_exact(inv, PATH_INV, INV_COLS)

        messagebox.showinfo("Éxito", "Devolución registrada y máquina marcada DISPONIBLE.")
        self.destroy()

class VentanaMantenimiento(ttk.Toplevel):
    """
    Muestra SIEMPRE el checkbox 'Esperando pieza' y un campo '¿Qué pieza?'.
    - Si se marca: registra la reparación como PENDIENTE (Dia vacío) y, si existe
      una columna de flag (p. ej. 'Esperando_Pieza'), la marca con 'X'.
    - Si no se marca: registra la reparación con fecha ahora (Dia = now).
    - Finalización: si hay pendiente (Dia vacío o flag 'X'), permite cerrar el MISMO registro.
    """
    def __init__(self, master, num_prop):
        super().__init__(master)
        self.title(f"Mantenimientos — {num_prop}")
        self.resizable(False, False); self.grab_set()
        self.num_prop = num_prop
        pad = {"padx":10,"pady":6}

        # Bloquear si decomisada
        if _exists_decomisada(num_prop):
            ttk.Label(self, text="Esta máquina está DECOMISADA.", foreground="red").grid(row=0, column=0, columnspan=3, **pad)
            return

        # Cargar mantenimientos y detectar columna de flag si existe
        self.df_mant = _read_xlsx(PATH_MANT, MANT_COLS)
        self.pending_flag_col = _find_pending_flag_col(self.df_mant)

        # ¿Hay reparación pendiente?
        self.pending_idx = self._buscar_reparacion_pendiente(self.df_mant, self.num_prop, self.pending_flag_col)

        ttk.Label(self, text=f"Máquina: {num_prop}", font=("Segoe UI",10,"bold")).grid(row=0, column=0, columnspan=3, **pad)

        if self.pending_idx is not None:
            # --- Finalizar reparación pendiente ---
            ttk.Label(self, text="Se encontró una reparación pendiente (esperando pieza).").grid(row=1, column=0, columnspan=3, **pad)

            ttk.Label(self, text="Técnico:").grid(row=2, column=0, sticky=E, **pad)
            self.e_tec = ttk.Entry(self, width=28); self.e_tec.grid(row=2, column=1, columnspan=2, sticky=EW, **pad)

            ttk.Label(self, text="Descripción final (qué se hizo):").grid(row=3, column=0, columnspan=3, sticky=W, **pad)
            self.t_rep_final = tk.Text(self, width=48, height=5)
            desc_prev = str(self.df_mant.loc[self.pending_idx, "Desc_Reparacion"]) if "Desc_Reparacion" in self.df_mant.columns else ""
            self.t_rep_final.insert("1.0", desc_prev)
            self.t_rep_final.grid(row=4, column=0, columnspan=3, sticky=EW, **pad)

            ttk.Button(self, text="Finalizar reparación", bootstyle="success",
                       command=self._finalizar_reparacion).grid(row=99, column=0, columnspan=3, pady=(6,12))
            return

        # --- Registrar NUEVO (Mantenimiento / Reparación) ---
        ttk.Label(self, text="Tipo:").grid(row=1, column=0, sticky=E, **pad)
        self.tipo_var = tk.StringVar(value="Mantenimiento")
        ttk.Radiobutton(self, text="Mantenimiento", variable=self.tipo_var, value="Mantenimiento",
                        command=self._toggle).grid(row=1, column=1, sticky=W, **pad)
        ttk.Radiobutton(self, text="Reparación", variable=self.tipo_var, value="Reparación",
                        command=self._toggle).grid(row=1, column=2, sticky=W, **pad)

        ttk.Label(self, text="Técnico:").grid(row=2, column=0, sticky=E, **pad)
        self.e_tec = ttk.Entry(self, width=28); self.e_tec.grid(row=2, column=1, columnspan=2, sticky=EW, **pad)

        # ----- Bloque Mantenimiento (checks) -----
        mant_opts = ["Nombre","Descripcion","Dominio","Check Update","Dell Command Updates",
                     "Bios Update","Upgrade Windows 10 - 11","Office 2019 Installed",
                     "PatchMyPC Installed","Dell Support Assist Installed"]
        self.mant_vars = {k: tk.IntVar(value=0) for k in mant_opts}
        self.box_m = ttk.LabelFrame(self, text="Marcar tareas de mantenimiento")
        self.box_m.grid(row=3, column=0, columnspan=3, sticky=EW, padx=10, pady=(2,8))
        i=0
        for k,var in self.mant_vars.items():
            ttk.Checkbutton(self.box_m, text=k, variable=var).grid(row=i//2, column=i%2, sticky=W, padx=8, pady=2)
            i+=1

        # ----- Bloque Reparación (desc + pendiente SIEMPRE visible) -----
        self.box_r = ttk.LabelFrame(self, text="Reparación")
        self.box_r.grid_forget()

        # Checkbox SIEMPRE visible
        self.var_pend = tk.IntVar(value=0)
        self.chk_pend = ttk.Checkbutton(self.box_r, text="Esperando pieza (dejar pendiente sin fecha)", variable=self.var_pend, command=self._toggle_pieza_field)
        self.chk_pend.pack(anchor="w", padx=8, pady=(8,4))

        # Campo para especificar la pieza (solo habilitado si chk está marcado)
        pieza_frame = ttk.Frame(self.box_r); pieza_frame.pack(fill=X, padx=8, pady=(0,6))
        ttk.Label(pieza_frame, text="¿Qué pieza?:").pack(side=LEFT)
        self.e_pieza = ttk.Entry(pieza_frame, width=32)
        self.e_pieza.pack(side=LEFT, padx=(6,0))
        self.e_pieza.configure(state="disabled")

        ttk.Label(self.box_r, text="Descripción de la reparación:").pack(anchor="w", padx=8, pady=(6,0))
        self.t_rep = tk.Text(self.box_r, width=48, height=4)
        self.t_rep.pack(fill=BOTH, expand=True, padx=8, pady=6)

        ttk.Button(self, text="Registrar", bootstyle="success",
                   command=self._registrar).grid(row=99, column=0, columnspan=3, pady=(6,12))

    # ---------- helpers ---------
    # --- Utils de fecha (pegar junto a los otros helpers) ---
    def _toggle(self):
        if self.tipo_var.get()=="Mantenimiento":
            self.box_r.grid_forget(); self.box_m.grid()
        else:
            self.box_m.grid_forget(); self.box_r.grid()

    def _toggle_pieza_field(self):
        if self.var_pend.get()==1:
            self.e_pieza.configure(state="normal")
        else:
            self.e_pieza.delete(0, tk.END)
            self.e_pieza.configure(state="disabled")

    @staticmethod
    def _is_blank(x) -> bool:
        return (x is None) or (str(x).strip() == "") or (pd.isna(x))

    def _buscar_reparacion_pendiente(self, df: pd.DataFrame, num_prop: str, pending_flag_col: str | None):
        if df is None or df.empty:
            return None
        mask_np = df["Num_Propiedad"].astype(str).str.upper() == str(num_prop).upper()
        if "Tipo" not in df.columns:
            return None
        mask_tipo = df["Tipo"].astype(str).str.strip().str.lower() == "reparación".lower()
        mask_dia = df["Dia"].apply(self._is_blank) if "Dia" in df.columns else pd.Series([False]*len(df))
        if pending_flag_col and pending_flag_col in df.columns:
            mask_flag = df[pending_flag_col].astype(str).str.strip().str.upper() == "X"
        else:
            mask_flag = pd.Series([False]*len(df))
        pend = df[mask_np & mask_tipo & (mask_dia | mask_flag)]
        if pend.empty:
            return None
        return pend.tail(1).index[0]

    # ---------- registrar nuevo ----------
    def _registrar(self):
        tec = self.e_tec.get().strip()
        if not tec:
            messagebox.showwarning("Atención","Debes indicar el técnico."); return

        fila = {c:"" for c in MANT_COLS}
        fila["Num_Propiedad"] = self.num_prop
        fila["tecnico"] = tec
        tipo = self.tipo_var.get()
        fila["Tipo"] = tipo

        if tipo == "Mantenimiento":
            fila["Dia"] = _now_full()
            for k, var in self.mant_vars.items():
                if k in fila:
                    fila[k] = "X" if var.get() else ""
            fila["Desc_Reparacion"] = ""
        else:
            desc = self.t_rep.get("1.0","end").strip()
            # Si está esperando pieza, anexa el detalle de la pieza (si se escribió)
            if self.var_pend.get()==1:
                pieza = self.e_pieza.get().strip()
                if pieza:
                    desc = (desc + ("\n" if desc else "")) + f"Pieza en espera: {pieza}"
                fila["Dia"] = ""  # PENDIENTE
            else:
                fila["Dia"] = _now_full()
            fila["Desc_Reparacion"] = desc

        df = self.df_mant.copy()
        df = pd.concat([df, pd.DataFrame([fila])], ignore_index=True)

        # Si existe columna de flag y se marcó pendiente, poner 'X' en esa columna (sin crear columnas nuevas)
        if self.pending_flag_col and tipo == "Reparación" and fila["Dia"] == "":
            if self.pending_flag_col in df.columns:
                df.iloc[-1, df.columns.get_loc(self.pending_flag_col)] = "X"

        _write_xlsx_exact(df, PATH_MANT, MANT_COLS)
        messagebox.showinfo("Éxito", f"{fila['Tipo']} registrada.")
        self.destroy()

    # ---------- finalizar pendiente ----------
    def _finalizar_reparacion(self):
        tec = self.e_tec.get().strip()
        if not tec:
            messagebox.showwarning("Atención","Debes indicar el técnico."); return
        if self.pending_idx is None:
            messagebox.showwarning("Atención","No se encontró reparación pendiente."); return

        desc_final = self.t_rep_final.get("1.0","end").strip()
        if "Desc_Reparacion" in self.df_mant.columns:
            self.df_mant.at[self.pending_idx, "Desc_Reparacion"] = desc_final
        if "tecnico" in self.df_mant.columns:
            self.df_mant.at[self.pending_idx, "tecnico"] = tec
        if "Dia" in self.df_mant.columns:
            self.df_mant.at[self.pending_idx, "Dia"] = _now_full()
        if self.pending_flag_col and self.pending_flag_col in self.df_mant.columns:
            self.df_mant.at[self.pending_idx, self.pending_flag_col] = ""

        _write_xlsx_exact(self.df_mant, PATH_MANT, MANT_COLS)
        messagebox.showinfo("Éxito", "Reparación finalizada.")
        self.destroy()

# ------------------------------- App ---------------------------------

class App(ttk.Window):
    def __init__(self):
        super().__init__(title="OSI Arecibo — Inventario, Préstamos y Mantenimientos",
                         themename="superhero", size=(1120, 820))
        self.view_mode = "inv"  # 'inv' inventario | 'dec' decomisadas

        # --- Autenticación ---
        self.auth_until = None  # datetime o None
        self.timer_job = None

        self.header = ttk.Frame(self, bootstyle="dark")
        self.header.pack(fill=X, padx=12, pady=(12,0))
        ttk.Label(self.header, text="OSI Arecibo — Inventario, Préstamos y Mantenimientos",
                  font=("Segoe UI",16,"bold")).pack(side=LEFT, pady=6)

        # Timer (texto blanco, sin color de fondo)
        self.auth_label = ttk.Label(self.header, text="No autenticado", font=("Segoe UI", 10))
        self.auth_label.pack(side=RIGHT, padx=8, pady=6)

        self._build_toolbar()
        self._build_table()

        self._load_inventory()
        self._refresh_counts()
        self._setup_shortcuts()
        self._update_auth_timer()

    # ---------- UI ----------
    def _build_toolbar(self):
        # LabelFrame con padding interno para respirar
        box = ttk.Labelframe(self, text="Inventario (Registro Laptops.xlsx)")
        box.pack(fill=X, padx=12, pady=12)
        try:
            box.configure(padding=(12, 10))  # padding interno del marco
        except Exception:
            pass  # por si el tema no soporta 'padding'

        # --- Grupo izquierdo: acciones principales ---
        left = ttk.Frame(box)
        left.pack(side=LEFT, padx=6, pady=4)

        # Botones con ancho uniforme y padding consistente
        # Ejemplo para cada botón:
        
        style = ttk.Style()
        style.configure("Small.TButton", padding=(6, 1), font=("Segoe UI", 9))

        ttk.Button(left, text="Decomisar",
                bootstyle=("danger", "toolbutton"),
                style="Small.TButton", width=8,
                command=lambda: self._require_auth(self._decomisar)
        ).pack(side=LEFT, padx=(3, 3), pady=1)

        ttk.Button(left, text="Añadir",
                bootstyle=("success", "toolbutton"),
                style="Small.TButton", width=8,
                command=lambda: self._require_auth(self._add_machine)
        ).pack(side=LEFT, padx=3, pady=1)

        ttk.Button(left, text="Importar…",
                bootstyle=("secondary", "toolbutton"),
                style="Small.TButton", width=9,
                command=lambda: self._require_auth(self._importar_lote)
        ).pack(side=LEFT, padx=3, pady=1)

        ttk.Button(left, text="Autenticar…",
                bootstyle=("light", "toolbutton"),
                style="Small.TButton", width=10,
                command=self._autenticar
        ).pack(side=LEFT, padx=(6, 3), pady=1)

        ttk.Button(left, text="Refrescar",
                bootstyle=("secondary", "toolbutton"),
                style="Small.TButton", width=9,
                command=self._refresh_view
        ).pack(side=LEFT, padx=3, pady=1)

        # Separador vertical para dividir acciones vs. búsqueda
        ttk.Separator(box, orient="vertical").pack(side=LEFT, fill=Y, padx=10, pady=6)

        # --- Grupo centro: búsqueda ---
        center = ttk.Frame(box)
        center.pack(side=LEFT, padx=6, pady=4)

        ttk.Label(center, text="Número de Propiedad:").pack(side=LEFT, padx=(2, 8))

        self.q_var = tk.StringVar()
        self.entry_q = ttk.Combobox(center, width=26, textvariable=self.q_var, values=[], state="normal")
        # ipady para hacer la caja un poquito más alta (si el tema lo soporta)
        self.entry_q.pack(side=LEFT, padx=(0, 8), pady=6, ipady=2)

        ttk.Button(center, text="Buscar", bootstyle="primary", width=10,
                command=self._buscar_info).pack(side=LEFT, padx=6, pady=6)
        ttk.Button(center, text="Préstamos", bootstyle="info", width=12,
                command=self._open_prestamo).pack(side=LEFT, padx=6, pady=6)
        ttk.Button(center, text="Mantenimientos", bootstyle="info", width=14,
                command=self._open_mant).pack(side=LEFT, padx=6, pady=6)

        # Separador vertical antes de stats
        ttk.Separator(box, orient="vertical").pack(side=LEFT, fill=Y, padx=10, pady=6)

        # --- Grupo derecho: estadísticas + Decomisadas ---
        stats = ttk.Frame(box)
        stats.pack(side=RIGHT, padx=8, pady=4)

        self.lbl_total = ttk.Button(stats, text="Total: 0", bootstyle="link",
                                    command=lambda: self._apply_filter(None))
        self.lbl_prest = ttk.Button(stats, text="Prestadas: 0", bootstyle="link",
                                    command=lambda: self._apply_filter("prestadas"))
        self.lbl_disp  = ttk.Button(stats, text="Disponibles: 0", bootstyle="link",
                                    command=lambda: self._apply_filter("disponibles"))
        self.btn_decos = ttk.Button(stats, text="Decomisadas", bootstyle="warning", width=14,
                                    command=self._show_decomisadas)

        # filas/columnas con un poco más de aire
        self.lbl_total.grid(row=0, column=0, padx=6, pady=6)
        self.lbl_prest.grid(row=0, column=1, padx=6, pady=6)
        self.lbl_disp.grid(row=0, column=2, padx=6, pady=6)
        self.btn_decos.grid(row=0, column=3, padx=(10, 6), pady=6)

    def _build_table(self):
        cont = ttk.Frame(self); cont.pack(fill=BOTH, expand=True, padx=12, pady=(0,12))
        self.tree = ttk.Treeview(cont, show="headings", height=22)
        vsb = ttk.Scrollbar(cont, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(cont, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky=NSEW)
        vsb.grid(row=0, column=1, sticky=NS)
        hsb.grid(row=1, column=0, sticky=EW)
        cont.rowconfigure(0, weight=1); cont.columnconfigure(0, weight=1)

    # ---------- Shortcuts ----------
    def _setup_shortcuts(self):
        self.bind("<Return>", lambda e: self._buscar_info())
        self.bind("<Control-p>", lambda e: self._open_prestamo())
        self.bind("<Control-P>", lambda e: self._open_prestamo())
        self.bind("<Control-m>", lambda e: self._open_mant())
        self.bind("<Control-M>", lambda e: self._open_mant())
        self.bind("<Control-n>", lambda e: self._require_auth(self._add_machine))
        self.bind("<Control-N>", lambda e: self._require_auth(self._add_machine))
        self.bind("<F5>", lambda e: self._refresh_view())
        self.bind("<Control-d>", lambda e: self._require_auth(self._decomisar))
        self.bind("<Control-D>", lambda e: self._require_auth(self._decomisar))
        self.bind("<Control-l>", lambda e: self._autenticar())
        self.bind("<Control-L>", lambda e: self._autenticar())

    # ---------- Data loading / view ----------
    def _refresh_view(self):
        if self.view_mode == "dec":
            self._load_decomisadas()
        else:
            self._load_inventory()

    def _load_inventory(self):
        self.view_mode = "inv"
        self.inv_df = _read_xlsx(PATH_INV, INV_COLS)

        # ✅ Ordenar por Num_Propiedad de mayor a menor (sin afectar el archivo original)
        self.inv_df = self.inv_df.sort_values(
            by="Num_Propiedad",
            ascending=False,
            key=lambda s: s.astype(str)
        )

        # actualizar lista para autocompletar
        self.entry_q["values"] = self.inv_df["Num_Propiedad"].astype(str).tolist() if not self.inv_df.empty else []

        # mostrar en la tabla y actualizar los contadores
        self._fill_table(self.inv_df, INV_COLS)
        self._refresh_counts()

    def _load_decomisadas(self):
        self.view_mode = "dec"
        self.dec_df = _read_xlsx(PATH_DEC, DEC_COLS)
        self._fill_table(self.dec_df, DEC_COLS)
        self._refresh_counts()

    def _fill_table(self, df, cols):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            w = 150 if c in ("Service_Tag","ID_Laptop") else 120
            if c in ("Modelo","Garantía","Fecha_Compra","Fecha_Dec"): w=140
            self.tree.column(c, width=w, anchor=W, stretch=True)
        for _, r in df.iterrows():
            vals = []
            for c in cols:
                v = r[c]
                if c in ("Garantía","Fecha_Compra","Fecha_Dec"):
                    vals.append(_fmt_date_only(v))
                else:
                    vals.append("" if pd.isna(v) else str(v))
            self.tree.insert("", tk.END, values=vals)

    def _apply_filter(self, kind):
        if self.view_mode != "inv":
            self._load_inventory()
            return
        if kind is None:
            self._fill_table(self.inv_df, INV_COLS)
        else:
            disp = self.inv_df["Disponible"].astype(str).str.strip().str.upper()
            if kind=="disponibles":
                df = self.inv_df[disp=="X"]
            else:
                df = self.inv_df[disp!="X"]
            self._fill_table(df, INV_COLS)

    def _refresh_counts(self):
        inv = getattr(self, "inv_df", pd.DataFrame(columns=INV_COLS))
        total = len(inv)
        disp = inv["Disponible"].astype(str).str.strip().str.upper()=="X" if not inv.empty else []
        n_disp = int(disp.sum()) if len(disp)>0 else 0
        n_prest = total - n_disp
        self.lbl_total.configure(text=f"Total: {total}")
        self.lbl_prest.configure(text=f"Prestadas: {n_prest}")
        self.lbl_disp.configure(text=f"Disponibles: {n_disp}")

    # ---------- Autenticación ----------
    def _autenticar(self):
        path = filedialog.askopenfilename(title="Seleccionar archivo de autenticación")
        if not path:
            return
        try:
            h = sha256_file(path)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo.\n{e}")
            return
        if h.lower() == AUTH_HASH.lower():
            self.auth_until = datetime.now() + timedelta(seconds=AUTH_WINDOW_SECS)
            messagebox.showinfo("Autenticación", "Autenticación exitosa. Tienes 5 minutos.")
            self._update_auth_timer()
        else:
            self.auth_until = None
            messagebox.showwarning("Autenticación", "Archivo no válido.")
            self._update_auth_timer()

    def _is_authed(self) -> bool:
        return self.auth_until is not None and datetime.now() < self.auth_until

    def _require_auth(self, func):
        if not self._is_authed():
            messagebox.showwarning("Autenticación requerida", "Acción protegida. Autentícate primero (botón «Autenticar…»).")
            return
        func()

    def _update_auth_timer(self):
        # etiqueta en blanco, solo texto
        if self._is_authed():
            remaining = int((self.auth_until - datetime.now()).total_seconds())
            mins = remaining // 60
            secs = remaining % 60
            self.auth_label.configure(text=f"Autenticado: {mins:02d}:{secs:02d} restantes")
        else:
            self.auth_label.configure(text="No autenticado")
        # reprogramar
        if self.timer_job:
            self.after_cancel(self.timer_job)
        self.timer_job = self.after(1000, self._update_auth_timer)

    # ---------- Acciones ----------
    def _show_decomisadas(self):
        self._load_decomisadas()

    def _buscar_info(self):
        num = self.q_var.get().strip()
        if not num:
            messagebox.showwarning("Atención","Ingresa un Num_Propiedad."); return

        if _exists_decomisada(num):
            messagebox.showinfo("Resultado", f"Num_Propiedad {num} está **DECOMISADA**.")
            if self.view_mode != "dec":
                self._load_decomisadas()
            dec = getattr(self, "dec_df", pd.DataFrame(columns=DEC_COLS))
            if not dec.empty:
                sel = dec[dec["Num_Propiedad"].astype(str).str.upper()==num.upper()]
                if not sel.empty: self._fill_table(sel, DEC_COLS)
            return

        if self.view_mode != "inv":
            self._load_inventory()
        row = self.inv_df[self.inv_df["Num_Propiedad"].astype(str).str.upper()==num.upper()]
        if row.empty:
            messagebox.showinfo("Resultado","No se encontró en inventario.")
            return

        r = row.iloc[0]
        estado = "DISPONIBLE" if str(r["Disponible"]).strip().upper()=="X" else "PRESTADA"
        modelo = str(r["Modelo"]); st = str(r["Service_Tag"]); idl = str(r["ID_Laptop"])
        gar = _fmt_date_only(r["Garantía"]); fcomp = _fmt_date_only(r["Fecha_Compra"])

        m = _read_xlsx(PATH_MANT, MANT_COLS)
        mm = m[m["Num_Propiedad"].astype(str).str.upper()==num.upper()]
        cnt_m = int((mm["Tipo"].astype(str)=="Mantenimiento").sum())
        cnt_r = int((mm["Tipo"].astype(str)=="Reparación").sum())
        ult = _fmt_date_only(mm["Dia"].max()) if not mm.empty else "(sin registro)"

        p = _read_xlsx(PATH_PREST, PREST_COLS)
        pp = p[p["Num_Propiedad"].astype(str).str.upper()==num.upper()]
        cnt_p = len(pp)

        msg = (f"Num_Propiedad: {num}\nID_Laptop: {idl}\nService_Tag: {st}\nModelo: {modelo}\n"
               f"Estado: {estado}\n\nMantenimientos: {cnt_m}\nReparaciones: {cnt_r}\n"
               f"Último mant./rep.: {ult}\nPréstamos totales: {cnt_p}\n\n"
               f"Garantía: {gar}\nFecha de compra: {fcomp}")
        messagebox.showinfo("Resumen de la máquina", msg)

    def _open_prestamo(self):
        num = self.q_var.get().strip()
        if not num:
            messagebox.showwarning("Atención","Ingresa un Num_Propiedad primero."); return
        if _exists_decomisada(num):
            messagebox.showwarning("Atención", f"La máquina {num} está DECOMISADA (no se pueden registrar préstamos).")
            return
        if not _inv_has(num):
            messagebox.showwarning("Atención", "Esta máquina NO existe en el inventario. No se puede registrar préstamo.")
            return
        VentanaPrestamo(self, num)

    def _open_mant(self):
        num = self.q_var.get().strip()
        if not num:
            messagebox.showwarning("Atención","Ingresa un Num_Propiedad primero."); return
        if _exists_decomisada(num):
            messagebox.showwarning("Atención", f"La máquina {num} está DECOMISADA (no se pueden registrar mantenimientos/reparaciones).")
            return
        if not _inv_has(num):
            messagebox.showwarning("Atención", "Esta máquina NO existe en el inventario. No se puede registrar mantenimiento/reparación.")
            return
        VentanaMantenimiento(self, num)

    def _add_machine(self):
        win = ttk.Toplevel(self); win.title("Añadir máquina"); win.resizable(False, False); win.grab_set()
        pad={"padx":10,"pady":6}
        ttk.Label(win, text="Número de Propiedad (ej. R40022104):").grid(row=0, column=0, sticky=E, **pad)
        e_np=ttk.Entry(win, width=28); e_np.grid(row=0, column=1, **pad)
        ttk.Label(win, text="ID de Laptop (ej. UIPRA-EST-L045):").grid(row=1, column=0, sticky=E, **pad)
        e_id=ttk.Entry(win, width=28); e_id.grid(row=1, column=1, **pad)
        ttk.Label(win, text="Service Tag (7 chars, ej. 4TR2M53):").grid(row=2, column=0, sticky=E, **pad)
        e_st=ttk.Entry(win, width=28); e_st.grid(row=2, column=1, **pad)
        ttk.Label(win, text="Modelo (ej. 5510):").grid(row=3, column=0, sticky=E, **pad)
        e_md=ttk.Entry(win, width=28); e_md.grid(row=3, column=1, **pad)
        ttk.Label(win, text="Garantía (YYYY-MM-DD):").grid(row=4, column=0, sticky=E, **pad)
        e_ga=ttk.Entry(win, width=28); e_ga.grid(row=4, column=1, **pad)
        ttk.Label(win, text="Fecha de compra (YYYY-MM-DD):").grid(row=5, column=0, sticky=E, **pad)
        e_fc=ttk.Entry(win, width=28); e_fc.grid(row=5, column=1, **pad)

        def guardar():
            npv=e_np.get().strip().upper(); idv=e_id.get().strip().upper(); stv=e_st.get().strip().upper()
            mdv=e_md.get().strip(); gav=e_ga.get().strip(); fcv=e_fc.get().strip()
            errs=[]
            if not re.fullmatch(r"R\d{8}", npv): errs.append("Número de Propiedad inválido (R + 8 dígitos).")
            if not re.fullmatch(r"UIPRA-(EST|FAC)-L\d{3}", idv): errs.append("ID_Laptop inválido (UIPRA-(EST|FAC)-L###).")
            if not re.fullmatch(r"[A-Z0-9]{7}", stv): errs.append("Service_Tag inválido (7 alfanuméricos en MAYÚSCULA).")
            try:
                dtg = pd.to_datetime(gav, format="%Y-%m-%d", errors="raise")
                if dtg.date() <= datetime.now().date(): errs.append("Garantía debe ser FUTURA (YYYY-MM-DD).")
            except Exception: errs.append("Garantía inválida (YYYY-MM-DD).")
            try:
                _ = pd.to_datetime(fcv, format="%Y-%m-%d", errors="raise")
            except Exception: errs.append("Fecha de compra inválida (YYYY-MM-DD).")
            inv=_read_xlsx(PATH_INV, INV_COLS)
            if (inv["Num_Propiedad"].astype(str).str.upper()==npv).any(): errs.append("Num_Propiedad duplicado.")
            if (inv["ID_Laptop"].astype(str).str.upper()==idv).any(): errs.append("ID_Laptop duplicado.")
            if (inv["Service_Tag"].astype(str).str.upper()==stv).any(): errs.append("Service_Tag duplicado.")
            if _exists_decomisada(npv): errs.append("Ese Num_Propiedad aparece en decomisados.")
            if errs:
                messagebox.showwarning("Datos inválidos", "\n".join(f"• {e}" for e in errs)); return
            new = pd.DataFrame([{
                "Num_Propiedad": npv, "ID_Laptop": idv, "Service_Tag": stv,
                "Modelo": mdv, "Disponible": "X", "Garantía": gav, "Fecha_Compra": fcv
            }])
            inv = pd.concat([inv, new], ignore_index=True)
            _write_xlsx_exact(inv, PATH_INV, INV_COLS)
            messagebox.showinfo("Éxito","Máquina añadida."); win.destroy(); self._load_inventory()
        ttk.Button(win, text="Guardar", bootstyle="success", command=guardar).grid(row=6, column=0, columnspan=2, pady=(6,12))

    def _decomisar(self):
        num = self.q_var.get().strip()
        if not num:
            messagebox.showwarning("Atención","Ingresa un Num_Propiedad para decomisar."); return
        if _exists_decomisada(num):
            messagebox.showwarning("Atención", f"La máquina {num} ya está DECOMISADA."); return

        inv = _read_xlsx(PATH_INV, INV_COLS)
        row = inv[inv["Num_Propiedad"].astype(str).str.upper()==num.upper()]
        if row.empty:
            messagebox.showwarning("Atención","No está en inventario. Si fue decomisada antes, usa 'Decomisadas'.")
            return

        if not messagebox.askyesno("Confirmar", f"¿Decomisar la máquina {num}?"):
            return

        r = row.iloc[0]
        id_lap = str(r["ID_Laptop"]); st = str(r["Service_Tag"]); modelo = str(r["Modelo"])

        mant = _read_xlsx(PATH_MANT, MANT_COLS)
        mm = mant[mant["Num_Propiedad"].astype(str).str.upper()==num.upper()]
        num_m = int((mm["Tipo"].astype(str)=="Mantenimiento").sum())
        num_r = int((mm["Tipo"].astype(str)=="Reparación").sum())
        prest = _read_xlsx(PATH_PREST, PREST_COLS)
        pp = prest[prest["Num_Propiedad"].astype(str).str.upper()==num.upper()]
        num_p = len(pp)

        dec = _read_xlsx(PATH_DEC, DEC_COLS)
        nueva = pd.DataFrame([{
            "Num_Propiedad": num, "ID_Laptop": id_lap, "Service_Tag": st, "Modelo": modelo,
            "Num_Mantenimiento": num_m, "Num_Reparaciones": num_r, "Num_Prestamos": num_p,
            "Fecha_Dec": _now_full()
        }])
        dec = pd.concat([dec, nueva], ignore_index=True)
        _write_xlsx_exact(dec, PATH_DEC, DEC_COLS)

        if messagebox.askyesno("Inventario", "Decomiso guardado.\n\n¿Quitar del inventario ahora?"):
            inv = inv[inv["Num_Propiedad"].astype(str).str.upper()!=num.upper()].copy()
            _write_xlsx_exact(inv, PATH_INV, INV_COLS)
            self._load_inventory()
            messagebox.showinfo("Hecho","Máquina retirada del inventario.")
        else:
            messagebox.showinfo("Hecho","Decomiso registrado (inventario se mantiene).")

    # ---------- Importación por lote (protegida) ----------
    def _importar_lote(self):
        # Solo informa estructura y luego abre archivo
        messagebox.showinfo(
            "Estructura requerida",
            "El archivo debe tener columnas EXACTAS:\n"
            "Num_Propiedad, ID_Laptop, Service_Tag, Modelo, Garantía, Fecha_Compra\n\n"
            "Formatos:\n"
            "• Num_Propiedad: R + 8 dígitos (p.ej., R40022104)\n"
            "• ID_Laptop: UIPRA-(EST|FAC)-L###\n"
            "• Service_Tag: 7 caracteres alfanuméricos en mayúsculas\n"
            "• Garantía y Fecha_Compra: YYYY-MM-DD (Garantía debe ser futura)"
        )
        path = filedialog.askopenfilename(title="Seleccionar archivo de máquinas", filetypes=[("Excel","*.xlsx *.xls")])
        if not path:
            return

        df = _read_xlsx(path, expected_cols=["Num_Propiedad","ID_Laptop","Service_Tag","Modelo","Garantía","Fecha_Compra"])
        if df.empty:
            messagebox.showwarning("Importación cancelada","No se encontraron filas válidas."); return

        errs = []
        inv = _read_xlsx(PATH_INV, INV_COLS)

        for i, row in df.iterrows():
            npv = str(row["Num_Propiedad"]).strip().upper()
            idv = str(row["ID_Laptop"]).strip().upper()
            stv = str(row["Service_Tag"]).strip().upper()
            mdv = str(row["Modelo"]).strip()

            # >>> LEE LAS FECHAS CRUDAS DE LA FILA <<<
            gav_raw = row["Garantía"]
            fcv_raw = row["Fecha_Compra"]

            # --- Normaliza fechas a ISO ---
            try:
                gav_iso = _to_iso_date(gav_raw)  # 'YYYY-MM-DD'
            except Exception:
                errs.append(f"Fila {i+2}: Garantía inválida (YYYY-MM-DD).")
                continue

            try:
                fcv_iso = _to_iso_date(fcv_raw)  # 'YYYY-MM-DD'
            except Exception:
                errs.append(f"Fila {i+2}: Fecha_Compra inválida (YYYY-MM-DD).")
                continue

            # --- Validaciones ---
            if not re.fullmatch(r"R\d{8}", npv):
                errs.append(f"Fila {i+2}: Num_Propiedad inválido."); continue
            if not re.fullmatch(r"UIPRA-(EST|FAC)-L\d{3}", idv):
                errs.append(f"Fila {i+2}: ID_Laptop inválido."); continue
            if not re.fullmatch(r"[A-Z0-9]{7}", stv):
                errs.append(f"Fila {i+2}: Service_Tag inválido."); continue

            # Garantía debe ser futura
            try:
                if pd.to_datetime(gav_iso).date() <= datetime.now().date():
                    errs.append(f"Fila {i+2}: Garantía no es futura."); continue
            except Exception:
                errs.append(f"Fila {i+2}: Garantía inválida."); continue

            # Duplicados y decomisados
            if (inv["Num_Propiedad"].astype(str).str.upper()==npv).any():
                errs.append(f"Fila {i+2}: Num_Propiedad duplicado."); continue
            if (inv["ID_Laptop"].astype(str).str.upper()==idv).any():
                errs.append(f"Fila {i+2}: ID_Laptop duplicado."); continue
            if (inv["Service_Tag"].astype(str).str.upper()==stv).any():
                errs.append(f"Fila {i+2}: Service_Tag duplicado."); continue
            if _exists_decomisada(npv):
                errs.append(f"Fila {i+2}: Num_Propiedad aparece en decomisados."); continue

            # --- Inserta usando las fechas normalizadas ---
            inv = pd.concat([inv, pd.DataFrame([{
                "Num_Propiedad": npv,
                "ID_Laptop": idv,
                "Service_Tag": stv,
                "Modelo": mdv,
                "Disponible": "X",
                "Garantía": gav_iso,
                "Fecha_Compra": fcv_iso
            }])], ignore_index=True)

        if errs:
            messagebox.showwarning("Importación cancelada", "Se encontraron problemas y NO se importó nada:\n\n• " + "\n• ".join(errs))
            return

        _write_xlsx_exact(inv, PATH_INV, INV_COLS)
        messagebox.showinfo("Éxito","Importación completada.")
        self._load_inventory()

# ------------------------------------ Run -----------------------------------

if __name__ == "__main__":
    App().mainloop()