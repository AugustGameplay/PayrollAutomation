#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================================
 VINCULADOR DE REGISTROS  ·  Record Linker
 Primary Staffing Inc.
----------------------------------------------------------------------------
 Enlaza dos archivos de Excel/CSV de personas que NO comparten un
 identificador confiable (sin CURP, INE, NSS). Usa una estrategia por
 niveles:
    1) Employee ID (cuando coincide)  -> match fuerte
    2) Nombre (fuzzy) + Monto neto     -> match probable
    3) Nombres ambiguos / sin match    -> revisión MANUAL
 Genera un CSV consolidado con: método de match, % de confianza,
 margen de error y estado (AUTOMÁTICO / REVISAR MANUAL).
============================================================================
 Requisitos (una sola vez):
    pip install pandas openpyxl rapidfuzz
 Ejecutar:
    python record_linker.py
============================================================================
"""

import os
import re
import csv
import queue
import threading
import unicodedata
import datetime as dt

import pandas as pd

try:
    from rapidfuzz import fuzz, process
    HAVE_RF = True
except Exception:
    HAVE_RF = False

# ===========================================================================
#  MOTOR DE VINCULACIÓN  (independiente de la interfaz)
# ===========================================================================

# Pistas para autodetectar columnas clave en cada archivo
HINTS_NAME   = ["name", "nombre", "employee name", "full name"]
HINTS_ID     = ["employee id", "emp id", "id", "ee number", "ee#", "id empleado"]
HINTS_AMOUNT = ["net pay", "net amount", "neto", "net", "deposit", "check amount"]


def _ratio(a, b):
    """token_sort_ratio si hay rapidfuzz; respaldo simple si no."""
    if HAVE_RF:
        return fuzz.token_sort_ratio(a, b)
    # Respaldo (Jaccard de tokens) por si falta rapidfuzz
    sa, sb = set(a.split()), set(b.split())
    if not sa or not sb:
        return 0
    return 100.0 * len(sa & sb) / len(sa | sb)


def norm_id(v):
    """Normaliza un Employee ID: quita prefijos ('Emp ID :', 'EE Number:'),
    deja solo letras y números en mayúscula."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).upper()
    s = re.sub(r"EMP\s*ID\s*:?", "", s)
    s = re.sub(r"EE\s*NUMBER\s*:?", "", s)
    s = re.sub(r"CHECK\s*NO\s*:?", "", s)
    return re.sub(r"[^A-Z0-9]", "", s)


def norm_name(v):
    """Normaliza un nombre para comparar: minúsculas, sin acentos,
    sin puntuación, tokens ordenados (robusto a 'Apellido, Nombre')."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = s.replace(",", " ")
    s = re.sub(r"[^a-z ]", " ", s)
    return " ".join(sorted(s.split()))


def to_num(v):
    try:
        return round(float(v), 2)
    except Exception:
        return None


# Patrones de filas que NO son personas (subtotales / pies de reporte)
JUNK_RE = re.compile(
    r"^\s*(totals?\b|total for|totals for|grand total|report total|"
    r"company name|selection criteria|page \d)", re.I)


def is_person(name):
    """False si el 'nombre' es en realidad un subtotal o pie de reporte."""
    if name is None:
        return False
    s = str(name).strip()
    if not s:
        return False
    if JUNK_RE.match(s):
        return False
    # debe contener al menos una letra
    return bool(re.search(r"[A-Za-z]", s))


def autodetect(columns, hints):
    """Devuelve la primera columna cuyo nombre contiene alguna pista."""
    low = {c: str(c).strip().lower() for c in columns}
    for h in hints:
        for c, cl in low.items():
            if cl == h:
                return c
    for h in hints:
        for c, cl in low.items():
            if h in cl:
                return c
    return None


def parse_blocks(df, name_col, id_col, amount_col):
    """Convierte un reporte tipo 'registro' (donde cada persona ocupa una
    fila cabecera + varias filas de detalle vacías) en un registro por
    persona. Una fila es cabecera cuando la columna Nombre NO está vacía.
    Conserva TODAS las columnas de la fila cabecera para no perder datos."""
    records = []
    header_idx = df.index[df[name_col].notna()].tolist()
    header_idx.append(len(df))
    for i in range(len(header_idx) - 1):
        start = header_idx[i]
        head = df.iloc[start]
        if not is_person(head[name_col]):
            continue  # salta subtotales / pies de reporte
        rec = {
            "_name": str(head[name_col]).strip(),
            "_eid": norm_id(head[id_col]) if id_col else "",
            "_namekey": norm_name(head[name_col]),
            "_amount": to_num(head[amount_col]) if amount_col else None,
        }
        # arrastra todas las columnas cabecera (datos completos)
        for c in df.columns:
            val = head[c]
            rec[c] = "" if (val is None or (isinstance(val, float) and pd.isna(val))) else val
        records.append(rec)
    return records


def confidence(method, name_sim, amt_match, ambiguous):
    """Calcula confianza 0-100 según la evidencia."""
    if method == "ID":
        if name_sim >= 80 and amt_match:
            return 100
        if name_sim >= 80 or amt_match:
            return 97
        return 88  # ID igual pero ni nombre ni monto corroboran -> ojo
    if method == "NOMBRE":
        c = float(name_sim)
        if amt_match:
            c = min(99, c + 8)
        if ambiguous:                 # dos candidatos casi idénticos
            c = min(c, 78)            # se manda a revisión
        return round(c, 1)
    return 0


def link(records_a, records_b, threshold, progress_cb=None):
    """Vincula A (izquierda) contra B (derecha). Devuelve filas de resultado."""
    # índice de B por ID
    b_by_id = {}
    for rb in records_b:
        if rb["_eid"]:
            b_by_id.setdefault(rb["_eid"], []).append(rb)
    b_names = [rb["_namekey"] for rb in records_b]

    matched_b = set()
    matched_eids = set()
    rows = []
    n = len(records_a)
    for i, ra in enumerate(records_a):
        method, b, name_sim, amt_match, ambiguous = "SIN MATCH", None, 0, False, False

        # --- Nivel 1: Employee ID exacto ---
        if ra["_eid"] and ra["_eid"] in b_by_id:
            cands = b_by_id[ra["_eid"]]
            # si hay varios (varios cheques), toma el de nombre más parecido
            b = max(cands, key=lambda x: _ratio(ra["_namekey"], x["_namekey"]))
            name_sim = _ratio(ra["_namekey"], b["_namekey"])
            amt_match = (ra["_amount"] is not None and b["_amount"] is not None
                         and abs(ra["_amount"] - b["_amount"]) < 0.01)
            method = "ID"

        # --- Nivel 2: Nombre fuzzy (+ monto) ---
        elif b_names:
            if HAVE_RF:
                top = process.extract(ra["_namekey"], b_names,
                                      scorer=fuzz.token_sort_ratio, limit=2)
            else:
                scored = sorted(
                    [(bn, _ratio(ra["_namekey"], bn), j) for j, bn in enumerate(b_names)],
                    key=lambda x: x[1], reverse=True)[:2]
                top = scored
            if top and top[0][1] >= 80:
                b = records_b[top[0][2]]
                name_sim = top[0][1]
                amt_match = (ra["_amount"] is not None and b["_amount"] is not None
                             and abs(ra["_amount"] - b["_amount"]) < 0.01)
                # ¿ambiguo? segundo candidato casi igual de bueno
                if len(top) > 1 and top[1][1] >= 85 and (top[0][1] - top[1][1]) < 5:
                    ambiguous = True
                method = "NOMBRE"

        conf = confidence(method, name_sim, amt_match, ambiguous)
        status = "AUTOMATICO" if conf >= threshold else "REVISAR MANUAL"
        if b is not None and conf >= threshold:
            matched_b.add(id(b))
            if b["_eid"]:
                matched_eids.add(b["_eid"])

        rows.append({
            "ra": ra, "rb": b, "method": method, "conf": conf,
            "margin": round(100 - conf, 1), "status": status,
            "name_sim": round(name_sim, 1), "amt_match": amt_match,
            "ambiguous": ambiguous,
        })
        if progress_cb and (i % 15 == 0 or i == n - 1):
            progress_cb(i + 1, n, ra["_name"])

    # registros de B que nunca se vincularon -> revisión manual (B-only).
    # Un 2do cheque de alguien ya vinculado (mismo EID) NO es huérfano.
    b_only = [rb for rb in records_b
              if id(rb) not in matched_b
              and not (rb["_eid"] and rb["_eid"] in matched_eids)]
    return rows, b_only


def build_output(rows, b_only, cols_a, cols_b):
    """Arma la tabla final consolidada (lista de dicts) para el CSV."""
    out = []
    for r in rows:
        ra, rb = r["ra"], r["rb"]
        row = {"Origen": "A (New)"}
        for c in cols_a:
            row[f"A_{c}"] = ra.get(c, "")
        for c in cols_b:
            row[f"B_{c}"] = (rb.get(c, "") if rb is not None else "")
        row["Metodo"] = r["method"]
        row["Confianza_%"] = r["conf"]
        row["Margen_error_%"] = r["margin"]
        row["Similitud_nombre_%"] = r["name_sim"]
        row["Monto_coincide"] = "SI" if r["amt_match"] else "NO"
        row["Nombre_ambiguo"] = "SI" if r["ambiguous"] else "NO"
        row["Estado"] = r["status"]
        out.append(row)
    # B sin pareja
    for rb in b_only:
        row = {"Origen": "B (Regular) SIN PAREJA"}
        for c in cols_a:
            row[f"A_{c}"] = ""
        for c in cols_b:
            row[f"B_{c}"] = rb.get(c, "")
        row["Metodo"] = "SIN MATCH"
        row["Confianza_%"] = 0
        row["Margen_error_%"] = 100
        row["Similitud_nombre_%"] = 0
        row["Monto_coincide"] = "NO"
        row["Nombre_ambiguo"] = "NO"
        row["Estado"] = "REVISAR MANUAL"
        out.append(row)
    return out


def load_table(path):
    """Carga xlsx/xls/csv como DataFrame (primera hoja)."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".tsv", ".txt"):
        sep = "\t" if ext == ".tsv" else None
        return pd.read_csv(path, sep=sep, engine="python", dtype=str)
    return pd.read_excel(path, sheet_name=0)


# ===========================================================================
#  INTERFAZ GRÁFICA  (Tkinter)
# ===========================================================================

def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    # ---- paleta ----
    BG     = "#0f172a"   # fondo
    PANEL  = "#1e293b"   # paneles
    CARD   = "#273449"
    ACCENT = "#3b82f6"   # azul
    GREEN  = "#22c55e"
    AMBER  = "#f59e0b"
    RED    = "#ef4444"
    TXT    = "#e2e8f0"
    MUTE   = "#94a3b8"

    app = tk.Tk()
    app.title("Vinculador de Registros · Primary Staffing")
    app.geometry("1180x760")
    app.configure(bg=BG)
    app.minsize(980, 640)

    state = {
        "path_a": tk.StringVar(),
        "path_b": tk.StringVar(),
        "df_a": None, "df_b": None,
        "name_a": tk.StringVar(), "id_a": tk.StringVar(), "amt_a": tk.StringVar(),
        "name_b": tk.StringVar(), "id_b": tk.StringVar(), "amt_b": tk.StringVar(),
        "threshold": tk.IntVar(value=90),
        "result_rows": None, "out_table": None,
        "cols_a": None, "cols_b": None,
    }
    q = queue.Queue()

    # ---- estilos ttk ----
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure(".", background=PANEL, foreground=TXT, fieldbackground=PANEL,
                    bordercolor=CARD, font=("Segoe UI", 10))
    style.configure("TFrame", background=PANEL)
    style.configure("Card.TFrame", background=CARD)
    style.configure("TLabel", background=PANEL, foreground=TXT)
    style.configure("Card.TLabel", background=CARD, foreground=TXT)
    style.configure("Mute.TLabel", background=PANEL, foreground=MUTE)
    style.configure("Title.TLabel", background=BG, foreground=TXT,
                    font=("Segoe UI Semibold", 18))
    style.configure("Sub.TLabel", background=BG, foreground=MUTE, font=("Segoe UI", 10))
    style.configure("Accent.TButton", background=ACCENT, foreground="white",
                    font=("Segoe UI Semibold", 11), borderwidth=0, padding=10)
    style.map("Accent.TButton", background=[("active", "#2563eb"), ("disabled", "#334155")])
    style.configure("TButton", background=CARD, foreground=TXT, borderwidth=0, padding=7)
    style.map("TButton", background=[("active", "#33415a")])
    style.configure("TCombobox", fieldbackground=PANEL, background=CARD,
                    foreground=TXT, arrowcolor=TXT)
    style.configure("Treeview", background=PANEL, fieldbackground=PANEL,
                    foreground=TXT, rowheight=26, borderwidth=0)
    style.configure("Treeview.Heading", background=CARD, foreground=TXT,
                    font=("Segoe UI Semibold", 10), borderwidth=0)
    style.map("Treeview", background=[("selected", ACCENT)])
    style.configure("Horizontal.TProgressbar", background=ACCENT, troughcolor=CARD,
                    borderwidth=0)

    # ---- encabezado ----
    head = tk.Frame(app, bg=BG)
    head.pack(fill="x", padx=22, pady=(18, 4))
    ttk.Label(head, text="Vinculador de Registros", style="Title.TLabel").pack(anchor="w")
    ttk.Label(head, text="Enlaza dos archivos de personas sin identificador único · ID › Nombre+Monto › Revisión manual",
              style="Sub.TLabel").pack(anchor="w")

    body = tk.Frame(app, bg=BG)
    body.pack(fill="both", expand=True, padx=22, pady=12)
    body.columnconfigure(0, weight=0, minsize=380)
    body.columnconfigure(1, weight=1)
    body.rowconfigure(0, weight=1)

    # ============ PANEL IZQUIERDO (configuración) ============
    left = tk.Frame(body, bg=PANEL)
    left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))

    def section(parent, title):
        f = tk.Frame(parent, bg=PANEL)
        f.pack(fill="x", padx=16, pady=(14, 0))
        ttk.Label(f, text=title, style="Mute.TLabel",
                  font=("Segoe UI Semibold", 10)).pack(anchor="w")
        return f

    def file_picker(parent, label, path_var, on_load):
        sec = section(parent, label)
        row = tk.Frame(sec, bg=PANEL)
        row.pack(fill="x", pady=6)
        ent = tk.Entry(row, textvariable=path_var, bg=CARD, fg=TXT,
                       insertbackground=TXT, relief="flat")
        ent.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))

        def browse():
            p = filedialog.askopenfilename(
                title=f"Selecciona {label}",
                filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv *.tsv"), ("Todos", "*.*")])
            if p:
                path_var.set(p)
                on_load(p)
        ttk.Button(row, text="Examinar…", command=browse).pack(side="left")
        return sec

    # combos para mapear columnas
    def col_mapper(parent, name_v, id_v, amt_v):
        f = tk.Frame(parent, bg=PANEL)
        f.pack(fill="x", padx=16, pady=(2, 4))
        grid = tk.Frame(f, bg=PANEL)
        grid.pack(fill="x")
        combos = {}
        for i, (lbl, var) in enumerate([("Nombre", name_v), ("Employee ID", id_v),
                                        ("Monto neto", amt_v)]):
            ttk.Label(grid, text=lbl, style="Mute.TLabel").grid(
                row=i, column=0, sticky="w", pady=3, padx=(0, 8))
            cb = ttk.Combobox(grid, textvariable=var, state="readonly", width=26)
            cb.grid(row=i, column=1, sticky="ew", pady=3)
            combos[lbl] = cb
        grid.columnconfigure(1, weight=1)
        return combos

    def load_file(which, path):
        try:
            df = load_table(path)
        except Exception as e:
            messagebox.showerror("Error al leer archivo", str(e))
            return
        cols = list(df.columns)
        state[f"df_{which}"] = df
        nm = autodetect(cols, HINTS_NAME)
        idd = autodetect(cols, HINTS_ID)
        amt = autodetect(cols, HINTS_AMOUNT)
        combos = combos_a if which == "a" else combos_b
        for lbl, cb in combos.items():
            cb["values"] = cols
        state[f"name_{which}"].set(nm or (cols[0] if cols else ""))
        state[f"id_{which}"].set(idd or "")
        state[f"amt_{which}"].set(amt or "")
        log(f"✓ Archivo {which.upper()} cargado: {os.path.basename(path)} "
            f"({len(df)} filas, {len(cols)} columnas)")
        log(f"   Autodetectado → Nombre: {nm} | ID: {idd} | Monto: {amt}")

    file_picker(left, "ARCHIVO A  (ej. Register New)", state["path_a"],
                lambda p: load_file("a", p))
    combos_a = col_mapper(left, state["name_a"], state["id_a"], state["amt_a"])

    file_picker(left, "ARCHIVO B  (ej. Register Regular)", state["path_b"],
                lambda p: load_file("b", p))
    combos_b = col_mapper(left, state["name_b"], state["id_b"], state["amt_b"])

    # umbral
    th_sec = section(left, "UMBRAL DE CONFIANZA (auto vs. manual)")
    th_row = tk.Frame(th_sec, bg=PANEL)
    th_row.pack(fill="x", pady=6)
    th_lbl = ttk.Label(th_row, text="90%", style="TLabel",
                       font=("Segoe UI Semibold", 12))
    th_lbl.pack(side="right")
    scale = ttk.Scale(th_row, from_=70, to=100, orient="horizontal",
                      variable=state["threshold"],
                      command=lambda v: th_lbl.config(text=f"{int(float(v))}%"))
    scale.pack(side="left", fill="x", expand=True, padx=(0, 10))
    ttk.Label(left, text="Arriba del umbral → AUTOMÁTICO.  Debajo → REVISAR MANUAL.",
              style="Mute.TLabel", font=("Segoe UI", 9)).pack(anchor="w", padx=16)

    # botón procesar
    btn_run = ttk.Button(left, text="▶  Procesar y vincular",
                         style="Accent.TButton", command=lambda: start_process())
    btn_run.pack(fill="x", padx=16, pady=(18, 4))

    # progreso
    prog = ttk.Progressbar(left, style="Horizontal.TProgressbar", maximum=100)
    prog.pack(fill="x", padx=16, pady=(6, 2))
    prog_lbl = ttk.Label(left, text="Listo.", style="Mute.TLabel", font=("Segoe UI", 9))
    prog_lbl.pack(anchor="w", padx=16)

    # ============ PANEL DERECHO (resultados + log) ============
    right = tk.Frame(body, bg=PANEL)
    right.grid(row=0, column=1, sticky="nsew")
    right.rowconfigure(2, weight=3)
    right.rowconfigure(4, weight=1)
    right.columnconfigure(0, weight=1)

    # tarjetas resumen
    cards = tk.Frame(right, bg=PANEL)
    cards.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 6))
    card_vals = {}

    def make_card(parent, key, title, color):
        c = tk.Frame(parent, bg=CARD, highlightbackground=color,
                     highlightthickness=2)
        c.pack(side="left", expand=True, fill="x", padx=4)
        v = tk.Label(c, text="—", bg=CARD, fg=color,
                     font=("Segoe UI Semibold", 22))
        v.pack(anchor="w", padx=12, pady=(8, 0))
        tk.Label(c, text=title, bg=CARD, fg=MUTE,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=12, pady=(0, 8))
        card_vals[key] = v

    make_card(cards, "auto", "Automáticos", GREEN)
    make_card(cards, "manual", "Revisar manual", AMBER)
    make_card(cards, "pct", "% Auto-match", ACCENT)
    make_card(cards, "total", "Total A", TXT)

    ttk.Label(right, text="RESULTADOS", style="Mute.TLabel",
              font=("Segoe UI Semibold", 10)).grid(row=1, column=0, sticky="w",
                                                   padx=16, pady=(8, 2))

    # tabla resultados
    tbl_frame = tk.Frame(right, bg=PANEL)
    tbl_frame.grid(row=2, column=0, sticky="nsew", padx=16)
    cols = ("nombre", "id", "metodo", "conf", "margen", "estado")
    tree = ttk.Treeview(tbl_frame, columns=cols, show="headings")
    for c, txt, w in [("nombre", "Nombre (A)", 230), ("id", "Employee ID", 110),
                      ("metodo", "Método", 100), ("conf", "Confianza %", 95),
                      ("margen", "Margen err %", 100), ("estado", "Estado", 140)]:
        tree.heading(c, text=txt)
        tree.column(c, width=w, anchor="w")
    vsb = ttk.Scrollbar(tbl_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    tree.pack(side="left", fill="both", expand=True)
    vsb.pack(side="right", fill="y")
    tree.tag_configure("auto", foreground=GREEN)
    tree.tag_configure("manual", foreground=AMBER)
    tree.tag_configure("none", foreground=RED)

    # exportar
    exp = tk.Frame(right, bg=PANEL)
    exp.grid(row=3, column=0, sticky="ew", padx=16, pady=(8, 4))
    btn_csv = ttk.Button(exp, text="⬇  Exportar CSV consolidado",
                         command=lambda: export_csv(False), state="disabled")
    btn_csv.pack(side="left", padx=(0, 8))
    btn_rev = ttk.Button(exp, text="⬇  Exportar solo 'Revisar manual'",
                         command=lambda: export_csv(True), state="disabled")
    btn_rev.pack(side="left")

    # log
    ttk.Label(right, text="REGISTRO DE ACTIVIDAD", style="Mute.TLabel",
              font=("Segoe UI Semibold", 10)).grid(row=4, column=0, sticky="nw",
                                                   padx=16, pady=(8, 2))
    log_box = tk.Text(right, height=7, bg="#0b1220", fg="#9fb3c8",
                      insertbackground=TXT, relief="flat", font=("Consolas", 9),
                      wrap="word")
    log_box.grid(row=5, column=0, sticky="nsew", padx=16, pady=(0, 14))

    def log(msg):
        ts = dt.datetime.now().strftime("%H:%M:%S")
        log_box.insert("end", f"[{ts}] {msg}\n")
        log_box.see("end")

    log("Selecciona los dos archivos. Las columnas clave se autodetectan; "
        "puedes corregirlas en los menús.")
    if not HAVE_RF:
        log("⚠ rapidfuzz no está instalado: usando comparación básica. "
            "Recomendado: pip install rapidfuzz")

    # ---------- proceso en hilo ----------
    def start_process():
        if state["df_a"] is None or state["df_b"] is None:
            messagebox.showwarning("Faltan archivos", "Carga el Archivo A y el Archivo B.")
            return
        if not state["name_a"].get() or not state["name_b"].get():
            messagebox.showwarning("Falta columna", "Selecciona la columna de Nombre en ambos archivos.")
            return
        btn_run.config(state="disabled")
        btn_csv.config(state="disabled")
        btn_rev.config(state="disabled")
        for i in tree.get_children():
            tree.delete(i)
        prog["value"] = 0
        log("──────── Iniciando vinculación ────────")

        def worker():
            try:
                df_a, df_b = state["df_a"], state["df_b"]
                na, ia, aa = state["name_a"].get(), state["id_a"].get() or None, state["amt_a"].get() or None
                nb, ib, ab = state["name_b"].get(), state["id_b"].get() or None, state["amt_b"].get() or None
                q.put(("log", "Separando registros por persona (Archivo A)…"))
                recs_a = parse_blocks(df_a, na, ia, aa)
                q.put(("log", f"   {len(recs_a)} personas en A."))
                q.put(("log", "Separando registros por persona (Archivo B)…"))
                recs_b = parse_blocks(df_b, nb, ib, ab)
                q.put(("log", f"   {len(recs_b)} personas en B."))
                q.put(("log", "Buscando coincidencias (ID › Nombre+Monto)…"))

                def pcb(done, total, name):
                    q.put(("prog", (done, total, name)))

                rows, b_only = link(recs_a, recs_b, state["threshold"].get(), pcb)
                cols_a = [c for c in df_a.columns]
                cols_b = [c for c in df_b.columns]
                out = build_output(rows, b_only, cols_a, cols_b)
                q.put(("done", (rows, b_only, out, cols_a, cols_b)))
            except Exception as e:
                q.put(("error", str(e)))

        threading.Thread(target=worker, daemon=True).start()
        app.after(60, poll_queue)

    def poll_queue():
        try:
            while True:
                kind, payload = q.get_nowait()
                if kind == "log":
                    log(payload)
                elif kind == "prog":
                    done, total, name = payload
                    prog["value"] = 100 * done / max(total, 1)
                    prog_lbl.config(text=f"Procesando {done}/{total} · {name[:30]}")
                elif kind == "error":
                    log(f"✗ ERROR: {payload}")
                    messagebox.showerror("Error", payload)
                    btn_run.config(state="normal")
                    return
                elif kind == "done":
                    finish(*payload)
                    return
        except queue.Empty:
            pass
        app.after(60, poll_queue)

    def finish(rows, b_only, out, cols_a, cols_b):
        state["result_rows"] = rows
        state["out_table"] = out
        state["cols_a"] = cols_a
        state["cols_b"] = cols_b
        prog["value"] = 100
        prog_lbl.config(text="Completado.")

        total = len(rows)
        auto = sum(1 for r in rows if r["status"] == "AUTOMATICO")
        manual = total - auto
        by_id = sum(1 for r in rows if r["method"] == "ID")
        by_name = sum(1 for r in rows if r["method"] == "NOMBRE")
        none = sum(1 for r in rows if r["method"] == "SIN MATCH")
        pct = 100 * auto / total if total else 0

        card_vals["auto"].config(text=str(auto))
        card_vals["manual"].config(text=str(manual + len(b_only)))
        card_vals["pct"].config(text=f"{pct:.0f}%")
        card_vals["total"].config(text=str(total))

        for r in rows:
            tag = "auto" if r["status"] == "AUTOMATICO" else (
                "none" if r["method"] == "SIN MATCH" else "manual")
            tree.insert("", "end", tags=(tag,), values=(
                r["ra"]["_name"], r["ra"]["_eid"], r["method"],
                f'{r["conf"]:.1f}', f'{r["margin"]:.1f}', r["status"]))
        for rb in b_only:
            tree.insert("", "end", tags=("none",), values=(
                rb["_name"] + "  (solo en B)", rb["_eid"], "SIN MATCH",
                "0.0", "100.0", "REVISAR MANUAL"))

        log("──────── Resultado ────────")
        log(f"Match por ID: {by_id} · por Nombre: {by_name} · sin match: {none}")
        log(f"AUTOMÁTICOS: {auto} ({pct:.1f}%) · MANUAL: {manual + len(b_only)} "
            f"({100 - pct:.1f}% + {len(b_only)} solo-en-B)")
        log("Revisa los marcados en ámbar/rojo y exporta el CSV.")
        btn_run.config(state="normal")
        btn_csv.config(state="normal")
        btn_rev.config(state="normal")

    def export_csv(only_review):
        if not state["out_table"]:
            return
        default = ("revisar_manual" if only_review else "vinculado_consolidado")
        default += f"_{dt.datetime.now():%Y%m%d_%H%M}.csv"
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", initialfile=default,
            filetypes=[("CSV", "*.csv")])
        if not path:
            return
        rows = state["out_table"]
        if only_review:
            rows = [r for r in rows if r["Estado"] == "REVISAR MANUAL"]
        if not rows:
            messagebox.showinfo("Sin filas", "No hay filas que exportar.")
            return
        fields = list(rows[0].keys())
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            w.writerows(rows)
        log(f"✓ Exportado: {os.path.basename(path)} ({len(rows)} filas)")
        messagebox.showinfo("Exportado", f"Guardado:\n{path}\n\n{len(rows)} filas.")

    app.mainloop()


if __name__ == "__main__":
    run_gui()
