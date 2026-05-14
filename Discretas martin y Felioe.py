import os
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import openpyxl

# COLORES Y FUENTES
COLOR_BG        = "#0D1B2A"
COLOR_PANEL     = "#1A2E44"
COLOR_CARD      = "#1F3A55"
COLOR_ACCENT    = "#F4C430"
COLOR_ACCENT2   = "#CE1126"
COLOR_TEXT      = "#E8EFF5"
COLOR_MUTED     = "#7B9AB0"
COLOR_SUCCESS   = "#27AE60"
COLOR_BORDER    = "#2C4A6B"
COLOR_HIGHLIGHT = "#2563EB"

FONT_TITLE  = ("Trebuchet MS", 20, "bold")
FONT_HEADER = ("Trebuchet MS", 13, "bold")
FONT_LABEL  = ("Trebuchet MS", 11)
FONT_SMALL  = ("Trebuchet MS", 9)
FONT_BTN    = ("Trebuchet MS", 11, "bold")
FONT_MONO   = ("Courier New", 10)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# CLASE GRAFO
class Grafo:
    def __init__(self):
        self.adyacencia = {}
        self.nombres = []

    def agregar_vertice(self, ciudad):
        if ciudad not in self.adyacencia:
            self.adyacencia[ciudad] = []
            self.nombres.append(ciudad)

    def agregar_arista(self, origen, destino, peso):
        ya_existe = any(v == destino for v, _ in self.adyacencia[origen])
        if not ya_existe:
            self.adyacencia[origen].append((destino, peso))
            self.adyacencia[destino].append((origen, peso))


# LÓGICA DE GRAFOS Y ALGORITMOS
def construir_grafo(ruta_excel):
    wb = openpyxl.load_workbook(ruta_excel, read_only=True, data_only=True)
    ws = wb.active
    filas = list(ws.iter_rows(values_only=True))
    datos_filas = [f for f in filas[3:] if f[1] is not None and str(f[1]).strip() != ""]
    n = len(datos_filas)
    ciudades = [str(f[1]).strip().replace("\n", " ").replace("\r", " ") for f in datos_filas]
    matriz = []
    for fila in datos_filas:
        fila_dist = []
        for v in fila[2: 2 + n]:
            try:
                fila_dist.append(int(v))
            except (TypeError, ValueError):
                fila_dist.append(0)
        matriz.append(fila_dist)
    wb.close()
    g = Grafo()
    for ciudad in ciudades:
        g.agregar_vertice(ciudad)
    for i in range(n):
        for j in range(i + 1, n):
            peso = matriz[i][j]
            if peso > 0:
                g.agregar_arista(ciudades[i], ciudades[j], peso)
    return g


def dijkstra(g, origen):
    distancias = {}
    for v in g.adyacencia:
        distancias[v] = float('inf')
    predecesores = {}
    for v in g.adyacencia:
        predecesores[v] = None
    visitados    = {}
    distancias[origen] = 0
    while True:
        actual, minimo = None, float('inf')
        for v in distancias:
            if v not in visitados and distancias[v] < minimo:
                minimo, actual = distancias[v], v
        if actual is None:
            break
        visitados[actual] = True
        for vecino, peso in g.adyacencia[actual]:
            if vecino not in visitados:
                nd = distancias[actual] + peso
                if nd < distancias[vecino]:
                    distancias[vecino]   = nd
                    predecesores[vecino] = actual
    return distancias, predecesores


def dijkstra_todos(g):
    todas_dist, todos_pred = {}, {}
    for ciudad in g.adyacencia:
        d, p = dijkstra(g, ciudad)
        todas_dist[ciudad] = d
        todos_pred[ciudad] = p
    return todas_dist, todos_pred


def reconstruir_camino(predecesores, destino):
    camino, actual = [], destino
    while actual is not None:
        camino.append(actual)
        actual = predecesores[actual]
    camino.reverse()
    return camino


def centralidad_grado(g):
    n = len(g.adyacencia)
    centralidad = {}
    for c in g.adyacencia:
        grado = len(g.adyacencia[c])
        centralidad[c] = grado / (n - 1)
    return centralidad


def centralidad_intermediacion(g, todos_pred):
    bc = {}
    for c in g.adyacencia:
        bc[c] = 0
    for s in g.adyacencia:
        for t in g.adyacencia:
            if s == t:
                continue
            for nodo in reconstruir_camino(todos_pred[s], t)[1:-1]:
                bc[nodo] += 1
    return bc


def centralidad_cercania(g, todas_dist):
    n = len(g.adyacencia)
    cc = {}
    for c in g.adyacencia:
        suma = 0
        for d in todas_dist[c].values():
            if d != 0 and d != float('inf'):
                suma = suma + d
        if suma > 0:
            cc[c] = (n - 1) / suma
        else:
            cc[c] = 0.0
    return cc


# APLICACIÓN PRINCIPAL
class AppRutas(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Rutas Optimas — Colombia  |  Matematicas Discretas")
        self.configure(bg=COLOR_BG)
        self.geometry("1060x680")
        self.minsize(900, 560)

        self.g = None
        self.todas_dist = None
        self.todos_pred = None

        self._build_ui()
        self._try_autoload()

    def _try_autoload(self):
        ruta = os.path.join(BASE_DIR, "Datos.xlsx")
        if os.path.exists(ruta):
            self._load_graph(ruta)
        else:
            self.status_lbl.config(
                text="Datos.xlsx no encontrado — colócalo junto a este script.",
                fg=COLOR_ACCENT2)

    # BUILD UI
    def _build_ui(self):
        # cabecera
        header = tk.Frame(self, bg=COLOR_PANEL, pady=12)
        header.pack(fill="x")

        flag = tk.Frame(header, bg=COLOR_PANEL)
        flag.pack(side="left", padx=18)
        for col, w in [(COLOR_ACCENT2, 7), (COLOR_ACCENT, 7), (COLOR_ACCENT2, 7)]:
            tk.Frame(flag, bg=col, width=w, height=42).pack(side="left")

        tk.Label(header, text="RUTAS OPTIMAS", font=FONT_TITLE,bg=COLOR_PANEL, fg=COLOR_ACCENT).pack(side="left", padx=14)
        tk.Label(header, text="Colombia  |  Teoria de Grafos  |  Dijkstra  |  Centralidades",font=FONT_SMALL, bg=COLOR_PANEL, fg=COLOR_MUTED).pack(side="left")

        self.status_lbl = tk.Label(header, text="Cargando...", font=FONT_SMALL,bg=COLOR_PANEL, fg=COLOR_MUTED)
        self.status_lbl.pack(side="right", padx=20)

        # cuerpo
        body = tk.Frame(self, bg=COLOR_BG)
        body.pack(fill="both", expand=True, padx=14, pady=10)

        # sidebar
        sidebar = tk.Frame(body, bg=COLOR_PANEL, width=228)
        sidebar.pack(side="left", fill="y", padx=(0, 12))
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text="MENU", font=FONT_HEADER,bg=COLOR_PANEL, fg=COLOR_ACCENT).pack(pady=(18, 4))
        self._sep(sidebar)

        btn_defs = [
            ("Ruta mas corta",                self._show_dijkstra),
            ("Centralidad de Grado",          self._show_grado),
            ("Centralidad de Intermediacion", self._show_intermediacion),
            ("Centralidad de Cercania",       self._show_cercania),
        ]
        self.menu_btns = []
        for txt, cmd in btn_defs:
            b = tk.Button(sidebar, text=txt, font=FONT_BTN,
                          bg=COLOR_CARD, fg=COLOR_ACCENT,
                          activebackground=COLOR_BORDER, activeforeground=COLOR_TEXT,
                          disabledforeground=COLOR_MUTED,
                          relief="flat", anchor="w", padx=12, pady=9,
                          width=22, command=cmd, state="disabled")
            b.pack(fill="x", padx=10, pady=3)
            self.menu_btns.append(b)

        self._sep(sidebar)
        self.info_ciudades = tk.Label(sidebar, text="Ciudades:    —",font=FONT_SMALL, bg=COLOR_PANEL, fg=COLOR_TEXT)
        self.info_ciudades.pack(anchor="w", padx=14, pady=2)
        self.info_aristas  = tk.Label(sidebar, text="Conexiones:  —", font=FONT_SMALL, bg=COLOR_PANEL, fg=COLOR_TEXT)
        self.info_aristas.pack(anchor="w", padx=14, pady=2)

        # panel de contenido
        self.content_frame = tk.Frame(body, bg=COLOR_BG)
        self.content_frame.pack(side="left", fill="both", expand=True)

        self._show_welcome()

    def _sep(self, parent):
        tk.Frame(parent, bg=COLOR_BORDER, height=1).pack(fill="x", padx=10, pady=5)

    # CARGA DEL GRAFO
    def _load_graph(self, path):
        self.status_lbl.config(text="Calculando rutas...", fg=COLOR_ACCENT)
        self.update_idletasks()

        def worker():
            try:
                g = construir_grafo(path)
                td, tp = dijkstra_todos(g)
                self.after(0, lambda: self._on_load_ok(g, td, tp))
            except Exception as e:
                self.after(0, lambda: self._on_load_err(str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_load_ok(self, g, td, tp):
        self.g, self.todas_dist, self.todos_pred = g, td, tp
        n = len(g.adyacencia)
        aris = sum(len(v) for v in g.adyacencia.values()) // 2
        self.status_lbl.config(
            text=f"Listo  —  {n} ciudades, {aris} conexiones", fg=COLOR_SUCCESS)
        self.info_ciudades.config(text=f"Ciudades:    {n}")
        self.info_aristas.config( text=f"Conexiones:  {aris}")
        for b in self.menu_btns:
            b.config(state="normal")
        self._show_welcome()

    def _on_load_err(self, msg):
        self.status_lbl.config(text="Error al cargar", fg=COLOR_ACCENT2)
        messagebox.showerror("Error", f"No se pudo cargar Datos.xlsx:\n{msg}")

    # PANTALLAS
    def _clear_content(self):
        for w in self.content_frame.winfo_children():
            w.destroy()

    def _show_welcome(self):
        self._clear_content()
        f = self.content_frame
        tk.Label(f, text="RUTAS OPTIMAS EN COLOMBIA",font=FONT_TITLE, bg=COLOR_BG, fg=COLOR_ACCENT).pack(pady=(60, 6))
        tk.Label(f, text="Proyecto Final  |  Matematicas Discretas",font=FONT_LABEL, bg=COLOR_BG, fg=COLOR_MUTED).pack()
        tk.Frame(f, bg=COLOR_BORDER, height=1, width=420).pack(pady=14)
        if self.g:
            msg = "Grafo cargado. Selecciona una opcion del menu lateral."
            col = COLOR_SUCCESS
        else:
            msg = "Coloca Datos.xlsx en la misma carpeta que este script."
            col = COLOR_MUTED
        tk.Label(f, text=msg, font=FONT_LABEL, bg=COLOR_BG, fg=col).pack()

    #Ruta mas corta 
    def _show_dijkstra(self):
        self._clear_content()
        f = self.content_frame

        tk.Label(f, text="Ruta Mas Corta — Algoritmo de Dijkstra",font=FONT_HEADER, bg=COLOR_BG, fg=COLOR_ACCENT).pack(pady=(10, 4))

        sel = tk.Frame(f, bg=COLOR_PANEL, pady=12, padx=18)
        sel.pack(fill="x", padx=8, pady=4)

        ciudades = self.g.nombres

        tk.Label(sel, text="Ciudad ORIGEN:", font=FONT_LABEL, bg=COLOR_PANEL, fg=COLOR_MUTED).grid(row=0, column=0, sticky="w", padx=6)
        self.cb_origen = ttk.Combobox(sel, values=ciudades, state="readonly", font=FONT_LABEL, width=26)
        self.cb_origen.grid(row=0, column=1, padx=8, pady=3)
        self.cb_origen.current(0)

        tk.Label(sel, text="Ciudad DESTINO:", font=FONT_LABEL,bg=COLOR_PANEL, fg=COLOR_MUTED).grid(row=1, column=0, sticky="w", padx=6)
        self.cb_destino = ttk.Combobox(sel, values=ciudades, state="readonly",font=FONT_LABEL, width=26)
        self.cb_destino.grid(row=1, column=1, padx=8, pady=3)
        self.cb_destino.current(1 if len(ciudades) > 1 else 0)

        tk.Button(sel, text="Calcular Ruta", font=FONT_BTN,
                  bg=COLOR_HIGHLIGHT, fg=COLOR_ACCENT,
                  activebackground="#1d4ed8", activeforeground=COLOR_ACCENT,
                  relief="flat", padx=14, pady=6,
                  command=self._run_dijkstra).grid(row=0, column=2, rowspan=2, padx=14)

        self.dijk_result = tk.Frame(f, bg=COLOR_BG)
        self.dijk_result.pack(fill="both", expand=True, padx=8, pady=4)

    def _run_dijkstra(self):
        origen  = self.cb_origen.get()
        destino = self.cb_destino.get()

        for w in self.dijk_result.winfo_children():
            w.destroy()

        if origen == destino:
            tk.Label(self.dijk_result, text="Origen y destino son la misma ciudad.", font=FONT_LABEL, bg=COLOR_BG, fg=COLOR_ACCENT2).pack(pady=16)
            return

        distancias, predecesores = dijkstra(self.g, origen)

        if distancias[destino] == float('inf'):
            tk.Label(self.dijk_result,
                     text=f"No existe camino entre {origen} y {destino}.",
                     font=FONT_LABEL, bg=COLOR_BG, fg=COLOR_ACCENT2).pack(pady=16)
            return

        camino = reconstruir_camino(predecesores, destino)

        # resumen
        res = tk.Frame(self.dijk_result, bg=COLOR_CARD, pady=8, padx=16)
        res.pack(fill="x", pady=4)
        arrow_txt = "  ->  ".join(camino)
        tk.Label(res, text=arrow_txt, font=("Courier New", 9, "bold"),
                 bg=COLOR_CARD, fg=COLOR_TEXT, wraplength=720, justify="left").pack(anchor="w")
        row = tk.Frame(res, bg=COLOR_CARD)
        row.pack(anchor="w", pady=4)
        for lbl, val, col in [
            ("Distancia total", f"{distancias[destino]:,} km", COLOR_ACCENT),
            ("Segmentos",       str(len(camino) - 1),          COLOR_MUTED),
        ]:
            tk.Label(row, text=f"{lbl}:", font=FONT_SMALL,
                     bg=COLOR_CARD, fg=COLOR_MUTED).pack(side="left", padx=(0, 4))
            tk.Label(row, text=val, font=("Trebuchet MS", 11, "bold"),
                     bg=COLOR_CARD, fg=col).pack(side="left", padx=(0, 18))

        self._tabla_segmentos(self.dijk_result, camino)

    def _tabla_segmentos(self, parent, camino):
        tk.Label(parent, text="Detalle del trayecto", font=FONT_SMALL,
                 bg=COLOR_BG, fg=COLOR_MUTED).pack(anchor="w", padx=4, pady=(6, 2))
        wrapper = tk.Frame(parent, bg=COLOR_BG)
        wrapper.pack(fill="both", expand=True)
        cols = ("Desde", "Hacia", "Distancia (km)")
        tree = ttk.Treeview(wrapper, columns=cols, show="headings", height=14)
        self._style_tree(tree, cols)
        for i in range(len(camino) - 1):
            a, b = camino[i], camino[i + 1]
            dist_ab = next((p for v, p in self.g.adyacencia[a] if v == b), 0)
            tree.insert("", "end", values=(a, b, f"{dist_ab:,}"))
        sb = ttk.Scrollbar(wrapper, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    #Centralidad de Grado
    def _show_grado(self):
        self._clear_content()
        f = self.content_frame
        tk.Label(f, text="Centralidad de Grado",
                 font=FONT_HEADER, bg=COLOR_BG, fg=COLOR_SUCCESS).pack(pady=(10, 2))
        tk.Label(f,
                 text="Formula:  Cd(v) = deg(v) / (n - 1)   |   "
                      "Mide las conexiones directas de cada ciudad, normalizado sobre el maximo posible.",
                 font=FONT_SMALL, bg=COLOR_BG, fg=COLOR_MUTED, wraplength=800).pack()
        grado = centralidad_grado(self.g)
        pares = sorted(grado.items(), key=lambda x: x[1], reverse=True)
        self._tabla_centralidad(f, pares,
                                ("Posicion", "Ciudad", "Cd(v) = deg(v) / (n-1)"),
                                fmt_val=lambda v: f"{v:.6f}")

    #Centralidad de Intermediacion
    def _show_intermediacion(self):
        self._clear_content()
        f = self.content_frame
        tk.Label(f, text="Centralidad de Intermediacion",
                 font=FONT_HEADER, bg=COLOR_BG, fg=COLOR_ACCENT2).pack(pady=(10, 2))
        tk.Label(f,
                 text="Formula:  Cb(v) = SUM [ sigma(s,t|v) / sigma(s,t) ]     "
                      "Normalizacion:  Cb_norm(v) = Cb(v) / [ (n-1)(n-2) ]",
                 font=FONT_SMALL, bg=COLOR_BG, fg=COLOR_MUTED, wraplength=800).pack()

        n    = len(self.g.adyacencia)
        norm = (n - 1) * (n - 2)
        bc   = centralidad_intermediacion(self.g, self.todos_pred)
        pares = sorted(bc.items(), key=lambda x: x[1], reverse=True)

        cols = ("Posicion", "Ciudad",
                "Cb(v)  [rutas que pasan por v]",
                "Cb_norm(v) = Cb(v) / [(n-1)(n-2)]")
        wrapper = tk.Frame(f, bg=COLOR_BG)
        wrapper.pack(fill="both", expand=True, padx=8, pady=6)
        tree = ttk.Treeview(wrapper, columns=cols, show="headings", height=22)
        self._style_tree(tree, cols)
        tree.column(cols[0], width=80,  anchor="center")
        tree.column(cols[1], width=200, anchor="w")
        tree.column(cols[2], width=240, anchor="center")
        tree.column(cols[3], width=270, anchor="center")

        ranks = {0: "[1]", 1: "[2]", 2: "[3]"}
        for i, (ciudad, val) in enumerate(pares):
            pos = ranks.get(i, f"  {i+1}")
            tree.insert("", "end", values=(pos, ciudad, f"{val:,}", f"{val/norm:.6f}"))

        sb = ttk.Scrollbar(wrapper, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    #Centralidad de Cercania
    def _show_cercania(self):
        self._clear_content()
        f = self.content_frame
        tk.Label(f, text="Centralidad de Cercania",
                 font=FONT_HEADER, bg=COLOR_BG, fg="#8B5CF6").pack(pady=(10, 2))
        tk.Label(f,
                 text="Formula:  Cc(v) = (n - 1) / SUM [ d(v, u) ]   |   "
                      "Mayor valor indica que la ciudad esta en promedio mas cerca del resto de la red.",
                 font=FONT_SMALL, bg=COLOR_BG, fg=COLOR_MUTED, wraplength=800).pack()
        cc = centralidad_cercania(self.g, self.todas_dist)
        pares = sorted(cc.items(), key=lambda x: x[1], reverse=True)
        self._tabla_centralidad(f, pares,
                                ("Posicion", "Ciudad", "Cc(v) = (n-1) / SUM d(v,u)"),
                                fmt_val=lambda v: f"{v:.6f}")

    # HELPERS
    def _tabla_centralidad(self, parent, pares, cols, fmt_val):
        wrapper = tk.Frame(parent, bg=COLOR_BG)
        wrapper.pack(fill="both", expand=True, padx=8, pady=6)
        tree = ttk.Treeview(wrapper, columns=cols, show="headings", height=22)
        self._style_tree(tree, cols)
        tree.column(cols[0], width=80,  anchor="center")
        tree.column(cols[1], width=220, anchor="w")
        tree.column(cols[2], width=300, anchor="center")

        ranks = {0: "[1]", 1: "[2]", 2: "[3]"}
        for i, (ciudad, val) in enumerate(pares):
            pos = ranks.get(i, f"  {i+1}")
            tree.insert("", "end", values=(pos, ciudad, fmt_val(val)))

        sb = ttk.Scrollbar(wrapper, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _style_tree(self, tree, cols):
        s = ttk.Style()
        s.theme_use("default")
        s.configure("Treeview",
                    background=COLOR_CARD, foreground=COLOR_TEXT,
                    rowheight=26, fieldbackground=COLOR_CARD, font=FONT_MONO)
        s.configure("Treeview.Heading",
                    background=COLOR_PANEL, foreground=COLOR_ACCENT,
                    font=("Trebuchet MS", 10, "bold"))
        s.map("Treeview", background=[("selected", COLOR_HIGHLIGHT)])
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, anchor="center", width=180)


# PUNTO DE ENTRADA
if __name__ == "__main__":
    AppRutas().mainloop()