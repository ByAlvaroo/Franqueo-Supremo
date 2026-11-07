# -*- coding: utf-8 -*-
"""
FRANQUEO SUPREMO
Calcula qué sellos pegar para cada envío, gastando primero los de más valor,
y cuántos envíos puedes hacer con lo que tienes.

HECHO POR @BYALVARO__
"""
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
import urllib.request
import webbrowser
from tkinter import filedialog, messagebox, ttk

import motor as M

APP = "Franqueo Supremo"
NOMBRE_EXCEL = "Mis_Sellos.xlsx"
ICONO_B64 = "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"

# ---------------------------------------------------------------- colores ---
BG = "#F6F1E7"
PAPEL = "#FFFDF8"
CARD = "#FFFFFF"
LINEA = "#E4DCCB"
SIDE = "#1F2A44"
SIDE_HOV = "#2C3A5C"
SIDE_ACT = "#34466E"
ROJO = "#C8322F"
ROJO_OSC = "#A52623"
ORO = "#E0A526"
VERDE = "#2E8B57"
TEXTO = "#2B2B2B"
SUAVE = "#7A7060"
SEL_BG = "#FFF4F2"
F = "Segoe UI"


COLS = 4            # tarjetas por fila


def colocar(tarifas):
    """
    Posiciones en la rejilla: fila de título ORDINARIO, sus tarjetas (de 4 en 4),
    fila de título CERTIFICADO y sus tarjetas. Devuelve (posiciones, titulos).
    """
    pos, titulos = [], []
    fila = 0
    for tipo, txt in (("ord", "ORDINARIO"), ("cert", "CERTIFICADO")):
        idx = [i for i, t in enumerate(tarifas) if t["tipo"] == tipo]
        if not idx:
            continue
        titulos.append((fila, txt, tipo))
        fila += 1
        for k, _ in enumerate(idx):
            pos.append((fila + k // COLS, k % COLS))
        fila += (len(idx) + COLS - 1) // COLS
    return pos, titulos


def info_build():
    """Versión y repositorio de GitHub (los pone GitHub al crear el .exe)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    try:
        with open(os.path.join(base, "build_info.json"), encoding="utf-8-sig") as fh:
            d = json.load(fh)
        return {"version": str(d.get("version", "dev")), "repo": str(d.get("repo", ""))}
    except Exception:
        return {"version": "dev", "repo": ""}


def version_tupla(v):
    return tuple(int(x) for x in re.findall(r"\d+", v or ""))


def color_valor(cent):
    if cent < 30:
        return "#3E8E6E"
    if cent < 50:
        return "#2F7F9E"
    if cent < 100:
        return "#3B5BA9"
    if cent < 200:
        return "#7B4BA0"
    return ROJO


def carpeta_app():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def abrir(ruta):
    try:
        if sys.platform.startswith("win"):
            os.startfile(ruta)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", ruta])
        else:
            subprocess.Popen(["xdg-open", ruta])
    except Exception as e:
        messagebox.showerror(APP, f"No se pudo abrir:\n{ruta}\n\n{e}")


def parse_euros(txt):
    txt = (txt or "").replace("€", "").replace(",", ".").strip()
    v = float(txt)
    if v <= 0:
        raise ValueError
    return M.a_cent(v)


# ------------------------------------------------------------- widgets -----
class Boton(tk.Label):
    """Botón plano con hover."""
    ESTILOS = {
        "primario": (ROJO, "#FFFFFF", ROJO_OSC),
        "verde": (VERDE, "#FFFFFF", "#24734A"),
        "secundario": (SIDE, "#FFFFFF", SIDE_HOV),
        "suave": ("#EFE8DA", TEXTO, "#E4DAC6"),
        "fantasma": (CARD, SUAVE, "#F3EEE4"),
    }

    def __init__(self, master, texto, comando=None, estilo="primario", tam=10, padx=16, pady=8, **kw):
        bg, fg, hov = self.ESTILOS[estilo]
        super().__init__(master, text=texto, bg=bg, fg=fg, font=(F, tam, "bold"),
                         padx=padx, pady=pady, cursor="hand2", **kw)
        self._bg, self._fg, self._hov = bg, fg, hov
        self.comando = comando
        self.activo = True
        self.bind("<Enter>", lambda e: self.activo and self.config(bg=self._hov))
        self.bind("<Leave>", lambda e: self.activo and self.config(bg=self._bg))
        self.bind("<Button-1>", self._click)

    def _click(self, _e):
        if self.activo and self.comando:
            self.comando()

    def habilitar(self, si=True):
        self.activo = si
        self.config(bg=self._bg if si else "#D8D2C6", fg=self._fg if si else "#9A9284",
                    cursor="hand2" if si else "arrow")


def bind_arbol(widget, evento, fn):
    widget.bind(evento, fn)
    for h in widget.winfo_children():
        bind_arbol(h, evento, fn)


class Desplazable(tk.Frame):
    """Frame con scroll vertical (rueda del ratón incluida)."""

    def __init__(self, master, bg=BG, **kw):
        super().__init__(master, bg=bg, **kw)
        self.cv = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.sb = ttk.Scrollbar(self, orient="vertical", command=self.cv.yview)
        self.dentro = tk.Frame(self.cv, bg=bg)
        self.win = self.cv.create_window(0, 0, window=self.dentro, anchor="nw")
        self.cv.configure(yscrollcommand=self.sb.set)
        self.cv.pack(side="left", fill="both", expand=True)
        self.sb.pack(side="right", fill="y")
        self.dentro.bind("<Configure>", lambda e: self.cv.configure(scrollregion=self.cv.bbox("all")))
        self.cv.bind("<Configure>", lambda e: self.cv.itemconfigure(self.win, width=e.width))

    def rueda(self, e):
        if self.cv.yview() == (0.0, 1.0):
            return
        self.cv.yview_scroll(int(-e.delta / 120) or (-1 if e.delta > 0 else 1), "units")


def dibujar_sello(cv, x, y, w, h, cent, cantidad=None, fondo=CARD, apagado=False):
    col = "#BDB5A6" if apagado else color_valor(cent)
    cv.create_rectangle(x, y, x + w, y + h, fill=PAPEL, outline="")
    r = max(2, int(w / 18))
    for lado in ("h", "v"):
        largo = w if lado == "h" else h
        n = max(3, round(largo / (r * 3.2)))
        for i in range(n + 1):
            p = (x if lado == "h" else y) + i * largo / n
            if lado == "h":
                pts = ((p, y), (p, y + h))
            else:
                pts = ((x, p), (x + w, p))
            for cx, cy in pts:
                cv.create_oval(cx - r, cy - r, cx + r, cy + r, fill=fondo, outline="")
    m = r * 2 + 2
    cv.create_rectangle(x + m, y + m, x + w - m, y + h - m, fill=col, outline="")
    cv.create_rectangle(x + m + 3, y + m + 3, x + w - m - 3, y + h - m - 3, outline="#FFFFFF", width=1)
    if h >= 60:
        cv.create_text(x + w / 2, y + m + h * 0.13, text="FRANQUEO", fill="#FFFFFF",
                       font=(F, -max(7, int(h * 0.08)), "bold"))
    cv.create_text(x + w / 2, y + h * 0.52, text=M.euros_corto(cent), fill="#FFFFFF",
                   font=(F, -max(9, int(h * (0.24 if cent < 1000 else 0.2))), "bold"))
    if h >= 60:
        cv.create_text(x + w / 2, y + h * 0.75, text="€", fill="#FFFFFF", font=(F, -int(h * 0.11), "bold"))
    if cantidad:
        rb = max(12, int(h * 0.19))
        bx, by = x + w - 4, y + 4
        cv.create_oval(bx - rb, by - rb, bx + rb, by + rb, fill=ORO if cantidad > 1 else SIDE,
                       outline="#FFFFFF", width=3)
        cv.create_text(bx, by, text=f"×{cantidad}", fill="#FFFFFF", font=(F, -int(rb * 0.85), "bold"))


# ------------------------------------------------------------------ app ----
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP)
        self.k = max(1.0, self.winfo_fpixels("1i") / 96.0)
        self.geometry(f"{self.S(1200)}x{self.S(780)}")
        self.minsize(self.S(1000), self.S(660))
        self.configure(bg=BG)
        try:
            self._icono = tk.PhotoImage(data=ICONO_B64)
            self.iconphoto(True, self._icono)
        except Exception:
            pass
        self._estilo_ttk()

        self.build = info_build()
        self.ruta = os.path.join(carpeta_app(), NOMBRE_EXCEL)
        try:
            self.alm = M.Almacen(self.ruta)
        except PermissionError:
            messagebox.showerror(APP, "No puedo crear/leer el Excel. Si está abierto en Excel, ciérralo y vuelve a abrir la app.")
            raise SystemExit
        except Exception as e:
            if messagebox.askyesno(APP, f"El archivo {NOMBRE_EXCEL} no se pudo leer:\n{e}\n\n¿Lo renombro y creo uno nuevo?"):
                os.replace(self.ruta, self.ruta + ".roto.xlsx")
                self.alm = M.Almacen(self.ruta)
            else:
                raise SystemExit

        self.cola = queue.Queue()
        self.modo = tk.StringVar(value="altos")
        self.sel = None              # ('t', indice) o ('x', cent)
        self.alts = []
        self.alt_i = 0
        self.deshacer = []           # pila de (usado)
        self.stats = {}              # cent -> resultado maximo_envios
        self.gen_stats = 0
        self.gen_alt = 0
        self.pendiente = False
        self.det_tarifa = 0
        self._tstats = None
        self._stats_id = None

        self._construir()
        self.mostrar("enviar")
        self.refrescar_todo()
        self.bind_all("<MouseWheel>", self._rueda_global)
        self.nueva = None
        self.after(1500, self._borrar_exe_viejo)
        self.after(2500, lambda: self.buscar_actualizacion(False))
        self.after(80, self._bombear)
        self.after(2000, self._vigilar)

    def _rueda_global(self, e):
        w = self.winfo_containing(e.x_root, e.y_root)
        while w is not None:
            if isinstance(w, Desplazable):
                w.rueda(e)
                return
            w = w.master

    def S(self, px):
        return int(px * self.k)

    def _estilo_ttk(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TEXTO,
                     rowheight=self.S(30), font=(F, 10), borderwidth=0)
        st.configure("Treeview.Heading", background=SIDE, foreground="#FFFFFF", font=(F, 10, "bold"),
                     relief="flat", padding=6)
        st.map("Treeview.Heading", background=[("active", SIDE_HOV)])
        st.map("Treeview", background=[("selected", "#F3D9D6")], foreground=[("selected", TEXTO)])
        st.configure("Vertical.TScrollbar", background="#E4DAC6", troughcolor=BG, borderwidth=0, arrowsize=12)

    # ------------------------------------------------------------ marco ---
    def _construir(self):
        side = tk.Frame(self, bg=SIDE, width=self.S(236))
        side.pack(side="left", fill="y")
        side.pack_propagate(False)

        logo = tk.Frame(side, bg=SIDE)
        logo.pack(fill="x", padx=18, pady=(22, 18))
        cv = tk.Canvas(logo, width=self.S(46), height=self.S(54), bg=SIDE, highlightthickness=0)
        cv.pack(side="left")
        dibujar_sello(cv, 2, 2, self.S(42), self.S(50), 96, fondo=SIDE)
        tx = tk.Frame(logo, bg=SIDE)
        tx.pack(side="left", padx=10)
        tk.Label(tx, text="Franqueo", bg=SIDE, fg="#FFFFFF", font=(F, 15, "bold")).pack(anchor="w")
        tk.Label(tx, text="SUPREMO", bg=SIDE, fg=ORO, font=(F, 9, "bold")).pack(anchor="w")

        self.nav = {}
        for clave, icono, txt in (("enviar", "✉", "Enviar"), ("sellos", "▦", "Mis sellos"),
                                  ("cuantos", "∑", "¿Cuántos envíos?"), ("historial", "↺", "Historial"),
                                  ("ajustes", "⚙", "Ajustes")):
            b = tk.Label(side, text=f"  {icono}   {txt}", anchor="w", bg=SIDE, fg="#DCE3F2",
                         font=(F, 11), padx=16, pady=11, cursor="hand2")
            b.pack(fill="x", padx=10, pady=1)
            b.bind("<Button-1>", lambda e, c=clave: self.mostrar(c))
            b.bind("<Enter>", lambda e, w=b, c=clave: self.pagina != c and w.config(bg=SIDE_HOV))
            b.bind("<Leave>", lambda e, w=b, c=clave: self.pagina != c and w.config(bg=SIDE))
            self.nav[clave] = b

        pie = tk.Frame(side, bg=SIDE)
        pie.pack(side="bottom", fill="x", padx=18, pady=16)
        tk.Label(pie, text="HECHO POR @BYALVARO__", bg=SIDE, fg="#8090B0", font=(F, 8, "bold")).pack(anchor="w")
        tk.Label(pie, text=f"versión {self.build['version']}", bg=SIDE, fg="#56648A", font=(F, 8)).pack(anchor="w")

        caja = tk.Frame(side, bg=SIDE_HOV)
        caja.pack(side="bottom", fill="x", padx=14, pady=(0, 8))
        tk.Label(caja, text="TIENES AHORA", bg=SIDE_HOV, fg="#9AA8C7", font=(F, 8, "bold")).pack(anchor="w", padx=14, pady=(12, 0))
        self.lbl_tot_u = tk.Label(caja, text="", bg=SIDE_HOV, fg="#FFFFFF", font=(F, 20, "bold"))
        self.lbl_tot_u.pack(anchor="w", padx=14)
        self.lbl_tot_v = tk.Label(caja, text="", bg=SIDE_HOV, fg=ORO, font=(F, 11, "bold"))
        self.lbl_tot_v.pack(anchor="w", padx=14, pady=(0, 12))

        self.cuerpo = tk.Frame(self, bg=BG)
        self.cuerpo.pack(side="left", fill="both", expand=True)
        self.paginas = {
            "enviar": self._pag_enviar(),
            "sellos": self._pag_sellos(),
            "cuantos": self._pag_cuantos(),
            "historial": self._pag_historial(),
            "ajustes": self._pag_ajustes(),
        }
        self.toast = tk.Label(self, text="", bg=SIDE, fg="#FFFFFF", font=(F, 10, "bold"), padx=18, pady=10)
        self.banner = tk.Frame(self.cuerpo, bg=ORO)
        self.lbl_banner = tk.Label(self.banner, text="", bg=ORO, fg=TEXTO, font=(F, 10, "bold"))
        self.lbl_banner.pack(side="left", padx=(34, 12), pady=8)
        self.b_banner = Boton(self.banner, "Actualizar ahora", self.actualizar, "secundario", tam=9, pady=5)
        self.b_banner.pack(side="left")
        self.b_banner_x = Boton(self.banner, "Más tarde", lambda: self.banner.pack_forget(), "suave", tam=9, pady=5)
        self.b_banner_x.pack(side="left", padx=8)
        self._toast_id = None

    def mostrar(self, clave):
        self.pagina = clave
        for c, b in self.nav.items():
            b.config(bg=SIDE_ACT if c == clave else SIDE, fg="#FFFFFF" if c == clave else "#DCE3F2",
                     font=(F, 11, "bold") if c == clave else (F, 11))
        for c, p in self.paginas.items():
            if c == clave:
                p.pack(fill="both", expand=True)
            else:
                p.pack_forget()
        if clave == "historial":
            self._pintar_historial()

    def aviso(self, texto, color=SIDE, ms=3200):
        self.toast.config(text=texto, bg=color)
        self.toast.place(in_=self.cuerpo, relx=0.5, rely=1.0, y=-24, anchor="s")
        self.toast.lift()
        if self._toast_id:
            self.after_cancel(self._toast_id)
        self._toast_id = self.after(ms, self.toast.place_forget)

    def _cabecera(self, pag, titulo, sub):
        tk.Label(pag, text=titulo, bg=BG, fg=TEXTO, font=(F, 22, "bold")).pack(anchor="w", padx=34, pady=(26, 0))
        tk.Label(pag, text=sub, bg=BG, fg=SUAVE, font=(F, 10), justify="left").pack(anchor="w", padx=34, pady=(2, 14))

    # ---------------------------------------------------------- ENVIAR ---
    def _pag_enviar(self):
        pag = tk.Frame(self.cuerpo, bg=BG)
        top = tk.Frame(pag, bg=BG)
        top.pack(fill="x")
        izq = tk.Frame(top, bg=BG)
        izq.pack(side="left", fill="x", expand=True)
        self._cabecera(izq, "¿Qué envío vas a hacer?", "Pulsa una tarifa y te digo qué sellos pegar. El importe siempre es exacto.")

        # selector de modo
        seg = tk.Frame(top, bg="#E9E1D2")
        seg.pack(side="right", padx=34, pady=(30, 0), anchor="n")
        self.seg_btn = {}
        for val, txt in (("altos", "▲  Gastar los más altos"), ("envios", "✉  Sacar más envíos")):
            b = tk.Label(seg, text=txt, font=(F, 10, "bold"), padx=14, pady=8, cursor="hand2")
            b.pack(side="left", padx=3, pady=3)
            b.bind("<Button-1>", lambda e, v=val: self.cambiar_modo(v))
            self.seg_btn[val] = b
        self.lbl_modo = tk.Label(pag, text="", bg=BG, fg=SUAVE, font=(F, 9, "italic"))
        self.lbl_modo.place(relx=1.0, x=-36, y=self.S(78), anchor="ne")

        # tarjetas de tarifas
        self.grid_t = tk.Frame(pag, bg=BG)
        self.grid_t.pack(fill="x", padx=28)
        self.cards = []

        otro = tk.Frame(pag, bg=BG)
        otro.pack(fill="x", padx=34, pady=(4, 10))
        tk.Label(otro, text="¿Otro importe?", bg=BG, fg=SUAVE, font=(F, 10)).pack(side="left")
        self.ent_otro = tk.Entry(otro, width=8, font=(F, 11), relief="flat", bg=CARD,
                                 highlightthickness=1, highlightbackground=LINEA, highlightcolor=ROJO, justify="center")
        self.ent_otro.pack(side="left", padx=8, ipady=4)
        tk.Label(otro, text="€", bg=BG, fg=SUAVE, font=(F, 10)).pack(side="left")
        Boton(otro, "Calcular", self.calcular_otro, "suave", tam=9, pady=5).pack(side="left", padx=10)
        self.ent_otro.bind("<Return>", lambda e: self.calcular_otro())

        # resultado
        res = tk.Frame(pag, bg=CARD, highlightthickness=1, highlightbackground=LINEA)
        res.pack(fill="both", expand=True, padx=34, pady=(4, 26))
        cab = tk.Frame(res, bg=CARD)
        cab.pack(fill="x", padx=24, pady=(18, 0))
        self.lbl_res_t = tk.Label(cab, text="Elige una tarifa", bg=CARD, fg=TEXTO, font=(F, 16, "bold"))
        self.lbl_res_t.pack(side="left")
        self.lbl_opc = tk.Label(cab, text="", bg=CARD, fg=SUAVE, font=(F, 10))
        self.lbl_opc.pack(side="right")

        self.cv_res = tk.Canvas(res, bg=CARD, highlightthickness=0, height=self.S(150))
        self.cv_res.pack(fill="both", expand=True, padx=24, pady=10)
        self.cv_res.bind("<Configure>", lambda e: self._pintar_resultado())

        self.lbl_res_info = tk.Label(res, text="", bg=CARD, fg=TEXTO, font=(F, 12))
        self.lbl_res_info.pack(anchor="w", padx=24)

        acc = tk.Frame(res, bg=CARD)
        acc.pack(fill="x", padx=24, pady=(10, 20))
        self.b_usar = Boton(acc, "✔  Ya los he pegado — descontar", self.usar, "verde", tam=11, pady=10)
        self.b_usar.pack(side="left")
        self.b_otra = Boton(acc, "↻  Otra opción", self.otra_opcion, "suave", tam=10, pady=10)
        self.b_otra.pack(side="left", padx=10)
        self.b_desh = Boton(acc, "↶  Deshacer último envío", self.deshacer_ultimo, "fantasma", tam=10, pady=10)
        self.b_desh.pack(side="right")
        return pag

    def _construir_cards(self):
        for w in self.grid_t.winfo_children():
            w.destroy()
        self.cards = []
        for i in range(COLS):
            self.grid_t.columnconfigure(i, weight=1, uniform="t")
        pos, titulos = colocar(self.alm.tarifas)
        for fila, txt, tipo in titulos:
            tk.Label(self.grid_t, text=txt, bg=BG, fg=ROJO if tipo == "ord" else SIDE,
                     font=(F, 9, "bold")).grid(row=fila, column=0, columnspan=COLS, sticky="w", padx=8, pady=(4, 0))
        for i, t in enumerate(self.alm.tarifas):
            grande = t["tipo"] == "ord"
            fila, col = pos[i]
            c = tk.Frame(self.grid_t, bg=CARD, highlightthickness=2, highlightbackground=LINEA, cursor="hand2")
            c.grid(row=fila, column=col, sticky="nsew", padx=5, pady=4)
            arriba = tk.Frame(c, bg=CARD)
            arriba.pack(fill="x", padx=14, pady=(10 if grande else 8, 0))
            tk.Frame(arriba, bg=ROJO if grande else SIDE, width=4, height=self.S(16)).pack(side="left", anchor="n", pady=2)
            tk.Label(arriba, text=t["nombre"], bg=CARD, fg=SUAVE, font=(F, 9, "bold"),
                     wraplength=self.S(160), justify="left").pack(side="left", padx=8)
            precio = tk.Label(c, text=M.euros(t["cent"]), bg=CARD, fg=TEXTO,
                              font=(F, 24 if grande else 16, "bold"))
            precio.pack(anchor="w", padx=14, pady=(2 if grande else 0, 0))
            quedan = tk.Label(c, text="calculando…", bg=CARD, fg=SUAVE, font=(F, 9))
            quedan.pack(anchor="w", padx=14, pady=(0, 10 if grande else 8))
            bind_arbol(c, "<Button-1>", lambda e, k=i: self.elegir(("t", k)))
            self.cards.append({"frame": c, "quedan": quedan, "t": t})
        self._marcar_cards()

    def _marcar_cards(self):
        for i, c in enumerate(self.cards):
            sel = self.sel == ("t", i)
            c["frame"].config(highlightbackground=ROJO if sel else LINEA)
            bg = SEL_BG if sel else CARD
            for w in [c["frame"]] + self._desc(c["frame"]):
                if isinstance(w, tk.Label) and w.cget("bg") in (CARD, SEL_BG):
                    w.config(bg=bg)
                elif isinstance(w, tk.Frame) and w.cget("bg") in (CARD, SEL_BG):
                    w.config(bg=bg)
        for v, b in self.seg_btn.items():
            on = self.modo.get() == v
            b.config(bg=ROJO if on else "#E9E1D2", fg="#FFFFFF" if on else SUAVE)
        self.lbl_modo.config(text="Usa primero los sellos de más valor, de ahí para abajo" if self.modo.get() == "altos"
                             else "Elige la combinación que deja más sobres posibles de esa tarifa")

    def _desc(self, w):
        out = []
        for h in w.winfo_children():
            out.append(h)
            out.extend(self._desc(h))
        return out

    def cambiar_modo(self, v):
        self.modo.set(v)
        self._marcar_cards()
        if self.sel:
            self.calcular()

    def elegir(self, sel):
        self.sel = sel
        self._marcar_cards()
        self.calcular()

    def calcular_otro(self):
        try:
            cent = parse_euros(self.ent_otro.get())
        except ValueError:
            self.aviso("Escribe un importe válido, p. ej. 1,30", ROJO)
            return
        self.elegir(("x", cent))

    def _sel_info(self):
        if not self.sel:
            return None, None, None
        if self.sel[0] == "t":
            if self.sel[1] >= len(self.alm.tarifas):
                return None, None, None
            t = self.alm.tarifas[self.sel[1]]
            return t["cent"], f"{M.etiqueta(t)}  ·  {M.euros(t['cent'])}", M.etiqueta(t)
        return self.sel[1], f"Importe libre  ·  {M.euros(self.sel[1])}", "Importe libre"

    def calcular(self):
        cent, titulo, _ = self._sel_info()
        if cent is None:
            return
        self.gen_alt += 1
        gen = self.gen_alt
        self.lbl_res_t.config(text=titulo)
        self.lbl_opc.config(text="")
        self.lbl_res_info.config(text="Calculando…", fg=SUAVE)
        self.alts = []
        self._pintar_resultado()
        stock = dict(self.alm.stock)
        modo = self.modo.get()
        plan = self.stats.get(cent) if modo == "envios" else None

        def trabajo():
            return M.alternativas(cent, stock, modo, n=6, segundos=0.8, plan=plan)

        def listo(alts):
            if gen != self.gen_alt:
                return
            self.alts = alts
            self.alt_i = 0
            self._mostrar_alt()

        self._en_hilo(trabajo, listo)

    def _mostrar_alt(self):
        cent, _, _ = self._sel_info()
        if cent is None:
            return
        if not self.alts:
            self.lbl_res_info.config(text=f"✖  Con los sellos que tienes no hay forma de sumar exactamente {M.euros(cent)}.",
                                     fg=ROJO)
            self.lbl_opc.config(text="")
            self.b_usar.habilitar(False)
            self.b_otra.habilitar(False)
        else:
            r = self.alts[self.alt_i]
            n = M.num_sellos(r)
            self.lbl_res_info.config(
                text=f"✔  {n} sello{'s' if n != 1 else ''}   ·   {M.texto_reparto(r)}   =   {M.euros(cent)} exacto",
                fg=VERDE)
            self.lbl_opc.config(text=f"Opción {self.alt_i + 1} de {len(self.alts)}" +
                                ("  ·  recomendada" if self.alt_i == 0 else ""))
            self.b_usar.habilitar(True)
            self.b_otra.habilitar(len(self.alts) > 1)
        self._pintar_resultado()

    def _pintar_resultado(self):
        cv = self.cv_res
        cv.delete("all")
        W = max(cv.winfo_width(), 200)
        if not self.sel:
            cv.create_text(W / 2, self.S(70), text="Pulsa una tarifa de arriba para ver qué sellos pegar",
                           fill="#B5AC9A", font=(F, 13))
            self.b_usar.habilitar(False)
            self.b_otra.habilitar(False)
            self.b_desh.habilitar(bool(self.deshacer))
            return
        self.b_desh.habilitar(bool(self.deshacer))
        if not self.alts:
            return
        r = self.alts[self.alt_i]
        w, h, gap = self.S(92), self.S(112), self.S(30)
        x, y = 8, self.S(16)
        filas = 1
        for v in sorted(r, reverse=True):
            if x + w > W - 8:
                x = 8
                y += h + self.S(24)
                filas += 1
            dibujar_sello(cv, x, y, w, h, v, r[v])
            cv.create_text(x + w / 2, y + h + self.S(11), text=f"{r[v]} × {M.euros(v)}", fill=SUAVE, font=(F, 9, "bold"))
            x += w + gap
        alto = filas * (h + self.S(24)) + self.S(20)
        if int(float(cv.cget("height"))) != alto:
            cv.config(height=alto)

    def otra_opcion(self):
        if len(self.alts) > 1:
            self.alt_i = (self.alt_i + 1) % len(self.alts)
            self._mostrar_alt()

    def usar(self):
        if not self.alts:
            return
        cent, _, etiqueta = self._sel_info()
        r = self.alts[self.alt_i]
        try:
            self.alm.descontar(r, etiqueta, cent)
        except ValueError as e:
            self.aviso(str(e), ROJO)
            return
        self.deshacer.append(r)
        self.guardar()
        self.aviso(f"Descontado: {M.texto_reparto(r)}", VERDE)
        self.refrescar_todo()

    def deshacer_ultimo(self):
        if not self.deshacer:
            return
        r = self.deshacer.pop()
        self.alm.devolver(r)
        self.guardar()
        self.aviso("Último envío deshecho: sellos devueltos", SIDE)
        self.refrescar_todo()

    # ---------------------------------------------------------- SELLOS ---
    def _pag_sellos(self):
        pag = tk.Frame(self.cuerpo, bg=BG)
        top = tk.Frame(pag, bg=BG)
        top.pack(fill="x")
        izq = tk.Frame(top, bg=BG)
        izq.pack(side="left")
        self._cabecera(izq, "Mis sellos", "Suma o resta lo que vayas gastando. Se guarda solo en el Excel y todo se recalcula.")
        self.lbl_sellos_tot = tk.Label(top, text="", bg=BG, fg=TEXTO, font=(F, 12, "bold"))
        self.lbl_sellos_tot.pack(side="right", padx=36, pady=(36, 0), anchor="n")

        add = tk.Frame(pag, bg=CARD, highlightthickness=1, highlightbackground=LINEA)
        add.pack(fill="x", padx=34, pady=(0, 12))
        dentro = tk.Frame(add, bg=CARD)
        dentro.pack(fill="x", padx=18, pady=14)
        tk.Label(dentro, text="＋ Añadir sellos", bg=CARD, fg=TEXTO, font=(F, 12, "bold")).pack(side="left", padx=(0, 18))

        def entrada(txt, ancho):
            tk.Label(dentro, text=txt, bg=CARD, fg=SUAVE, font=(F, 10)).pack(side="left")
            e = tk.Entry(dentro, width=ancho, font=(F, 11), relief="flat", bg=BG, justify="center",
                         highlightthickness=1, highlightbackground=LINEA, highlightcolor=ROJO)
            e.pack(side="left", padx=(6, 16), ipady=4)
            return e

        self.ent_val = entrada("Valor (€)", 8)
        self.ent_cant = entrada("Cantidad", 6)
        Boton(dentro, "Añadir", self.anadir, "primario", tam=10, pady=6).pack(side="left")
        tk.Label(dentro, text="Si ese valor ya existe, se suma a lo que tienes.", bg=CARD, fg=SUAVE,
                 font=(F, 9, "italic")).pack(side="left", padx=14)
        self.ent_cant.bind("<Return>", lambda e: self.anadir())
        self.ent_val.bind("<Return>", lambda e: self.ent_cant.focus_set())

        cab = tk.Frame(pag, bg=BG)
        cab.pack(fill="x", padx=34)
        for txt, w in (("Sello", 9), ("Valor", 10), ("Cantidad", 10), ("Subtotal", 11), ("Ajustar   ·   escribe 12 para fijar, +5 para sumar, -3 para restar", 0)):
            kw = {"width": w} if w else {}
            tk.Label(cab, text=txt.upper(), bg=BG, fg=SUAVE, font=(F, 8, "bold"), anchor="w",
                     **kw).pack(side="left", padx=(14 if txt == "Sello" else 4, 0))

        self.lista = Desplazable(pag)
        self.lista.pack(fill="both", expand=True, padx=34, pady=(4, 24))
        return pag

    def _pintar_sellos(self):
        try:
            y0 = self.lista.cv.yview()[0]
        except Exception:
            y0 = 0
        for w in self.lista.dentro.winfo_children():
            w.destroy()
        u, v = self.alm.total_unidades(), self.alm.total_valor()
        self.lbl_sellos_tot.config(text=f"{len(self.alm.stock)} tipos  ·  {u} sellos  ·  {M.euros(v)}")
        for i, val in enumerate(sorted(self.alm.stock, reverse=True)):
            cant = self.alm.stock[val]
            bg = CARD if i % 2 == 0 else "#FBF8F2"
            f = tk.Frame(self.lista.dentro, bg=bg)
            f.pack(fill="x", pady=1)
            cv = tk.Canvas(f, width=self.S(40), height=self.S(46), bg=bg, highlightthickness=0)
            cv.pack(side="left", padx=(12, 14), pady=5)
            dibujar_sello(cv, 2, 2, self.S(36), self.S(42), val, fondo=bg, apagado=cant == 0)
            tk.Label(f, text=M.euros(val), bg=bg, fg=TEXTO, font=(F, 12, "bold"), width=8, anchor="w").pack(side="left")
            tk.Label(f, text=str(cant), bg=bg, fg=ROJO if cant == 0 else TEXTO, font=(F, 16, "bold"),
                     width=6, anchor="w").pack(side="left", padx=(8, 0))
            tk.Label(f, text=M.euros(val * cant), bg=bg, fg=SUAVE, font=(F, 10), width=10, anchor="w").pack(side="left")
            Boton(f, "−", lambda v=val: self.ajustar(v, -1), "suave", tam=12, padx=12, pady=2).pack(side="left", padx=(10, 3))
            Boton(f, "+", lambda v=val: self.ajustar(v, +1), "suave", tam=12, padx=12, pady=2).pack(side="left", padx=3)
            e = tk.Entry(f, width=6, font=(F, 11), relief="flat", bg=PAPEL, justify="center",
                         highlightthickness=1, highlightbackground=LINEA, highlightcolor=ROJO)
            e.pack(side="left", padx=(12, 4), ipady=3)
            ok = Boton(f, "OK", None, "secundario", tam=9, padx=10, pady=4)
            ok.pack(side="left")
            ok.comando = (lambda v=val, en=e: self.ajustar_texto(v, en.get()))
            e.bind("<Return>", lambda ev, v=val, en=e: self.ajustar_texto(v, en.get()))
            Boton(f, "✕", lambda v=val: self.borrar_tipo(v), "fantasma", tam=10, padx=10, pady=2).pack(side="right", padx=10)
            for w in f.winfo_children():
                if isinstance(w, Boton) and w._bg == CARD:
                    w._bg = bg
                    w.config(bg=bg)
        self.lista.update_idletasks()
        self.lista.cv.yview_moveto(y0)

    def anadir(self):
        try:
            val = parse_euros(self.ent_val.get())
            cant = int(self.ent_cant.get().strip() or "0")
            if cant < 0:
                raise ValueError
        except ValueError:
            self.aviso("Pon un valor (p. ej. 0,45) y una cantidad entera", ROJO)
            return
        self.alm.stock[val] = self.alm.stock.get(val, 0) + cant
        self.ent_val.delete(0, "end")
        self.ent_cant.delete(0, "end")
        self.ent_val.focus_set()
        self.guardar()
        self.aviso(f"Añadidos {cant} sellos de {M.euros(val)}", VERDE)
        self.refrescar_todo()

    def ajustar(self, val, delta):
        self.alm.stock[val] = max(0, self.alm.stock.get(val, 0) + delta)
        self.guardar()
        self.refrescar_todo()

    def ajustar_texto(self, val, txt):
        txt = txt.strip().replace(" ", "")
        try:
            if txt.startswith(("+", "-")):
                nuevo = self.alm.stock.get(val, 0) + int(txt)
            else:
                nuevo = int(txt)
        except ValueError:
            self.aviso("Escribe 12 para fijar, +5 para sumar o -3 para restar", ROJO)
            return
        self.alm.stock[val] = max(0, nuevo)
        self.guardar()
        self.aviso(f"{M.euros(val)}: ahora tienes {self.alm.stock[val]}", VERDE)
        self.refrescar_todo()

    def borrar_tipo(self, val):
        if messagebox.askyesno(APP, f"¿Quitar el sello de {M.euros(val)} de la lista?\n(Tienes {self.alm.stock.get(val, 0)})"):
            self.alm.stock.pop(val, None)
            self.guardar()
            self.refrescar_todo()

    # --------------------------------------------------------- CUÁNTOS ---
    def _pag_cuantos(self):
        pag = tk.Frame(self.cuerpo, bg=BG)
        self._cabecera(pag, "¿Cuántos envíos puedo hacer?",
                       "Cada tarifa calculada por separado, como si todos tus sellos fueran para ella.\n"
                       "Pulsa la que te interese para ver cómo repartirlos. Se recalcula solo cuando cambias tus sellos.")
        self.grid_c = tk.Frame(pag, bg=BG)
        self.grid_c.pack(fill="x", padx=28)
        self.ccards = []

        det = tk.Frame(pag, bg=CARD, highlightthickness=1, highlightbackground=LINEA)
        det.pack(fill="both", expand=True, padx=34, pady=(8, 24))
        cab = tk.Frame(det, bg=CARD)
        cab.pack(fill="x", padx=22, pady=(16, 4))
        self.lbl_det = tk.Label(cab, text="", bg=CARD, fg=TEXTO, font=(F, 15, "bold"), justify="left")
        self.lbl_det.pack(side="left")
        Boton(cab, "Usar este plan al enviar", self.plan_a_enviar, "primario", tam=9, pady=6).pack(side="right")
        self.b_fondo = Boton(cab, "⌕  Buscar más a fondo", self.buscar_fondo, "suave", tam=9, pady=6)
        self.b_fondo.pack(side="right", padx=8)
        self.lbl_det_sub = tk.Label(det, text="", bg=CARD, fg=SUAVE, font=(F, 10), justify="left", anchor="w")
        self.lbl_det_sub.pack(fill="x", padx=22)
        self.lista_plan = Desplazable(det, bg=CARD)
        self.lista_plan.pack(fill="both", expand=True, padx=16, pady=(8, 14))
        return pag

    def _construir_ccards(self):
        for w in self.grid_c.winfo_children():
            w.destroy()
        self.ccards = []
        for i in range(COLS):
            self.grid_c.columnconfigure(i, weight=1, uniform="c")
        pos, titulos = colocar(self.alm.tarifas)
        for fila_t, txt, tipo in titulos:
            tk.Label(self.grid_c, text=txt, bg=BG, fg=ROJO if tipo == "ord" else SIDE,
                     font=(F, 9, "bold")).grid(row=fila_t, column=0, columnspan=COLS, sticky="w", padx=8, pady=(2, 0))
        for i, t in enumerate(self.alm.tarifas):
            c = tk.Frame(self.grid_c, bg=CARD, highlightthickness=2, highlightbackground=LINEA, cursor="hand2")
            c.grid(row=pos[i][0], column=pos[i][1], sticky="nsew", padx=5, pady=4)
            fila = tk.Frame(c, bg=CARD)
            fila.pack(fill="x", padx=16, pady=(10, 0))
            tk.Label(fila, text=t["nombre"], bg=CARD, fg=SUAVE, font=(F, 9, "bold"),
                     wraplength=self.S(120), justify="left").pack(side="left")
            tk.Label(fila, text=M.euros(t["cent"]), bg=CARD, fg=TEXTO, font=(F, 9, "bold")).pack(side="right")
            fila2 = tk.Frame(c, bg=CARD)
            fila2.pack(fill="x", padx=16)
            num = tk.Label(fila2, text="…", bg=CARD, fg=ROJO if t["tipo"] == "ord" else SIDE, font=(F, 26, "bold"))
            num.pack(side="left")
            tk.Label(fila2, text="sobres", bg=CARD, fg=SUAVE, font=(F, 10)).pack(side="left", padx=6, pady=(14, 0))
            barra = tk.Canvas(c, height=self.S(6), bg="#EFE8DA", highlightthickness=0)
            barra.pack(fill="x", padx=16, pady=(0, 12))
            bind_arbol(c, "<Button-1>", lambda e, k=i: self.ver_detalle(k))
            self.ccards.append({"frame": c, "num": num, "barra": barra})

    def ver_detalle(self, i):
        self.det_tarifa = i
        self._pintar_cuantos()

    def _pintar_cuantos(self):
        maxn = max([len(self.stats[t["cent"]]["sobres"]) for t in self.alm.tarifas if t["cent"] in self.stats] or [1]) or 1
        for i, (c, t) in enumerate(zip(self.ccards, self.alm.tarifas)):
            c["frame"].config(highlightbackground=ROJO if i == self.det_tarifa else LINEA)
            r = self.stats.get(t["cent"])
            c["num"].config(text=str(len(r["sobres"])) if r else "…")
            b = c["barra"]
            b.delete("all")
            if r:
                b.update_idletasks()
                wb = max(b.winfo_width(), 50)
                b.create_rectangle(0, 0, wb * len(r["sobres"]) / maxn, self.S(6), fill=ROJO if t["tipo"] == "ord" else SIDE, outline="")
        for w in self.lista_plan.dentro.winfo_children():
            w.destroy()
        if self.det_tarifa >= len(self.alm.tarifas):
            self.det_tarifa = 0
        if not self.alm.tarifas:
            return
        t = self.alm.tarifas[self.det_tarifa]
        r = self.stats.get(t["cent"])
        if not r:
            self.lbl_det.config(text=f"{M.etiqueta(t)} · {M.euros(t['cent'])}")
            self.lbl_det_sub.config(text="Calculando…")
            return
        n = len(r["sobres"])
        self.lbl_det.config(text=f"{M.etiqueta(t)} · {M.euros(t['cent'])}  →  {n} sobres")
        sob = {v: k for v, k in r["sobran"].items() if k > 0}
        txt_s = (f"Te sobrarían {sum(sob.values())} sellos ({M.euros(sum(v * k for v, k in sob.items()))}) (los guardas para otras tarifas)."
                 if sob else "No te sobraría ningún sello.")
        tope = r["tope"]
        extra = "  ¡Es el máximo posible!" if n >= tope else f"  (Máximo teórico por valor: {tope}.)"
        self.lbl_det_sub.config(text=txt_s + extra)
        cont = self.lista_plan.dentro
        grupos = sorted(M.agrupar(r["sobres"]), key=lambda g: sorted(g[0], reverse=True), reverse=True)
        for j, (rep, veces) in enumerate(grupos):
            bg = CARD if j % 2 == 0 else "#FBF8F2"
            f = tk.Frame(cont, bg=bg)
            f.pack(fill="x")
            tk.Label(f, text=f"×{veces}", bg=bg, fg=ROJO, font=(F, 14, "bold"), width=5, anchor="e").pack(side="left", padx=(6, 12), pady=4)
            tk.Label(f, text="sobres con", bg=bg, fg=SUAVE, font=(F, 9)).pack(side="left")
            cv = tk.Canvas(f, height=self.S(40), bg=bg, highlightthickness=0, width=self.S(46) * len(rep) + 10)
            cv.pack(side="left", padx=10, pady=3)
            x = 2
            for v in sorted(rep, reverse=True):
                dibujar_sello(cv, x, 2, self.S(32), self.S(36), v, rep[v] if rep[v] > 1 else None, fondo=bg)
                x += self.S(46)
            tk.Label(f, text=M.texto_reparto(rep), bg=bg, fg=TEXTO, font=(F, 10)).pack(side="left", padx=6)

    def buscar_fondo(self):
        if not self.alm.tarifas:
            return
        t = self.alm.tarifas[self.det_tarifa]
        stock = dict(self.alm.stock)
        self.b_fondo.habilitar(False)
        self.b_fondo.config(text="Buscando (6 s)…")
        gen = self.gen_stats

        def trabajo():
            return M.maximo_envios(t["cent"], stock, segundos=6.0)

        def listo(r):
            self.b_fondo.habilitar(True)
            self.b_fondo.config(text="⌕  Buscar más a fondo")
            if gen != self.gen_stats:
                return
            ant = self.stats.get(t["cent"])
            if not ant or len(r["sobres"]) > len(ant["sobres"]):
                self.stats[t["cent"]] = r
                self.aviso(f"¡Encontrado un reparto mejor: {len(r['sobres'])} sobres!", VERDE)
            else:
                self.aviso("No hay forma de sacar más sobres que esos", SIDE)
            self._pintar_cuantos()
            self._pintar_quedan()

        self._en_hilo(trabajo, listo)

    def plan_a_enviar(self):
        self.modo.set("envios")
        self.mostrar("enviar")
        self.elegir(("t", self.det_tarifa))

    def _recalcular_stats(self):
        """Programa el recálculo (espera un poco por si sigues pulsando + / −)."""
        self.gen_stats += 1
        self.stats = {}
        if self._stats_id:
            self.after_cancel(self._stats_id)
        self._stats_id = self.after(450, self._lanzar_stats)

    def _lanzar_stats(self):
        """Recalcula en segundo plano cuántos sobres salen de cada tarifa."""
        self._stats_id = None
        self.gen_stats += 1
        gen = self.gen_stats
        self.stats = {}
        stock = dict(self.alm.stock)
        tarifas = [t["cent"] for t in self.alm.tarifas]

        def trabajo():
            out = {}
            for c in tarifas:
                if gen != self.gen_stats:
                    return None
                out[c] = M.maximo_envios(c, stock, segundos=0.7)
                self.cola.put((self._stat_parcial, (gen, c, out[c])))
            return out

        self._en_hilo(trabajo, lambda r: None)

    def _stat_parcial(self, datos):
        gen, c, r = datos
        if gen != self.gen_stats:
            return
        self.stats[c] = r
        self._pintar_quedan()
        self._pintar_cuantos()

    def _pintar_quedan(self):
        for c in self.cards:
            r = self.stats.get(c["t"]["cent"])
            if r is None:
                c["quedan"].config(text="calculando…", fg=SUAVE)
            else:
                n = len(r["sobres"])
                c["quedan"].config(text=f"Te da para {n} sobres" if n else "Sin combinación exacta",
                                   fg=VERDE if n else ROJO)

    # ------------------------------------------------------- HISTORIAL ---
    def _pag_historial(self):
        pag = tk.Frame(self.cuerpo, bg=BG)
        self._cabecera(pag, "Historial de envíos", "Cada vez que descuentas sellos queda apuntado aquí y en la hoja «Historial» del Excel.")
        marco = tk.Frame(pag, bg=CARD, highlightthickness=1, highlightbackground=LINEA)
        marco.pack(fill="both", expand=True, padx=34, pady=(0, 24))
        cols = ("fecha", "tarifa", "importe", "sellos")
        self.tree = ttk.Treeview(marco, columns=cols, show="headings")
        for c, txt, w in (("fecha", "Fecha", 150), ("tarifa", "Tarifa", 200), ("importe", "Importe", 90), ("sellos", "Sellos usados", 460)):
            self.tree.heading(c, text=txt)
            self.tree.column(c, width=self.S(w), anchor="w" if c in ("tarifa", "sellos") else "center")
        sb = ttk.Scrollbar(marco, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        sb.pack(side="right", fill="y", pady=10)
        return pag

    def _pintar_historial(self):
        self.tree.delete(*self.tree.get_children())
        for fe, t, imp, txt in self.alm.historial:
            self.tree.insert("", "end", values=(fe, t, M.euros(imp), txt))

    # --------------------------------------------------------- AJUSTES ---
    def _pag_ajustes(self):
        pag = tk.Frame(self.cuerpo, bg=BG)
        self._cabecera(pag, "Ajustes", "Tus tarifas (añade las que quieras), tu Excel y las actualizaciones.")
        scroll = Desplazable(pag)
        scroll.pack(fill="both", expand=True, padx=(34, 20), pady=(0, 16))
        cont = scroll.dentro

        tar = tk.Frame(cont, bg=CARD, highlightthickness=1, highlightbackground=LINEA)
        tar.pack(fill="x", pady=(0, 14), padx=(0, 10))
        cab = tk.Frame(tar, bg=CARD)
        cab.pack(fill="x", padx=22, pady=(18, 4))
        tk.Label(cab, text="Tarifas de envío", bg=CARD, fg=TEXTO, font=(F, 14, "bold")).pack(side="left")
        Boton(cab, "Guardar tarifas", self.guardar_tarifas, "primario", tam=10, pady=6).pack(side="right")
        tk.Label(tar, text="Dos tablas separadas, igual que en el Excel. Pulsa «＋ Añadir» para crear más.",
                 bg=CARD, fg=SUAVE, font=(F, 9)).pack(anchor="w", padx=22, pady=(0, 10))
        cols = tk.Frame(tar, bg=CARD)
        cols.pack(fill="x", padx=22, pady=(0, 18))
        cols.columnconfigure(0, weight=1, uniform="a")
        cols.columnconfigure(1, weight=1, uniform="a")
        self.ed_tar = {}
        self.ed_marco = {}
        for k, (tipo, titulo, color) in enumerate((("ord", "ORDINARIO", ROJO), ("cert", "CERTIFICADO", SIDE))):
            col = tk.Frame(cols, bg=CARD)
            col.grid(row=0, column=k, sticky="new", padx=(0, 14) if k == 0 else (14, 0))
            tk.Label(col, text=titulo, bg=color, fg="#FFFFFF", font=(F, 10, "bold"), padx=10, pady=5,
                     anchor="w").pack(fill="x")
            enc = tk.Frame(col, bg=CARD)
            enc.pack(fill="x", pady=(6, 2))
            tk.Label(enc, text="NOMBRE", bg=CARD, fg=SUAVE, font=(F, 8, "bold")).pack(side="left")
            tk.Label(enc, text="PRECIO (€)", bg=CARD, fg=SUAVE, font=(F, 8, "bold")).pack(side="right", padx=(0, 46))
            marco = tk.Frame(col, bg=CARD)
            marco.pack(fill="x")
            self.ed_marco[tipo] = marco
            self.ed_tar[tipo] = []
            Boton(col, "＋ Añadir", lambda tp=tipo: self._fila_tarifa(tp, "", ""), "suave", tam=9,
                  pady=4).pack(anchor="w", pady=(6, 0))

        ex = tk.Frame(cont, bg=CARD, highlightthickness=1, highlightbackground=LINEA)
        ex.pack(fill="x", pady=(0, 14), padx=(0, 10))
        tk.Label(ex, text="Tu Excel", bg=CARD, fg=TEXTO, font=(F, 14, "bold")).pack(anchor="w", padx=22, pady=(18, 2))
        tk.Label(ex, text=("Sellos: columna B = cantidad, columna C = valor (como el tuyo).\n"
                           "Tarifas: tablas ORDINARIO (columnas F-G) y CERTIFICADO (columnas I-J); añade filas debajo.\n"
                           "Puedes editarlo a mano: guarda en Excel y la app lo recarga sola.\n"
                           "Mientras lo tengas abierto en Excel la app no puede escribir en él."),
                 bg=CARD, fg=SUAVE, font=(F, 9), justify="left").pack(anchor="w", padx=22)
        tk.Label(ex, text=self.ruta, bg=BG, fg=TEXTO, font=("Consolas", 9), wraplength=self.S(700),
                 justify="left", padx=10, pady=8).pack(anchor="w", padx=22, pady=12, fill="x")
        bot = tk.Frame(ex, bg=CARD)
        bot.pack(fill="x", padx=22, pady=(0, 18))
        for txt, fn, est in (("▤  Abrir el Excel", lambda: abrir(self.ruta), "secundario"),
                             ("❐  Abrir la carpeta", lambda: abrir(os.path.dirname(self.ruta)), "suave"),
                             ("⟳  Recargar", self.recargar, "suave"),
                             ("⇪  Importar otro Excel…", self.importar, "suave")):
            Boton(bot, txt, fn, est, tam=10, pady=8).pack(side="left", padx=(0, 8))

        up = tk.Frame(cont, bg=CARD, highlightthickness=1, highlightbackground=LINEA)
        up.pack(fill="x", pady=(0, 14), padx=(0, 10))
        tk.Label(up, text="Actualizaciones", bg=CARD, fg=TEXTO, font=(F, 14, "bold")).pack(anchor="w", padx=22, pady=(18, 2))
        repo = self.build["repo"] or "(sin GitHub: esta copia no se creó desde GitHub)"
        tk.Label(up, text=f"Versión instalada: {self.build['version']}     ·     GitHub: {repo}",
                 bg=CARD, fg=SUAVE, font=(F, 9)).pack(anchor="w", padx=22)
        tk.Label(up, text="Al abrir la app mira si hay una versión nueva en GitHub y te avisa arriba. "
                          "Tus sellos (el Excel) no se tocan.", bg=CARD, fg=SUAVE, font=(F, 9)).pack(anchor="w", padx=22)
        fila = tk.Frame(up, bg=CARD)
        fila.pack(fill="x", padx=22, pady=(10, 18))
        Boton(fila, "⟳  Buscar actualizaciones", lambda: self.buscar_actualizacion(True), "secundario",
              tam=10, pady=8).pack(side="left")
        self.lbl_up = tk.Label(fila, text="", bg=CARD, fg=SUAVE, font=(F, 9))
        self.lbl_up.pack(side="left", padx=12)
        return pag

    def _fila_tarifa(self, tipo, nombre, precio):
        f = tk.Frame(self.ed_marco[tipo], bg=CARD)
        f.pack(fill="x", pady=2)
        en = tk.Entry(f, font=(F, 11), relief="flat", bg=BG,
                      highlightthickness=1, highlightbackground=LINEA, highlightcolor=ROJO)
        en.pack(side="left", fill="x", expand=True, ipady=3)
        en.insert(0, nombre)
        ep = tk.Entry(f, width=8, font=(F, 11), relief="flat", bg=BG, justify="center",
                      highlightthickness=1, highlightbackground=LINEA, highlightcolor=ROJO)
        ep.pack(side="left", padx=8, ipady=3)
        ep.insert(0, precio)
        item = (f, en, ep)
        Boton(f, "✕", lambda: self._quitar_fila(tipo, item), "fantasma", tam=10, padx=8, pady=2).pack(side="left")
        self.ed_tar[tipo].append(item)
        if not nombre:
            en.focus_set()

    def _quitar_fila(self, tipo, item):
        if item in self.ed_tar[tipo]:
            self.ed_tar[tipo].remove(item)
            item[0].destroy()

    def _rellenar_ajustes(self):
        for tipo in ("ord", "cert"):
            for item in list(self.ed_tar[tipo]):
                item[0].destroy()
            self.ed_tar[tipo] = []
        for t in self.alm.tarifas:
            self._fila_tarifa(t["tipo"], t["nombre"], M.euros_corto(t["cent"]))

    def guardar_tarifas(self):
        nuevas = []
        for tipo, txt in (("ord", "ordinario"), ("cert", "certificado")):
            for k, (_f, en, ep) in enumerate(self.ed_tar[tipo], 1):
                if not ep.get().strip() and not en.get().strip():
                    continue
                try:
                    cent = parse_euros(ep.get())
                except ValueError:
                    self.aviso(f"Revisa el precio de la fila {k} de {txt}", ROJO)
                    return
                nuevas.append(M.tarifa(tipo, cent, en.get()))
        if not nuevas:
            self.aviso("Pon al menos una tarifa", ROJO)
            return
        self.alm.tarifas = nuevas
        self.sel = None
        self.det_tarifa = 0
        self.guardar()
        self.aviso("Tarifas guardadas", VERDE)
        self.refrescar_todo(reconstruir=True)

    def recargar(self):
        try:
            self.alm.cargar()
        except Exception as e:
            messagebox.showerror(APP, f"No se pudo leer el Excel:\n{e}")
            return
        self.deshacer.clear()
        self.aviso("Excel recargado", SIDE)
        self.refrescar_todo(reconstruir=True)

    def importar(self):
        ruta = filedialog.askopenfilename(title="Elige el Excel con tus sellos",
                                          filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if not ruta:
            return
        try:
            stock, tarifas, _ = M.Almacen.leer_libro(ruta)
        except Exception as e:
            messagebox.showerror(APP, f"No se pudo leer:\n{e}")
            return
        if not stock:
            messagebox.showwarning(APP, "No encontré sellos (cantidad en columna B y valor en columna C).")
            return
        msg = f"Encontrados {len(stock)} tipos de sello ({sum(stock.values())} sellos)"
        msg += f" y {len(tarifas)} tarifas." if tarifas else "."
        if not messagebox.askyesno(APP, msg + "\n\n¿Sustituyo tus sellos actuales por estos?"):
            return
        self.alm.stock = stock
        if tarifas:
            self.alm.tarifas = tarifas
        self.deshacer.clear()
        self.guardar()
        self.aviso("Importado correctamente", VERDE)
        self.refrescar_todo(reconstruir=True)

    # -------------------------------------------------- ACTUALIZACIONES ---
    def buscar_actualizacion(self, manual):
        repo = self.build["repo"]
        if not repo:
            if manual:
                self.lbl_up.config(text="Esta copia no viene de GitHub: descarga el .exe desde tu repositorio.")
            return
        if manual:
            self.lbl_up.config(text="Buscando…")

        def trabajo():
            req = urllib.request.Request(f"https://api.github.com/repos/{repo}/releases/latest",
                                         headers={"User-Agent": "FranqueoSupremo", "Accept": "application/vnd.github+json"})
            with urllib.request.urlopen(req, timeout=10) as r:
                d = json.loads(r.read().decode("utf-8"))
            exe = next((a for a in d.get("assets", []) if a.get("name", "").lower().endswith(".exe")), None)
            return {"version": d.get("tag_name", "").lstrip("vV"), "url": exe and exe["browser_download_url"],
                    "web": d.get("html_url", f"https://github.com/{repo}/releases")}

        def listo(r):
            if not r:
                if manual:
                    self.lbl_up.config(text="No se pudo conectar con GitHub (¿internet? ¿repositorio público?).")
                return
            if version_tupla(r["version"]) > version_tupla(self.build["version"]):
                self.nueva = r
                self.lbl_banner.config(text=f"✦  Hay una versión nueva: {r['version']}  (tienes {self.build['version']})")
                self.b_banner.habilitar(True)
                self.b_banner.config(text="Actualizar ahora")
                self.banner.pack(side="top", fill="x", before=self.paginas[self.pagina])
                if manual:
                    self.lbl_up.config(text=f"Hay versión nueva: {r['version']}")
            elif manual:
                self.lbl_up.config(text=f"Ya tienes la última versión ({self.build['version']}).")

        self._en_hilo(trabajo, listo)

    def actualizar(self):
        r = self.nueva
        if not r:
            return
        if not getattr(sys, "frozen", False) or not r["url"]:
            webbrowser.open(r["web"])
            return
        self.b_banner.habilitar(False)
        self.lbl_banner.config(text="Descargando la versión nueva… 0 %")
        exe = sys.executable
        nuevo = exe + ".nuevo"

        def progreso(p):
            self.lbl_banner.config(text=f"Descargando la versión nueva… {p} %")

        def trabajo():
            req = urllib.request.Request(r["url"], headers={"User-Agent": "FranqueoSupremo"})
            with urllib.request.urlopen(req, timeout=30) as resp, open(nuevo, "wb") as fh:
                total = int(resp.headers.get("Content-Length") or 0)
                hecho, ultimo = 0, -1
                while True:
                    trozo = resp.read(65536)
                    if not trozo:
                        break
                    fh.write(trozo)
                    hecho += len(trozo)
                    if total:
                        p = int(hecho * 100 / total)
                        if p != ultimo:
                            ultimo = p
                            self.cola.put((progreso, p))
            with open(nuevo, "rb") as fh:
                ok = fh.read(2) == b"MZ" and os.path.getsize(nuevo) > 1_000_000
            return ok

        def listo(ok):
            if not ok:
                try:
                    os.remove(nuevo)
                except OSError:
                    pass
                self.lbl_banner.config(text="No se pudo descargar la actualización. Prueba más tarde.")
                self.b_banner.habilitar(True)
                return
            self._cambiar_exe(exe, nuevo)

        self._en_hilo(trabajo, listo)

    def _cambiar_exe(self, exe, nuevo):
        if self.pendiente:
            self.guardar()
        viejo = os.path.splitext(exe)[0] + ".old.exe"
        try:
            if os.path.exists(viejo):
                try:
                    os.remove(viejo)
                except OSError:
                    viejo = os.path.splitext(exe)[0] + f".old{int(time.time())}.exe"
            os.replace(exe, viejo)      # Windows deja renombrar el .exe que está abierto
            os.replace(nuevo, exe)
        except OSError as e:
            messagebox.showerror(APP, f"No se pudo instalar la actualización:\n{e}")
            return
        env = {k: v for k, v in os.environ.items() if not k.startswith(("_PYI", "_MEI"))}
        env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
        flags = 0x00000008 | 0x00000200 if sys.platform.startswith("win") else 0
        subprocess.Popen([exe], cwd=os.path.dirname(exe), env=env, creationflags=flags, close_fds=True)
        self.destroy()

    def _borrar_exe_viejo(self):
        if not getattr(sys, "frozen", False):
            return
        carpeta = os.path.dirname(sys.executable)
        base = os.path.splitext(os.path.basename(sys.executable))[0].lower()
        for nom in os.listdir(carpeta):
            n = nom.lower()
            if n.startswith(base + ".old") and n.endswith(".exe") or n == base + ".exe.nuevo":
                try:
                    os.remove(os.path.join(carpeta, nom))
                except OSError:
                    pass

    # ---------------------------------------------------- infraestructura ---
    def guardar(self):
        try:
            self.alm.guardar()
            self.pendiente = False
        except PermissionError:
            self.pendiente = True
            self.aviso("⚠  No puedo guardar: cierra el Excel y se guardará solo", ORO, ms=5000)
        except Exception as e:
            self.pendiente = True
            self.aviso(f"⚠  Error al guardar: {e}", ROJO, ms=6000)

    def _vigilar(self):
        if self.pendiente:
            try:
                self.alm.guardar()
                self.pendiente = False
                self.aviso("Guardado en el Excel ✔", VERDE)
            except Exception:
                pass
        elif self.alm.cambiado_fuera():
            try:
                self.alm.cargar()
                self.deshacer.clear()
                self.aviso("Has cambiado el Excel: datos recargados y recalculados", SIDE)
                self.refrescar_todo(reconstruir=True)
            except Exception:
                pass  # probablemente Excel aún guardando; se reintenta
        self.after(2000, self._vigilar)

    def _en_hilo(self, trabajo, listo):
        def run():
            try:
                r = trabajo()
            except Exception as e:  # pragma: no cover
                r = None
                print("Error en cálculo:", e)
            self.cola.put((listo, r))
        threading.Thread(target=run, daemon=True).start()

    def _bombear(self):
        try:
            while True:
                fn, r = self.cola.get_nowait()
                fn(r)
        except queue.Empty:
            pass
        self.after(80, self._bombear)

    def refrescar_todo(self, reconstruir=False):
        u, v = self.alm.total_unidades(), self.alm.total_valor()
        self.lbl_tot_u.config(text=f"{u} sellos")
        self.lbl_tot_v.config(text=f"{M.euros(v)} en total")
        if reconstruir or [c["t"] for c in self.cards] != self.alm.tarifas:
            self._construir_cards()
            self._construir_ccards()
            self._rellenar_ajustes()
        self._marcar_cards()
        self._pintar_sellos()
        self._recalcular_stats()
        self._pintar_quedan()
        self._pintar_cuantos()
        if self.pagina == "historial":
            self._pintar_historial()
        if self.sel:
            self.calcular()
        else:
            self._pintar_resultado()


def main():
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
