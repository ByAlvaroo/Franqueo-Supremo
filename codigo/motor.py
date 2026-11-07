# -*- coding: utf-8 -*-
"""
Franqueo Supremo - motor de cálculo y Excel.
HECHO POR @BYALVARO__

Todo se calcula en céntimos (enteros) para que 0,96 € sea exactamente 96.
"""
import os
import random
import time
from collections import Counter
from datetime import datetime

# ----------------------------------------------------------------------------
# Datos por defecto (los de tu Excel)
# ----------------------------------------------------------------------------
SELLOS_INICIALES = [
    (0.15, 12), (0.21, 9), (0.24, 179), (0.25, 36), (0.26, 70), (0.27, 31),
    (0.28, 29), (0.29, 31), (0.30, 13), (0.31, 15), (0.32, 15), (0.34, 24),
    (0.35, 16), (0.38, 3), (0.39, 2), (0.42, 6), (0.45, 112), (0.50, 15),
    (0.55, 21), (1.00, 9), (1.25, 50), (1.80, 12), (1.85, 11), (1.90, 10),
    (1.95, 3), (2.15, 3), (2.21, 5), (2.33, 4), (2.44, 2), (2.49, 2),
    (2.60, 2), (2.70, 2),
]
# (tipo, precio, nombre)  tipo: "ord" = ordinario, "cert" = certificado
TARIFAS_INICIALES = [
    ("ord", 0.96, "Hasta 20g normalizada - Tipo A"),
    ("ord", 1.15, "20 hasta 50g - Tipo A2"),
    ("ord", 1.75, "50 hasta 100g - Tipo A3"),
    ("ord", 3.70, "100 hasta 500g"),
    ("cert", 5.74, "Hasta 20g normalizada"),
    ("cert", 5.93, "20 hasta 50g"),
    ("cert", 6.53, "50 hasta 100g"),
    ("cert", 8.48, "100 hasta 500g"),
]
# Tarifas de la primera versión: si el Excel aún las tiene, se cambian por las nuevas
TARIFAS_VIEJAS = [96, 115, 370, 593, 653, 848]


def tarifa(tipo, cent, nombre):
    return {"tipo": tipo, "cent": int(cent), "nombre": (nombre or "").strip() or euros(cent)}


def tarifas_iniciales():
    return [tarifa(t, a_cent(p), n) for t, p, n in TARIFAS_INICIALES]


def etiqueta(t):
    """Nombre completo para historial y títulos."""
    return ("Certificado - " + t["nombre"]) if t["tipo"] == "cert" else t["nombre"]


def ordenar_tarifas(lista):
    """Ordinarias primero y después certificadas (respetando el orden de cada tabla)."""
    return [t for t in lista if t["tipo"] == "ord"] + [t for t in lista if t["tipo"] == "cert"]


def a_cent(euros):
    return int(round(float(euros) * 100))


def euros(cent):
    """96 -> '0,96 €'"""
    s = f"{cent / 100:.2f}".replace(".", ",")
    return s + " €"


def euros_corto(cent):
    return f"{cent / 100:.2f}".replace(".", ",")


# ----------------------------------------------------------------------------
# Cálculo
# ----------------------------------------------------------------------------
def _alcanzables(tipos, stock, objetivo):
    """
    suf[i] = bitset de importes alcanzables usando los tipos tipos[i:].
    Bit k encendido = se puede pagar exactamente k céntimos.
    Usa partición binaria de las cantidades (rápido aunque tengas 179 sellos).
    """
    mask = (1 << (objetivo + 1)) - 1
    n = len(tipos)
    suf = [0] * (n + 1)
    suf[n] = 1
    for i in range(n - 1, -1, -1):
        v = tipos[i]
        c = min(stock[v], objetivo // v)
        b = suf[i + 1]
        k = 1
        while c > 0:
            t = min(k, c)
            b = (b | (b << (v * t))) & mask
            c -= t
            k <<= 1
        suf[i] = b
    return suf


def reparto(objetivo, stock, orden=None):
    """
    Devuelve {valor_cent: cantidad} que suma EXACTAMENTE `objetivo`,
    o None si no hay forma.

    Recorre los tipos en `orden` (por defecto: de mayor a menor valor) y de
    cada uno coge TODOS los que pueda siempre que el resto siga siendo
    pagable exacto con los demás. Así se gastan primero los más altos.
    """
    if objetivo <= 0:
        return None
    tipos = [v for v, c in stock.items() if c > 0 and v <= objetivo]
    if orden is None:
        tipos.sort(reverse=True)
    else:
        pos = {v: i for i, v in enumerate(orden)}
        tipos.sort(key=lambda v: pos.get(v, 10 ** 9))
    if not tipos:
        return None
    suf = _alcanzables(tipos, stock, objetivo)
    if not (suf[0] >> objetivo) & 1:
        return None
    resto = objetivo
    usado = {}
    for i, v in enumerate(tipos):
        if resto == 0:
            break
        maxk = min(stock[v], resto // v)
        for k in range(maxk, -1, -1):
            r = resto - k * v
            if (suf[i + 1] >> r) & 1:
                if k:
                    usado[v] = k
                resto = r
                break
    if resto != 0:  # no debería pasar
        return None
    return usado


def _restar(stock, usado):
    for v, k in usado.items():
        stock[v] -= k


def _simular(objetivo, stock, orden_fn, limite):
    s = dict(stock)
    sobres = []
    while len(sobres) < limite:
        r = reparto(objetivo, s, orden_fn(s))
        if not r:
            break
        _restar(s, r)
        sobres.append(r)
    return sobres


def maximo_envios(objetivo, stock, segundos=1.0, semilla=None):
    """
    Busca el mayor número de sobres de `objetivo` céntimos que se pueden
    franquear con el stock (cada sobre con importe exacto).

    Es un problema combinatorio difícil, así que prueba varias estrategias
    y muchas mezclas aleatorias dentro del tiempo dado, y se queda con la
    que más sobres consigue. Devuelve dict con:
        sobres  -> lista de repartos (uno por sobre)
        tope    -> máximo teórico (valor útil / importe)
        sobran  -> stock que queda tras el plan
    """
    if objetivo <= 0:
        return {"sobres": [], "tope": 0, "sobran": dict(stock)}
    util = sum(v * c for v, c in stock.items() if v <= objetivo and c > 0)
    tope = util // objetivo
    if tope == 0:
        return {"sobres": [], "tope": 0, "sobran": dict(stock)}

    rnd = random.Random(semilla)
    estrategias = [
        lambda s: None,                                                 # más altos primero
        lambda s: sorted(s, key=lambda v: (-s[v] * v, -v)),            # más valor acumulado primero
        lambda s: sorted(s, key=lambda v: (-s[v], -v)),                # los más abundantes primero
        lambda s: sorted(s),                                            # más bajos primero
    ]
    mejor = None
    fin = time.time() + segundos
    i = 0
    while True:
        if i < len(estrategias):
            fn = estrategias[i]
        else:
            # orden aleatorio, sesgado hacia los altos y abundantes
            pesos = {v: rnd.random() * (1 + (v / objetivo)) for v in stock}
            fn = (lambda p: (lambda s: sorted(s, key=lambda v: -p[v])))(pesos)
        sobres = _simular(objetivo, stock, fn, tope)
        if mejor is None or len(sobres) > len(mejor):
            mejor = sobres
        i += 1
        if len(mejor) >= tope:
            break
        if i >= len(estrategias) + 8 or (i >= len(estrategias) and time.time() > fin):
            break

    # Mejora: deshacer unos cuantos sobres al azar y volver a montarlos de otra forma
    while len(mejor) < tope and time.time() < fin:
        quitar = rnd.randint(2, max(2, min(12, len(mejor))))
        idx = set(rnd.sample(range(len(mejor)), min(quitar, len(mejor))))
        base = [r for j, r in enumerate(mejor) if j not in idx]
        s = dict(stock)
        for r in base:
            _restar(s, r)
        pesos = {v: rnd.random() for v in stock}
        nuevos = _simular(objetivo, s, lambda st, p=pesos: sorted(st, key=lambda v: -p[v]), tope - len(base))
        if len(base) + len(nuevos) >= len(mejor):
            mejor = base + nuevos
    sobran = dict(stock)
    for r in mejor:
        _restar(sobran, r)
    return {"sobres": mejor, "tope": tope, "sobran": sobran}


def agrupar(sobres):
    """Agrupa sobres iguales: [(reparto_dict, veces), ...] en orden de aparición."""
    claves = []
    cuenta = Counter()
    ejemplo = {}
    for r in sobres:
        k = tuple(sorted(r.items(), reverse=True))
        if k not in cuenta:
            claves.append(k)
            ejemplo[k] = r
        cuenta[k] += 1
    return [(ejemplo[k], cuenta[k]) for k in claves]


def alternativas(objetivo, stock, modo="altos", n=6, segundos=0.6, plan=None):
    """
    Lista de combinaciones distintas para UN sobre, la mejor primero según el modo.
      modo 'altos'  -> gastar primero los sellos de más valor
      modo 'envios' -> el primer sobre del plan que más envíos consigue
    """
    vistos = []

    def add(r):
        if r and r not in vistos:
            vistos.append(r)

    if modo == "envios":
        if plan is None:
            plan = maximo_envios(objetivo, stock, segundos)
        # primero las combinaciones que más se repiten en el mejor plan
        for r, _ in sorted(agrupar(plan["sobres"]), key=lambda g: -g[1]):
            add(r)
    add(reparto(objetivo, stock))
    # variaciones: saltarse el sello más alto usado, luego el siguiente...
    base = reparto(objetivo, stock)
    if base:
        for v in sorted(base, reverse=True):
            s2 = dict(stock)
            s2[v] = 0
            add(reparto(objetivo, s2))
            if len(vistos) >= n:
                break
    rnd = random.Random(objetivo)
    intentos = 0
    while len(vistos) < n and intentos < 40:
        orden = sorted(stock, key=lambda v: -rnd.random() * (1 + v))
        add(reparto(objetivo, stock, orden))
        intentos += 1
    return vistos[:n]


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
class Almacen:
    """Guarda sellos, tarifas e historial en un .xlsx (fila 5 cabecera, como el tuyo)."""

    FILA_CAB = 5
    FILA_INI = 7

    def __init__(self, ruta):
        self.ruta = ruta
        self.stock = {}          # cent -> unidades
        self.tarifas = []        # [{tipo, cent, nombre}]
        self.historial = []      # [(fecha, tarifa, importe_cent, texto)]
        self.mtime = 0
        if os.path.exists(ruta):
            self.cargar()
        else:
            self.stock = {a_cent(v): c for v, c in SELLOS_INICIALES}
            self.tarifas = tarifas_iniciales()
            self.guardar()

    # --- lectura ------------------------------------------------------------
    @staticmethod
    def leer_libro(ruta):
        """Lee un Excel con tu formato: devuelve (stock, tarifas, historial)."""
        import openpyxl
        wb = openpyxl.load_workbook(ruta, data_only=True)
        ws = wb["Sellos"] if "Sellos" in wb.sheetnames else wb.worksheets[0]
        stock = {}
        tarifas = {}

        def num(x):
            if isinstance(x, bool):
                return None
            if isinstance(x, (int, float)):
                return float(x)
            if isinstance(x, str):
                try:
                    return float(x.replace("€", "").replace(",", ".").strip())
                except ValueError:
                    return None
            return None

        nuevas = []      # formato actual: tablas ORDINARIO (F:G) y CERTIFICADO (I:J)
        viejas = {}      # formato antiguo: nº en columna A, precio en B, nombre en C
        hay_tablas = False
        for fila in ws.iter_rows(min_row=1, max_col=10, values_only=True):
            fila = (list(fila) + [None] * 10)[:10]
            a, b, c = fila[0], fila[1], fila[2]
            na, nb, nc = num(a), num(b), num(c)
            if a is None and nb is not None and nc is not None and nc > 0:
                v = a_cent(nc)
                stock[v] = stock.get(v, 0) + max(0, int(round(nb)))
            elif na is not None and float(na).is_integer() and 1 <= na <= 20 and nb is not None:
                n = int(na)
                nombre = c if isinstance(c, str) and c.strip() else f"Tarifa {n}"
                viejas[n] = (a_cent(nb), nombre.strip())
            for col_n, tipo in ((5, "ord"), (8, "cert")):
                nom, pre = fila[col_n], fila[col_n + 1]
                if isinstance(nom, str) and nom.strip().upper() in ("ORDINARIO", "CERTIFICADO", "NOMBRE"):
                    hay_tablas = True
                    continue
                np_ = num(pre)
                if np_ is not None and np_ > 0:
                    nuevas.append(tarifa(tipo, a_cent(np_), nom if isinstance(nom, str) else ""))
        if hay_tablas:
            tarifas = ordenar_tarifas(nuevas)
        else:
            tarifas = []
            for n in sorted(viejas):
                cent, nombre = viejas[n]
                if nombre.lower().startswith("certificado"):
                    tarifas.append(tarifa("cert", cent, nombre.split("-", 1)[-1].split("·", 1)[-1]))
                else:
                    tarifas.append(tarifa("ord", cent, nombre))
            tarifas = ordenar_tarifas(tarifas)
        hist = []
        if "Historial" in wb.sheetnames:
            for fila in wb["Historial"].iter_rows(min_row=2, max_col=4, values_only=True):
                if fila and fila[0]:
                    f, t, imp, txt = (list(fila) + [None] * 4)[:4]
                    fs = f.strftime("%d/%m/%Y %H:%M") if hasattr(f, "strftime") else str(f)
                    hist.append((fs, str(t or ""), a_cent(num(imp) or 0), str(txt or "")))
        return stock, tarifas, hist

    def cargar(self):
        stock, tarifas, hist = self.leer_libro(self.ruta)
        self.stock = stock
        if tarifas:
            self.tarifas = tarifas
        elif not self.tarifas:
            self.tarifas = tarifas_iniciales()
        self.historial = hist
        self.mtime = os.path.getmtime(self.ruta)
        if [t["cent"] for t in self.tarifas] == TARIFAS_VIEJAS:
            self.tarifas = tarifas_iniciales()
            try:
                self.guardar()
            except OSError:
                pass

    def cambiado_fuera(self):
        try:
            return os.path.getmtime(self.ruta) > self.mtime + 0.5
        except OSError:
            return False

    # --- escritura ----------------------------------------------------------
    def guardar(self):
        """Escribe el Excel completo. Lanza PermissionError si está abierto en Excel."""
        import openpyxl
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sellos"
        rojo = PatternFill("solid", fgColor="C8322F")
        azul = PatternFill("solid", fgColor="1F2A44")
        crema = PatternFill("solid", fgColor="F6F1E7")
        blanco = Font(color="FFFFFF", bold=True)
        fino = Side(style="thin", color="D9CFBC")
        borde = Border(left=fino, right=fino, top=fino, bottom=fino)
        centro = Alignment(horizontal="center", vertical="center")

        ws["A1"] = "FRANQUEO SUPREMO · Mis sellos"
        ws["A1"].font = Font(size=16, bold=True, color="1F2A44")
        ws["A2"] = "Columna B = cantidad · Columna C = valor. Puedes editarlo aquí o en la app."
        ws["A2"].font = Font(italic=True, color="7A7060")
        ws["A3"] = "HECHO POR @BYALVARO__"
        ws["A3"].font = Font(size=9, bold=True, color="C8322F")

        f = self.FILA_CAB
        for col, txt in (("B", "Unidades"), ("C", "Valor euros unidad"), ("D", "Subtotal €")):
            ws[f"{col}{f}"] = txt
            ws[f"{col}{f}"].fill = azul
            ws[f"{col}{f}"].font = blanco
            ws[f"{col}{f}"].alignment = centro

        fila = self.FILA_INI
        for v in sorted(self.stock):
            ws[f"B{fila}"] = int(self.stock[v])
            ws[f"C{fila}"] = v / 100
            ws[f"D{fila}"] = f"=B{fila}*C{fila}"
            ws[f"C{fila}"].number_format = '0.00 "€"'
            ws[f"D{fila}"].number_format = '0.00 "€"'
            for col in "BCD":
                ws[f"{col}{fila}"].border = borde
                ws[f"{col}{fila}"].alignment = centro
                if fila % 2:
                    ws[f"{col}{fila}"].fill = crema
            fila += 1
        ultima = fila - 1
        ws[f"A{fila}"] = "TOTAL"
        ws[f"A{fila}"].font = Font(bold=True)
        ws[f"B{fila}"] = f"=SUM(B{self.FILA_INI}:B{ultima})"
        ws[f"D{fila}"] = f"=SUM(D{self.FILA_INI}:D{ultima})"
        ws[f"D{fila}"].number_format = '0.00 "€"'
        for col in "BD":
            ws[f"{col}{fila}"].font = Font(bold=True)
            ws[f"{col}{fila}"].alignment = centro

        # Tablas de tarifas a la derecha: cada una crece hacia abajo por separado
        ws["F2"] = "Tarifas: añade filas debajo de cada tabla (nombre + precio) para crear nuevas."
        ws["F2"].font = Font(italic=True, color="7A7060")
        for col_n, col_p, titulo, relleno, tipo in (("F", "G", "ORDINARIO", rojo, "ord"),
                                                    ("I", "J", "CERTIFICADO", azul, "cert")):
            ws[f"{col_n}4"] = titulo
            ws[f"{col_n}4"].font = Font(size=12, bold=True, color="C8322F" if tipo == "ord" else "1F2A44")
            for col, txt in ((col_n, "Nombre"), (col_p, "Precio €")):
                ws[f"{col}{f}"] = txt
                ws[f"{col}{f}"].fill = relleno
                ws[f"{col}{f}"].font = blanco
                ws[f"{col}{f}"].alignment = centro
            fila_t = f + 1
            for t in self.tarifas:
                if t["tipo"] != tipo:
                    continue
                ws[f"{col_n}{fila_t}"] = t["nombre"]
                ws[f"{col_p}{fila_t}"] = t["cent"] / 100
                ws[f"{col_p}{fila_t}"].number_format = '0.00 "€"'
                ws[f"{col_p}{fila_t}"].alignment = centro
                for col in (col_n, col_p):
                    ws[f"{col}{fila_t}"].border = borde
                fila_t += 1
        ws.column_dimensions["E"].width = 4
        ws.column_dimensions["F"].width = 32
        ws.column_dimensions["G"].width = 11
        ws.column_dimensions["H"].width = 4
        ws.column_dimensions["I"].width = 28
        ws.column_dimensions["J"].width = 11

        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 26
        ws.column_dimensions["D"].width = 14

        h = wb.create_sheet("Historial")
        for i, txt in enumerate(("Fecha", "Tarifa", "Importe", "Sellos usados"), 1):
            c = h.cell(row=1, column=i, value=txt)
            c.fill = azul
            c.font = blanco
        for i, (fe, t, imp, txt) in enumerate(self.historial, 2):
            h.cell(row=i, column=1, value=fe)
            h.cell(row=i, column=2, value=t)
            c = h.cell(row=i, column=3, value=imp / 100)
            c.number_format = '0.00 "€"'
            h.cell(row=i, column=4, value=txt)
        h.column_dimensions["A"].width = 18
        h.column_dimensions["B"].width = 26
        h.column_dimensions["C"].width = 10
        h.column_dimensions["D"].width = 60

        tmp = self.ruta + ".tmp.xlsx"
        wb.save(tmp)
        try:
            os.replace(tmp, self.ruta)
        except PermissionError:
            try:
                os.remove(tmp)
            except OSError:
                pass
            raise
        self.mtime = os.path.getmtime(self.ruta)

    # --- operaciones ---------------------------------------------------------
    def total_unidades(self):
        return sum(self.stock.values())

    def total_valor(self):
        return sum(v * c for v, c in self.stock.items())

    def descontar(self, usado, etiqueta, importe):
        for v, k in usado.items():
            if self.stock.get(v, 0) < k:
                raise ValueError("No tienes suficientes sellos de " + euros(v))
        for v, k in usado.items():
            self.stock[v] -= k
        texto = " + ".join(f"{k}×{euros_corto(v)}" for v, k in sorted(usado.items(), reverse=True))
        self.historial.insert(0, (datetime.now().strftime("%d/%m/%Y %H:%M"), etiqueta, importe, texto))

    def devolver(self, usado):
        for v, k in usado.items():
            self.stock[v] = self.stock.get(v, 0) + k
        if self.historial:
            self.historial.pop(0)


def texto_reparto(r):
    return " + ".join(f"{k} × {euros_corto(v)}" for v, k in sorted(r.items(), reverse=True))


def num_sellos(r):
    return sum(r.values())
