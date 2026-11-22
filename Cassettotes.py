# -*- coding: utf-8 -*-
"""
Cassettotes - Lista de canciones por duración máxima (27, 30, 45, 60 o 90 minutos).

- Arrastra archivos de audio a la lista: el programa lee su duración (metadatos)
  y va sumando. Las canciones que superan el máximo elegido se marcan en rojo.
- Borrar Cara vacía la cara seleccionada. Borrar Cinta vacía las dos caras.
  Borrar sobrante quita las canciones que no caben en el máximo. Con la tecla
  Supr se quitan las canciones seleccionadas.
- Copiar a carpeta: crea "Cara A" y "Cara B" dentro de la carpeta que elijas y
  copia en cada una las canciones de su cara.
- Una cinta tiene dos caras (A y B). El selector Cara A / Cara B decide qué
  lista se está editando; cada cara se mide contra el máximo por separado.
- Una sesión puede tener varias cintas. El contador "Cinta: n de m" de abajo a
  la derecha sirve para cambiar de cinta; "Nueva cinta" crea otra.
- Archivo > Guardar cinta: guarda todos los datos de la cinta actual (las dos
  caras, el máximo y las rutas) en un archivo .cinta que se puede cargar después.
- Archivo > Cargar cinta: recupera una cinta guardada.
- Archivo > Exportar cinta a TXT: guarda una copia legible de las dos caras.
- Archivo > Guardar sesión: guarda todas las cintas en una hoja de cálculo (.xlsx).
- Archivo > Cargar sesión: recupera todas las cintas de la sesión y se pueden
  recorrer con el contador de cintas.
- Archivo > Añadir archivos: alternativa a arrastrar y soltar.

Requisitos para ejecutar desde Python:  pip install mutagen openpyxl tkinterdnd2
"""

import json
import math
import os
import queue
import re
import shutil
import sys
import threading
import time
import tkinter as tk
import unicodedata
from collections import namedtuple
from tkinter import filedialog, messagebox, ttk

try:
    from mutagen import File as MutagenFile
except ImportError:  # se avisa al arrancar
    MutagenFile = None

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
except ImportError:  # se avisa al arrancar
    Workbook = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:  # sin arrastrar y soltar: queda Archivo > Añadir archivos
    DND_FILES = None
    TkinterDnD = None

try:
    import pygame
    pygame.mixer.init()
    PYGAME_OK = True
except Exception:  # sin pygame: el reproductor queda desactivado
    pygame = None
    PYGAME_OK = False

EXTENSIONES = {
    ".mp3", ".flac", ".wav", ".ogg", ".oga", ".opus", ".m4a", ".mp4",
    ".aac", ".wma", ".aiff", ".aif", ".ape", ".wv", ".mpc",
}

ENCABEZADOS = ("Cinta", "Cara", "Canción", "Duración", "Ruta")
ENCABEZADOS_CINTAS = ("Cinta", "Máximo", "Canciones", "Duración total")
COLOR_NORMAL = "#000000"
COLOR_ROJO = "#C00000"

Cancion = namedtuple("Cancion", "ruta nombre segundos")


def cinta_vacia(maximo=30):
    """Una cinta con las dos caras vacías."""
    return {"maximo": int(maximo),
            "caras": {"A": [], "B": []},
            "copiada": {"A": False, "B": False},
            "nombre": ""}


def cinta_tiene_canciones(cinta):
    return bool(cinta["caras"]["A"] or cinta["caras"]["B"])


def _clase_base():
    """TkinterDnD.Tk si arrastrar y soltar funciona en este equipo; si no, tk.Tk normal."""
    if TkinterDnD is not None:
        try:
            prueba = TkinterDnD.Tk()
            prueba.withdraw()
            prueba.destroy()
            return TkinterDnD.Tk
        except Exception:
            pass
    return tk.Tk


_BaseTk = _clase_base()
DND_OK = _BaseTk is not tk.Tk


# ----------------------------------------------------------------------------
# Lógica (sin interfaz)
# ----------------------------------------------------------------------------
def formato_tiempo(segundos):
    segundos = int(segundos)
    h, resto = divmod(segundos, 3600)
    m, s = divmod(resto, 60)
    if h:
        return f"{h}:{m:02d}:{s:02d}"
    return f"{m:02d}:{s:02d}"


def leer_texto_duracion(texto):
    """'mm:ss' o 'h:mm:ss' -> segundos. None si no se entiende."""
    if not isinstance(texto, str):
        return None
    partes = [p.strip() for p in texto.strip().split(":")]
    if len(partes) not in (2, 3) or not all(p.isdigit() for p in partes):
        return None
    nums = [int(p) for p in partes]
    if len(nums) == 2:
        return nums[0] * 60 + nums[1]
    return nums[0] * 3600 + nums[1] * 60 + nums[2]


def clave_ruta(ruta):
    return os.path.normcase(os.path.abspath(ruta))


def ruta_recurso(nombre):
    """Ruta de un recurso (icono) tanto en el .py como dentro del .exe."""
    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, nombre)


def nombre_archivo(ruta):
    """Nombre del archivo sin carpeta, valga la ruta de Windows o de otro sistema."""
    nombre = re.split(r"[\\/]", str(ruta))[-1]
    return unicodedata.normalize("NFC", nombre)


def clave_cancion(ruta, segundos):
    """
    Identifica una canción por nombre de archivo + duración. Así sigue
    reconociéndose aunque la carpeta de música haya cambiado de sitio o de unidad.
    """
    return (nombre_archivo(ruta).lower(), segundos)


def esta_excluida(cancion, claves):
    n = nombre_archivo(cancion.ruta).lower()
    # (n, None): fila de sesión sin duración legible; se excluye por nombre
    return (n, cancion.segundos) in claves or (n, None) in claves


def buscar_audios(carpeta, con_subcarpetas):
    rutas = []
    if con_subcarpetas:
        for raiz, _dirs, archivos in os.walk(carpeta):
            for a in archivos:
                if os.path.splitext(a)[1].lower() in EXTENSIONES:
                    rutas.append(os.path.join(raiz, a))
    else:
        for a in os.listdir(carpeta):
            ruta = os.path.join(carpeta, a)
            if os.path.isfile(ruta) and os.path.splitext(a)[1].lower() in EXTENSIONES:
                rutas.append(ruta)
    return sorted(rutas, key=lambda r: r.lower())


def expandir_rutas(rutas):
    """Archivos de audio de lo arrastrado (las carpetas se recorren). -> (audios, ignorados)"""
    audios, ignorados = [], 0
    for r in rutas:
        if os.path.isdir(r):
            audios.extend(buscar_audios(r, True))
        elif os.path.isfile(r) and os.path.splitext(r)[1].lower() in EXTENSIONES:
            audios.append(r)
        else:
            ignorados += 1
    return audios, ignorados


def leer_duracion(ruta):
    """Duración en segundos (redondeada hacia arriba) o None si no se puede leer."""
    try:
        audio = MutagenFile(ruta)
        if audio is None or audio.info is None:
            return None
        largo = float(audio.info.length)
        if largo <= 0:
            return None
        return int(math.ceil(largo))
    except Exception:
        return None


def leer_canciones(rutas, cache):
    """Lee la duración de cada archivo. -> (canciones, n_ilegibles)"""
    canciones, ilegibles = [], 0
    for ruta in rutas:
        try:
            fecha = os.path.getmtime(ruta)
        except OSError:
            ilegibles += 1
            continue
        guardado = cache.get(ruta)
        if guardado and guardado[0] == fecha:
            seg = guardado[1]
        else:
            seg = leer_duracion(ruta)
            cache[ruta] = (fecha, seg)
        if seg is None:
            ilegibles += 1
        else:
            nombre = os.path.splitext(os.path.basename(ruta))[0]
            canciones.append(Cancion(ruta, nombre, seg))
    return canciones, ilegibles


# -- copia -------------------------------------------------------------------
def nombre_destino_unico(carpeta, nombre):
    base, ext = os.path.splitext(nombre)
    ruta = os.path.join(carpeta, nombre)
    n = 2
    while os.path.exists(ruta):
        ruta = os.path.join(carpeta, f"{base} ({n}){ext}")
        n += 1
    return ruta


def copiar_canciones(canciones, destino):
    """Copia los archivos a 'destino'. Devuelve (copiadas, fallos[(nombre, error)])."""
    copiadas, fallos = [], []
    for c in canciones:
        final = nombre_destino_unico(destino, os.path.basename(c.ruta))
        try:
            shutil.copy2(c.ruta, final)
            copiadas.append(c)
        except (OSError, shutil.Error) as e:
            fallos.append((c.nombre, str(e)))
            try:  # no dejar un archivo a medias
                if os.path.exists(final):
                    os.remove(final)
            except OSError:
                pass
    return copiadas, fallos


def carpeta_de_cara(destino, clave):
    return os.path.join(destino, f"Cara {clave}")


def nombre_carpeta_cinta(indice, nombre_personalizado):
    """Devuelve el nombre de la carpeta para la cinta.
    Si hay un nombre personalizado lo usa, sino usa 'Cinta X'."""
    if nombre_personalizada := nombre_personalizado.strip():
        return nombre_personalizada
    return f"Cinta {indice + 1}"


def copiar_cinta(caras, destino, indice, nombre_cinta):
    """Copia cada cara en <destino>\\Cinta X\\Cara A y <destino>\\Cinta X\\Cara B.
    
    Si el nombre de la cinta está vacío, usa 'Cinta X' donde X es el índice + 1.

    Devuelve {clave: (copiadas, fallos[(nombre, error)])}.
    """
    nombre_carpeta = nombre_carpeta_cinta(indice, nombre_cinta)
    carpeta_cinta = os.path.join(destino, nombre_carpeta)
    
    resultado = {}
    for clave in ("A", "B"):
        carpeta = os.path.join(carpeta_cinta, f"Cara {clave}")
        canciones = caras.get(clave) or []
        try:
            os.makedirs(carpeta, exist_ok=True)
        except OSError as e:
            resultado[clave] = ([], [(f"Cara {clave}", str(e))])
            continue
        resultado[clave] = copiar_canciones(canciones, carpeta) if canciones else ([], [])
    return resultado


# -- cinta en TXT -------------------------------------------------------------
def guardar_cinta_txt(ruta, caras, maximo_min, nombre_cinta=""):
    """Escribe las dos caras de la cinta en un TXT."""
    if nombre_cinta:
        lineas = [f"{nombre_cinta} (máximo {maximo_min} minutos por cara)", ""]
    else:
        lineas = [f"Cinta (máximo {maximo_min} minutos por cara)", ""]
    for etiqueta, clave in (("CARA A", "A"), ("CARA B", "B")):
        canciones = caras.get(clave, [])
        total = sum(c.segundos for c in canciones)
        lineas.append(f"=== {etiqueta} ({formato_tiempo(total)}) ===")
        if not canciones:
            lineas.append("(vacía)")
        for i, c in enumerate(canciones, 1):
            lineas.append(f"{i:02d}. {c.nombre}  -  {formato_tiempo(c.segundos)}")
        lineas.append("")
    with open(ruta, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write("\n".join(lineas) + "\n")


# -- cinta completa (.cinta, JSON) ---------------------------------------------
def guardar_cinta_json(ruta, cinta):
    """Guarda todos los datos de una cinta para poder cargarla después."""
    datos = {
        "programa": "Cassettotes",
        "version": 1,
        "maximo": int(cinta["maximo"]),
        "nombre": cinta.get("nombre", ""),
        "caras": {
            clave: [
                {"ruta": c.ruta, "nombre": c.nombre, "segundos": int(c.segundos)}
                for c in cinta["caras"][clave]
            ]
            for clave in ("A", "B")
        },
    }
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)


def _cancion_desde_json(item):
    if isinstance(item, str):  # por si el archivo es de una versión anterior
        item = {"ruta": item}
    if not isinstance(item, dict):
        return None
    ruta = str(item.get("ruta") or "").strip()
    if not ruta:
        return None
    segundos = item.get("segundos")
    try:
        segundos = max(0, int(segundos))
    except (TypeError, ValueError):
        segundos = leer_texto_duracion(item.get("duracion")) or 0
    nombre = str(item.get("nombre") or "").strip()
    if not nombre:
        nombre = os.path.splitext(os.path.basename(ruta))[0]
    return Cancion(ruta, nombre, segundos)


def cargar_cinta_json(ruta):
    """Lee una cinta guardada. Lanza ValueError si el archivo no es válido."""
    with open(ruta, encoding="utf-8") as f:
        datos = json.load(f)
    if not isinstance(datos, dict):
        raise ValueError("El archivo no contiene una cinta válida.")
    try:
        maximo = int(datos.get("maximo", 30))
    except (TypeError, ValueError):
        maximo = 30
    cinta = cinta_vacia(maximo if maximo in (27, 30, 45, 60, 90) else 30)
    # Cargar el nombre de la cinta si existe
    if nombre := datos.get("nombre"):
        cinta["nombre"] = str(nombre)
    caras = datos.get("caras") or {}
    for clave in ("A", "B"):
        for item in caras.get(clave) or []:
            cancion = _cancion_desde_json(item)
            if cancion is not None:
                cinta["caras"][clave].append(cancion)
    if not cinta_tiene_canciones(cinta):
        raise ValueError("La cinta guardada no tiene canciones.")
    return cinta


# -- sesión (hoja de cálculo, por cintas) --------------------------------------
def _estilo_cabecera(ws):
    for celda in ws[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F3A5F")
        celda.alignment = Alignment(horizontal="center", vertical="center")


def guardar_sesion_xlsx(ruta_xlsx, cintas):
    """Guarda todas las cintas: una hoja con las canciones y otra con el resumen."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sesión"
    ws.append(list(ENCABEZADOS))
    _estilo_cabecera(ws)

    banda = PatternFill("solid", fgColor="EAF0F7")
    for numero, cinta in enumerate(cintas, 1):
        for clave in ("A", "B"):
            for c in cinta["caras"][clave]:
                ws.append([numero, f"Cara {clave}", c.nombre,
                           formato_tiempo(c.segundos), c.ruta])
                fila = ws[ws.max_row]
                for i in (0, 1, 3):
                    fila[i].alignment = Alignment(horizontal="center")
                if numero % 2 == 0:
                    for celda in fila:
                        celda.fill = banda

    for col, ancho in zip("ABCDE", (8, 9, 50, 11, 80)):
        ws.column_dimensions[col].width = ancho
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    resumen = wb.create_sheet("Cintas")
    resumen.append(list(ENCABEZADOS_CINTAS) + ["Nombre"])
    _estilo_cabecera(resumen)
    for numero, cinta in enumerate(cintas, 1):
        canciones = cinta["caras"]["A"] + cinta["caras"]["B"]
        total = sum(c.segundos for c in canciones)
        nombre = cinta.get("nombre", "")
        resumen.append([numero, cinta["maximo"], len(canciones), formato_tiempo(total), nombre])
        fila = resumen[resumen.max_row]
        for celda in fila:
            celda.alignment = Alignment(horizontal="center")
    for col, ancho in zip("ABCDE", (8, 10, 12, 16, 30)):
        resumen.column_dimensions[col].width = ancho
    resumen.freeze_panes = "A2"

    wb.save(ruta_xlsx)


def _normalizar(texto):
    t = unicodedata.normalize("NFD", str(texto or ""))
    return t.encode("ascii", "ignore").decode().strip().lower()


def _valor_columna(fila, indice, nombre):
    i = indice.get(nombre)
    if i is None or i >= len(fila):
        return None
    return fila[i]


def _numero_positivo(valor):
    try:
        numero = int(float(valor))
    except (TypeError, ValueError):
        return None
    return numero if numero > 0 else None


def _cara_de(valor):
    cara = _normalizar(valor).upper()
    if cara.startswith("CARA "):
        cara = cara[5:].strip()
    return cara if cara in ("A", "B") else "A"


def _maximo_de(valor):
    try:
        maximo = int(float(valor))
    except (TypeError, ValueError):
        return None
    return maximo if maximo in (27, 30, 45, 60, 90) else None


def _cancion_de_fila(fila, indice):
    ruta = _valor_columna(fila, indice, "ruta")
    if ruta is None or str(ruta).strip() == "":
        return None
    ruta = str(ruta).strip()
    nombre = str(_valor_columna(fila, indice, "cancion") or "").strip()
    if not nombre:
        nombre = os.path.splitext(os.path.basename(ruta))[0]
    segundos = leer_texto_duracion(_valor_columna(fila, indice, "duracion"))
    return Cancion(ruta, nombre, segundos or 0)


def cargar_sesion_xlsx(ruta_xlsx):
    """Devuelve (cintas, omitidas). Lanza ValueError si no es una sesión válida."""
    wb = load_workbook(ruta_xlsx, read_only=True, data_only=True)
    try:
        hojas = {_normalizar(n): wb[n] for n in wb.sheetnames}
        ws = hojas.get("sesion") or wb.worksheets[0]
        filas_hoja = ws.iter_rows(values_only=True)
        cabecera = next(filas_hoja, None)
        if cabecera is None:
            raise ValueError("La hoja de cálculo está vacía.")
        indice = {_normalizar(v): i for i, v in enumerate(cabecera) if v is not None}
        if any(n not in indice for n in ("cancion", "duracion", "ruta")):
            raise ValueError(
                "Este archivo no parece una sesión de este programa.\n"
                "Debe tener las columnas: Cinta, Cara, Canción, Duración y Ruta.")

        formato_nuevo = "cinta" in indice and "cara" in indice
        formato_antiguo = "lista" in indice
        if not formato_nuevo and not formato_antiguo:
            raise ValueError(
                "Este archivo no parece una sesión de este programa.\n"
                "Debe tener las columnas: Cinta, Cara, Canción, Duración y Ruta.")

        omitidas = 0
        if formato_nuevo:
            cintas = {}
            for fila in filas_hoja:
                if all(v is None or str(v).strip() == "" for v in fila):
                    continue
                numero = _numero_positivo(_valor_columna(fila, indice, "cinta"))
                cancion = _cancion_de_fila(fila, indice)
                if numero is None or cancion is None:
                    omitidas += 1
                    continue
                cinta = cintas.setdefault(numero, cinta_vacia())
                cinta["caras"][_cara_de(_valor_columna(fila, indice, "cara"))].append(cancion)

            if "cintas" in hojas:  # conserva los máximos y las cintas vacías
                filas_resumen = hojas["cintas"].iter_rows(values_only=True)
                cab_resumen = next(filas_resumen, None) or ()
                idx_resumen = {_normalizar(v): i for i, v in enumerate(cab_resumen)
                               if v is not None}
                for fila in filas_resumen:
                    numero = _numero_positivo(_valor_columna(fila, idx_resumen, "cinta"))
                    if numero is None:
                        continue
                    cinta = cintas.setdefault(numero, cinta_vacia())
                    maximo = _maximo_de(_valor_columna(fila, idx_resumen, "maximo"))
                    if maximo is not None:
                        cinta["maximo"] = maximo
                    # Cargar el nombre de la cinta si existe
                    if nombre := _valor_columna(fila, idx_resumen, "nombre"):
                        cinta["nombre"] = str(nombre)
            ordenadas = [cintas[n] for n in sorted(cintas)]
        else:
            # Sesiones antiguas: cada "Lista" era una cara copiada. Se emparejan
            # de dos en dos (Cara A y Cara B de la misma cinta).
            listas = {}
            for fila in filas_hoja:
                if all(v is None or str(v).strip() == "" for v in fila):
                    continue
                numero = _numero_positivo(_valor_columna(fila, indice, "lista"))
                cancion = _cancion_de_fila(fila, indice)
                if numero is None or cancion is None:
                    omitidas += 1
                    continue
                listas.setdefault(numero, []).append(cancion)
            numeros = sorted(listas)
            ordenadas = []
            for i in range(0, len(numeros), 2):
                cinta = cinta_vacia()
                cinta["caras"]["A"] = listas[numeros[i]]
                if i + 1 < len(numeros):
                    cinta["caras"]["B"] = listas[numeros[i + 1]]
                if any(sum(c.segundos for c in cinta["caras"][k]) > 30 * 60
                       for k in ("A", "B")):
                    cinta["maximo"] = 45
                ordenadas.append(cinta)
    finally:
        wb.close()

    if not ordenadas:
        raise ValueError("La sesión no tiene ninguna cinta con canciones.")
    return ordenadas, omitidas


# ----------------------------------------------------------------------------
# Interfaz
# ----------------------------------------------------------------------------
class App(_BaseTk):
    def __init__(self):
        super().__init__()
        self.title("Cassettotes")
        try:
            self.iconbitmap(ruta_recurso("icon.ico"))
        except Exception:
            pass
        self.geometry("1000x620")
        self.minsize(1000, 620)
        self.maxsize(1000, 620)
        self.resizable(False, False)

        self.cintas = [cinta_vacia()]   # todas las cintas de la sesión
        self.indice_cinta = 0           # cinta que se está editando (0-based)
        self.copiando = False
        self.nombre_sesion = ""
        self.sin_guardar = False        # hay cambios que no están en una sesión guardada
        self.cache = {}                 # ruta -> (fecha de modificación, segundos o None)

        self.maximo = tk.IntVar(value=self.cintas[0]["maximo"])
        self.cara = tk.StringVar(value="A")
        self.var_cinta = tk.StringVar(value="1")
        self.texto_num_cintas = tk.StringVar(value="de 1")
        self.nombre_cinta = tk.StringVar(value="")  # Nombre personalizado de la cinta
        # Sincroniza el nombre de la cinta actual con el campo de texto
        self.nombre_cinta.trace_add("write", lambda *_: self._sincronizar_nombre_cinta())
        # Reproductor de música
        self._reproduciendo = False
        self._pausado = False
        self._idx_reproduccion = None  # índice en self.canciones de la pista cargada
        self._buscando = False
        self._duracion_actual = 0.0
        self._tiempo_inicio = None
        self.texto_total = tk.StringVar(value="")
        self.texto_cinta = tk.StringVar(value="")
        self.texto_estado = tk.StringVar(value="")
        self.texto_sesion = tk.StringVar(value="")

        self._menu()
        self._widgets()
        self._actualizar_navegacion()
        self.protocol("WM_DELETE_WINDOW", self.cerrar)
        self._refrescar()
        self._registrar_dnd()
        if PYGAME_OK:
            self.after(500, self._poll_reproductor)

        faltan = []
        if MutagenFile is None:
            faltan.append("mutagen")
        if Workbook is None:
            faltan.append("openpyxl")
        if faltan:
            messagebox.showerror(
                "Faltan librerías",
                "No se encuentra: " + ", ".join(faltan) + "\n\n"
                "Instálalas con:  pip install " + " ".join(faltan))

    # -- cinta y cara actuales ----------------------------------------------
    @property
    def cinta(self):
        """Cinta que se está editando."""
        return self.cintas[self.indice_cinta]

    @property
    def caras(self):
        """Las dos caras de la cinta actual."""
        return self.cinta["caras"]

    @property
    def cara_actual(self):
        return self.cara.get()

    @property
    def canciones(self):
        """Canciones de la cara que se está editando."""
        return self.caras[self.cara_actual]

    @canciones.setter
    def canciones(self, valor):
        self.caras[self.cara_actual] = list(valor)

    @property
    def lista_registrada(self):
        """True si la cara mostrada ya se copió a una carpeta."""
        return self.cinta["copiada"][self.cara_actual]

    @lista_registrada.setter
    def lista_registrada(self, valor):
        self.cinta["copiada"][self.cara_actual] = bool(valor)

    def _cambiar_cara(self):
        # Guardar el nombre de cinta actual antes de cambiar
        self.cintas[self.indice_cinta]["nombre"] = self.nombre_cinta.get().strip()
        
        self._refrescar()
        # Obtener el nombre de cinta para mostrar (si existe)
        nombre = self.cinta.get("nombre", "").strip()
        parte_nombre = f" - {nombre}" if nombre else ""
        self.texto_estado.set(
            f"Editando la Cara {self.cara_actual} de la cinta {self.indice_cinta + 1}{parte_nombre}.")

    def _cambiar_maximo(self):
        # Guardar el nombre de cinta actual antes de cambiar
        self.cintas[self.indice_cinta]["nombre"] = self.nombre_cinta.get().strip()
        
        self.cinta["maximo"] = self.maximo.get()
        self.sin_guardar = True
        self._refrescar()

    def _sincronizar_nombre_cinta(self):
        if not self.cintas:
            return
        try:
            self.cintas[self.indice_cinta]["nombre"] = self.nombre_cinta.get().strip()
        except (IndexError, KeyError):
            pass

    def _actualizar_navegacion(self):
        n = len(self.cintas)
        self.spin_cinta.config(from_=1, to=n)
        self.texto_num_cintas.set(f"de {n}")
        # Actualizar el campo de entrada con el nombre de la nueva cinta
        if self.cintas:
            self.nombre_cinta.set(self.cinta.get("nombre", ""))

    def _ir_a_cinta(self, numero, avisar=True):
        numero = max(1, min(int(numero), len(self.cintas)))
        self.indice_cinta = numero - 1
        self.maximo.set(self.cinta["maximo"])
        self.var_cinta.set(str(numero))
        # Cargar el nombre de cinta de la nueva cinta (si existe)
        self.nombre_cinta.set(self.cinta.get("nombre", ""))
        self.cara.set("A")  # Cambiar automáticamente a Cara A al cambiar de cinta
        n = len(self.cintas)
        self.spin_cinta.config(from_=1, to=n)
        self.texto_num_cintas.set(f"de {n}")
        self._refrescar()
        if avisar:
            self.texto_estado.set(
                f"Editando la cinta {numero} de {len(self.cintas)}.")

    def _cambiar_cinta_spin(self):
        try:
            numero = int(self.var_cinta.get())
        except ValueError:
            numero = self.indice_cinta + 1
        # Guardar el nombre de cinta actual antes de cambiar
        if self.cintas:
            self.cintas[self.indice_cinta]["nombre"] = self.nombre_cinta.get().strip()
        self._ir_a_cinta(numero)

    def nueva_cinta(self):
        if self.copiando:
            return
        # Guardar el nombre de la cinta actual antes de cambiar
        self.cintas[self.indice_cinta]["nombre"] = self.nombre_cinta.get().strip()
        
        self.cintas.append(cinta_vacia(self.maximo.get()))
        self.sin_guardar = True
        self._ir_a_cinta(len(self.cintas))
        self.texto_estado.set(
            f"Cinta {len(self.cintas)} creada. La sesión tiene "
            f"{len(self.cintas)} cintas.")

    # -- construcción -------------------------------------------------------
    def _menu(self):
        barra = tk.Menu(self)
        archivo = tk.Menu(barra, tearoff=0)
        archivo.add_command(label="Añadir archivos...", command=self.anadir_desde_dialogo)
        archivo.add_separator()
        archivo.add_command(label="Guardar cinta...", command=self.guardar_cinta)
        archivo.add_command(label="Cargar cinta...", command=self.cargar_cinta)
        archivo.add_command(label="Exportar cinta a TXT...", command=self.exportar_cinta_txt)
        archivo.add_separator()
        archivo.add_command(label="Nueva cinta", command=self.nueva_cinta)
        archivo.add_separator()
        archivo.add_command(label="Guardar sesión...", command=self.guardar_sesion)
        archivo.add_command(label="Cargar sesión...", command=self.cargar_sesion)
        barra.add_cascade(label="Archivo", menu=archivo)
        barra.add_command(label="Info", command=self._mostrar_info)
        self.config(menu=barra)

    def _mostrar_info(self):
        ventana = tk.Toplevel(self)
        ventana.title("Info")
        try:
            ventana.iconbitmap(ruta_recurso("icon.ico"))
        except Exception:
            pass
        ventana.resizable(False, False)
        texto = (
            "Cassettotes 1.0\n"
            "ACUERDO DE LICENCIA DE USO (FREEWARE)\n"
            "Copyright (c) 2026 [Alejandro Garrido Trujillo (M\u00e9ldinor)]\n"
            "Contacto: meldinor@gmail.com\n"
            "\n"
            "1. CONCESI\u00d3N DE LICENCIA: Se concede al usuario el derecho gratuito y no exclusivo de descargar, instalar y utilizar este software.\n"
            "2. RESTRICCIONES: El autor conserva todos los derechos de propiedad intelectual. Queda estrictamente prohibido:\n"
            "   - Vender, alquilar, sublicenciar o comercializar el programa o cualquier parte del mismo.\n"
            "   - Copiar, distribuir o alojar el software en plataformas de terceros sin autorizaci\u00f3n previa del autor.\n"
            "   - Descompilar, realizar ingenier\u00eda inversa, modificar o crear obras derivadas a partir de este programa.\n"
            "3. EXENCI\u00d3N DE GARANT\u00cdA: El software se proporciona \"tal cual\", sin garant\u00eda de ning\u00fan tipo. El autor no se hace responsable de posibles da\u00f1os derivados de su uso.\n"
        )
        frame = ttk.Frame(ventana, padding=12)
        frame.pack(fill="both", expand=True)
        txt = tk.Text(frame, wrap="word", width=70, height=20, font=("Segoe UI", 10))
        txt.pack(fill="both", expand=True)
        txt.insert("1.0", texto)
        txt.config(state="disabled")
        ttk.Button(frame, text="Cerrar", command=ventana.destroy).pack(pady=(8, 0))

    def _widgets(self):
        estilo = ttk.Style(self)
        estilo.configure("Player.TButton", font=("Segoe UI Symbol", 12), padding=3)
        marco = ttk.Frame(self, padding=12)
        marco.pack(fill="both", expand=True)
        self.marco = marco

        fila1 = ttk.Frame(marco)
        fila1.pack(fill="x")
        ttk.Label(fila1, text="Duración de la Cara:").pack(side="left")
        ttk.Radiobutton(fila1, text="27 minutos", variable=self.maximo, value=27,
                        command=self._cambiar_maximo).pack(side="left", padx=(6, 0))
        ttk.Radiobutton(fila1, text="30 minutos", variable=self.maximo, value=30,
                        command=self._cambiar_maximo).pack(side="left", padx=(6, 0))
        ttk.Radiobutton(fila1, text="45 minutos", variable=self.maximo, value=45,
                        command=self._cambiar_maximo).pack(side="left", padx=(6, 0))
        ttk.Radiobutton(fila1, text="60 minutos", variable=self.maximo, value=60,
                        command=self._cambiar_maximo).pack(side="left", padx=(6, 0))
        ttk.Radiobutton(fila1, text="90 minutos", variable=self.maximo, value=90,
                        command=self._cambiar_maximo).pack(side="left", padx=(6, 0))

        ttk.Separator(fila1, orient="vertical").pack(side="left", fill="y", padx=12)
        ttk.Label(fila1, text="Cara:").pack(side="left")
        ttk.Radiobutton(fila1, text="Cara A", variable=self.cara, value="A",
                        command=self._cambiar_cara).pack(side="left", padx=(6, 0))
        ttk.Radiobutton(fila1, text="Cara B", variable=self.cara, value="B",
                        command=self._cambiar_cara).pack(side="left", padx=(6, 0))

        fila2 = ttk.Frame(marco)
        fila2.pack(fill="x", pady=10)
        self.btn_borrar_cara = ttk.Button(fila2, text="Borrar Cara", command=self.borrar_cara)
        self.btn_borrar_cara.pack(side="left")
        self.btn_borrar_cinta = ttk.Button(fila2, text="Borrar Cinta", command=self.borrar_cinta)
        self.btn_borrar_cinta.pack(side="left", padx=(8, 0))
        self.btn_borrar_sobrante = ttk.Button(fila2, text="Borrar sobrante",
                                              command=self.borrar_sobrante)
        self.btn_borrar_sobrante.pack(side="left", padx=(8, 0))
        self.btn_copiar = ttk.Button(fila2, text="Copiar a carpeta", command=self.copiar_a_carpeta)
        self.btn_copiar.pack(side="left", padx=(8, 0))

        self.texto_ayuda = tk.StringVar(value="Arrastra aquí los archivos de audio. "
                                               "Cada cinta tiene dos caras (A y B). "
                                               "Abajo a la derecha cambias de cinta. "
                                               "Arrastra las canciones para reordenarlas. "
                                               "Supr quita las canciones seleccionadas.")
        ttk.Label(marco, textvariable=self.texto_ayuda, foreground="#555555").pack(anchor="w", pady=(0, 4))

        cuerpo = ttk.Frame(marco)
        cuerpo.pack(fill="both", expand=True)
        self.tabla = ttk.Treeview(cuerpo, columns=("nombre", "duracion", "formato", "acumulado"),
                                  show="headings", selectmode="extended")
        self.tabla.heading("nombre", text="Canción")
        self.tabla.heading("duracion", text="Duración")
        self.tabla.heading("formato", text="Formato")
        self.tabla.heading("acumulado", text="Acumulado")
        self.tabla.column("nombre", anchor="w", width=320, stretch=True)
        self.tabla.column("duracion", anchor="center", width=80, stretch=False)
        self.tabla.column("formato", anchor="center", width=70, stretch=False)
        self.tabla.column("acumulado", anchor="center", width=90, stretch=False)
        self.tabla.tag_configure("excede", foreground=COLOR_ROJO)
        scroll = ttk.Scrollbar(cuerpo, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tabla.bind("<Delete>", self.quitar_seleccion)
        self.tabla.bind("<ButtonPress-1>", self._tabla_press)
        self.tabla.bind("<B1-Motion>", self._tabla_motion)
        self.tabla.bind("<ButtonRelease-1>", self._tabla_release)
        self.tabla.bind("<Double-1>", self._doble_click)
        self.tabla.bind("<Button-3>", self._menu_contextual)
        self._arrastre_iid = None
        self._arrastre_movido = False

        pie = ttk.Frame(marco)
        pie.pack(fill="x", pady=(8, 0))

        # Controles del reproductor + campo de nombre (al margen derecho e inferior)
        marco_nombre = ttk.Frame(pie)
        marco_nombre.pack(side="right", anchor="s", padx=(12, 0))
        # Fila de controles del reproductor (sobre el campo de nombre)
        fila_player = ttk.Frame(marco_nombre)
        fila_player.pack(anchor="e")
        _fuente_player = ("Segoe UI Symbol", 24)
        ttk.Button(fila_player, text="\u23EE", width=2, command=self._retroceder,
                   style="Player.TButton").pack(side="left", padx=1)
        self.btn_play = ttk.Button(fila_player, text="\u25B6", width=2, command=self._play_pause,
                                   style="Player.TButton")
        self.btn_play.pack(side="left", padx=1)
        ttk.Button(fila_player, text="\u23ED", width=2, command=self._avanzar,
                   style="Player.TButton").pack(side="left", padx=1)
        ttk.Button(fila_player, text="\u23F9", width=2, command=self._parar,
                   style="Player.TButton").pack(side="left", padx=1)
        # Barra de progreso a la derecha de los botones
        frame_progreso = ttk.Frame(fila_player)
        frame_progreso.pack(side="left", padx=(8, 0))
        self.barra_progreso = ttk.Scale(frame_progreso, from_=0, to=1, orient="horizontal",
                                       length=140, command=self._on_seek_drag)
        self.barra_progreso.pack()
        self.barra_progreso.bind("<ButtonPress-1>", self._on_seek_press)
        self.barra_progreso.bind("<ButtonRelease-1>", self._on_seek_release)
        self.lbl_progreso = ttk.Label(frame_progreso, text="00:00 / 00:00",
                                     font=("Segoe UI", 9))
        self.lbl_progreso.pack()
        if not PYGAME_OK:
            for w in fila_player.winfo_children():
                w.config(state="disabled")
        # Fila del nombre
        fila_nombre = ttk.Frame(marco_nombre)
        fila_nombre.pack(anchor="e", pady=(4, 0))
        ttk.Label(fila_nombre, text="Nombre:", font=("Segoe UI", 9)).pack(side="left")
        self.entry_nombre_cinta = ttk.Entry(fila_nombre, textvariable=self.nombre_cinta,
                                            width=25)
        self.entry_nombre_cinta.pack(side="left", padx=(4, 0))

        derecha = ttk.Frame(pie)
        derecha.pack(side="right", anchor="s", padx=(12, 0))
        ttk.Label(derecha, text="Cinta:", font=("Segoe UI", 10, "bold")).pack(side="left")
        self.spin_cinta = ttk.Spinbox(derecha, from_=1, to=1, width=4,
                                      justify="center", state="readonly",
                                      textvariable=self.var_cinta,
                                      command=self._cambiar_cinta_spin)
        self.spin_cinta.pack(side="left", padx=(4, 2))
        ttk.Label(derecha, textvariable=self.texto_num_cintas,
                  foreground="#555555").pack(side="left", padx=(0, 8))
        self.btn_nueva_cinta = ttk.Button(derecha, text="Nueva cinta",
                                          command=self.nueva_cinta)
        self.btn_nueva_cinta.pack(side="left")

        izquierda = ttk.Frame(pie)
        izquierda.pack(side="left", fill="x", expand=True)

        self.lbl_total = ttk.Label(izquierda, textvariable=self.texto_total,
                                   font=("Segoe UI", 10, "bold"))
        self.lbl_total.pack(anchor="w")
        ttk.Label(izquierda, textvariable=self.texto_cinta,
                  foreground="#333333").pack(anchor="w")
        ttk.Label(izquierda, textvariable=self.texto_estado,
                  foreground="#555555").pack(anchor="w")
        ttk.Label(izquierda, textvariable=self.texto_sesion,
                  foreground="#555555").pack(anchor="w")

    def _registrar_dnd(self):
        if not DND_OK:
            self.texto_ayuda.set("Arrastrar y soltar no está disponible en este equipo: "
                                 "usa Archivo > Añadir archivos.")
            return
        try:
            for w in (self, self.marco, self.tabla):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._soltar)
        except Exception:
            self.texto_ayuda.set("Arrastrar y soltar no está disponible en este equipo: "
                                 "usa Archivo > Añadir archivos.")

    # -- estado -----------------------------------------------------------------
    def _ocupado(self, si):
        estado = "disabled" if si else "normal"
        self.btn_copiar.config(state=estado)
        self.btn_borrar_cara.config(state=estado)
        self.btn_borrar_sobrante.config(state=estado)
        self.btn_borrar_cinta.config(state=estado)
        self.btn_nueva_cinta.config(state=estado)
        self.spin_cinta.config(state="disabled" if si else "readonly")

    def claves_excluidas(self):
        """Canciones que no se pueden volver a añadir (ya están en otra cinta)."""
        claves = set()
        for i, cinta in enumerate(self.cintas):
            if i == self.indice_cinta:
                continue  # la cinta actual se controla con las dos caras visibles
            for clave in ("A", "B"):
                for c in cinta["caras"][clave]:
                    claves.add(clave_cancion(c.ruta, c.segundos))
        return frozenset(claves)

    def _actualizar_sesion(self):
        total = sum(len(c["caras"]["A"]) + len(c["caras"]["B"]) for c in self.cintas)
        n = len(self.cintas)
        palabra = "cinta" if n == 1 else "cintas"
        partes = []
        if self.nombre_sesion:
            partes.append(f"Sesión: {self.nombre_sesion}")
        
        # Añadir información de la cinta actual
        nombre_cinta_actual = self.cinta.get("nombre", "").strip()
        cinta_num = self.indice_cinta + 1
        total_cintas = len(self.cintas)
        if nombre_cinta_actual:
            partes.append(f"Cinta {cinta_num}/{total_cintas}: {nombre_cinta_actual}")
        else:
            partes.append(f"{n} {palabra} · {total} canciones")
        
        if self.sin_guardar:
            partes.append("cambios sin guardar")
        self.texto_sesion.set(" · ".join(partes))

    def _actualizar_cinta(self):
        """Resumen de las dos caras de la cinta."""
        # Obtener el nombre de cinta para mostrar (si existe)
        nombre = self.cinta.get("nombre", "").strip()
        parte_nombre = f" - {nombre}" if nombre else ""
        
        partes = []
        for clave in ("A", "B"):
            total = sum(c.segundos for c in self.caras[clave])
            n = len(self.caras[clave])
            palabra = "canción" if n == 1 else "canciones"
            marca = " (editando)" if clave == self.cara_actual else ""
            partes.append(f"Cara {clave}: {formato_tiempo(total)}, {n} {palabra}{marca}")
        # Actualizar el texto para incluir número de cinta y nombre
        cinta_num = self.indice_cinta + 1
        total_cintas = len(self.cintas)
        info_cinta = f"{nombre}" if nombre else f"Cinta {cinta_num}/{total_cintas}"
        self.texto_cinta.set(f"{info_cinta} · " + " · ".join(partes))

    def _refrescar(self):
        """Redibuja la tabla y marca en rojo lo que supera el máximo elegido."""
        limite = self.maximo.get() * 60
        self._actualizar_cinta()
        self.tabla.delete(*self.tabla.get_children())
        acumulado, excedidas = 0, 0
        for i, c in enumerate(self.canciones):
            acumulado += c.segundos
            pasa = acumulado > limite
            excedidas += 1 if pasa else 0
            formato = os.path.splitext(c.ruta)[1].lstrip(".").upper() or "-"
            self.tabla.insert("", "end", iid=str(i),
                              values=(c.nombre, formato_tiempo(c.segundos), formato, formato_tiempo(acumulado)),
                              tags=("excede",) if pasa else ())

        if not self.canciones:
            self.texto_total.set("")
            self.lbl_total.config(foreground=COLOR_NORMAL)
            self._actualizar_sesion()
            return
        n = len(self.canciones)
        palabra = "canción" if n == 1 else "canciones"
        if acumulado > limite:
            self.texto_total.set(
                f"Total: {formato_tiempo(acumulado)} de {formato_tiempo(limite)} · "
                f"te pasas {formato_tiempo(acumulado - limite)} · en rojo: {excedidas}")
            self.lbl_total.config(foreground=COLOR_ROJO)
        else:
            self.texto_total.set(
                f"Total: {formato_tiempo(acumulado)} de {formato_tiempo(limite)} · "
                f"quedan {formato_tiempo(limite - acumulado)} ({n} {palabra})")
            self.lbl_total.config(foreground=COLOR_NORMAL)
        self._actualizar_sesion()

    # -- añadir canciones -------------------------------------------------------
    def _soltar(self, evento):
        rutas = [os.path.normpath(r) for r in self.tk.splitlist(evento.data)]
        self.anadir_archivos(rutas)
        return evento.action

    def anadir_desde_dialogo(self):
        patron = " ".join("*" + e for e in sorted(EXTENSIONES))
        rutas = filedialog.askopenfilenames(
            title="Añadir archivos de audio",
            filetypes=[("Archivos de audio", patron), ("Todos los archivos", "*.*")],
        )
        if rutas:
            self.anadir_archivos([os.path.normpath(r) for r in rutas])

    def anadir_archivos(self, rutas):
        if self.copiando:
            self.texto_estado.set("Espera a que termine la copia para añadir canciones.")
            return
        if MutagenFile is None:
            messagebox.showerror("Falta mutagen", "Instala la librería con:  pip install mutagen")
            return
        if not rutas:
            return
        self.texto_estado.set("Leyendo archivos...")
        cola = queue.Queue()
        threading.Thread(target=self._trabajo_lectura, args=(cola, list(rutas)), daemon=True).start()
        self.after(100, self._esperar_lectura, cola)

    def _trabajo_lectura(self, cola, rutas):
        try:
            audios, ignorados = expandir_rutas(rutas)
            canciones, ilegibles = leer_canciones(audios, self.cache)
            cola.put(("ok", canciones, ignorados, ilegibles))
        except Exception as e:  # noqa: BLE001
            cola.put(("error", str(e)))

    def _esperar_lectura(self, cola):
        try:
            res = cola.get_nowait()
        except queue.Empty:
            self.after(100, self._esperar_lectura, cola)
            return
        if res[0] == "error":
            self.texto_estado.set("")
            messagebox.showerror("Error", f"No se pudieron leer los archivos:\n{res[1]}")
            return
        _, canciones, ignorados, ilegibles = res
        self._aplicar_nuevas(canciones, ignorados + ilegibles)

    def _aplicar_nuevas(self, canciones, no_validos):
        excluidas = self.claves_excluidas()
        actuales = {
            clave_cancion(c.ruta, c.segundos)
            for lista in self.caras.values()
            for c in lista
        }
        nuevas, repetidas, usadas, vistas = [], 0, [], set()
        for c in canciones:
            clave = clave_cancion(c.ruta, c.segundos)
            if esta_excluida(c, excluidas):
                usadas.append(c.nombre)
            elif clave in actuales or clave in vistas:
                repetidas += 1
            else:
                nuevas.append(c)
                vistas.add(clave)

        if nuevas:
            # Guardar el nombre de cinta actual antes de cambiar
            self.cintas[self.indice_cinta]["nombre"] = self.nombre_cinta.get().strip()
            
            if self.lista_registrada:  # la cara anterior ya se copió: empieza una nueva
                self.canciones = []
                self.lista_registrada = False
            self.canciones.extend(nuevas)
            self.sin_guardar = True
            self._refrescar()

        partes = [f"Añadidas: {len(nuevas)}"]
        if repetidas:
            partes.append(f"ya estaban: {repetidas}")
        if usadas:
            partes.append(f"ya usadas: {len(usadas)}")
        if no_validos:
            partes.append(f"no válidas o ilegibles: {no_validos}")
        self.texto_estado.set(" · ".join(partes))

        if usadas:
            detalle = "\n".join(f"- {n}" for n in usadas[:5])
            if len(usadas) > 5:
                detalle += f"\n... y {len(usadas) - 5} más"
            messagebox.showinfo(
                "Canciones excluidas",
                f"No se añadieron {len(usadas)} canciones porque ya están en otra "
                f"cinta de la sesión:\n{detalle}")

    def quitar_seleccion(self, evento=None):
        seleccion = self.tabla.selection()
        if not seleccion or self.copiando:
            return
        
        # Guardar el nombre de cinta actual antes de cambiar
        self.cintas[self.indice_cinta]["nombre"] = self.nombre_cinta.get().strip()
        
        quitar = {int(i) for i in seleccion}
        self.canciones = [c for i, c in enumerate(self.canciones) if i not in quitar]
        self.sin_guardar = True
        self._refrescar()

    def _tabla_press(self, evento):
        if self.copiando:
            self._arrastre_iid = None
            return
        iid = self.tabla.identify_row(evento.y)
        self._arrastre_iid = iid if iid else None
        self._arrastre_movido = False

    def _tabla_motion(self, evento):
        if not self._arrastre_iid or self.copiando:
            return
        iid = self._arrastre_iid
        destino = self.tabla.identify_row(evento.y)
        if not destino or destino == iid:
            return
        try:
            idx_destino = self.tabla.index(destino)
        except Exception:
            return
        self.tabla.move(iid, '', idx_destino)
        self._arrastre_movido = True

    def _tabla_release(self, evento):
        iid = self._arrastre_iid
        self._arrastre_iid = None
        if not iid or not self._arrastre_movido or self.copiando:
            self._arrastre_movido = False
            return
        self._arrastre_movido = False
        try:
            nuevo_orden = [int(x) for x in self.tabla.children('')]
        except ValueError:
            self._refrescar()
            return
        canciones = self.canciones
        if len(nuevo_orden) != len(canciones):
            self._refrescar()
            return
        if any(i < 0 or i >= len(canciones) for i in nuevo_orden):
            self._refrescar()
            return
        self.canciones = [canciones[i] for i in nuevo_orden]
        self.sin_guardar = True
        self._refrescar()
        try:
            self.tabla.selection_set(iid)
        except Exception:
            pass

    # -- reproductor de musica -------------------------------------------------
    def _indice_seleccionado(self):
        sel = self.tabla.selection()
        if sel:
            try:
                return int(sel[0])
            except ValueError:
                pass
        if self._idx_reproduccion is not None:
            return self._idx_reproduccion
        return 0

    def _play_idx(self, idx):
        if not self.canciones or not PYGAME_OK:
            return
        n = len(self.canciones)
        idx = idx % n
        c = self.canciones[idx]
        try:
            pygame.mixer.music.load(c.ruta)
            pygame.mixer.music.play()
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Reproductor", "No se pudo reproducir:\n" + c.nombre + "\n" + str(e))
            return
        self._idx_reproduccion = idx
        self._reproduciendo = True
        self._pausado = False
        # Duración y barra de progreso
        try:
            self._duracion_actual = float(c.segundos)
        except Exception:
            self._duracion_actual = 0.0
        self._posicion_actual = 0.0
        self._tiempo_inicio = time.time()
        try:
            self.barra_progreso.config(to=self._duracion_actual if self._duracion_actual > 0 else 1)
            self.barra_progreso.set(0)
        except Exception:
            pass
        self.lbl_progreso.config(text=f"00:00 / {self._formatear_mmss(self._duracion_actual)}")
        self.tabla.selection_set(str(idx))
        try:
            self.tabla.see(str(idx))
        except Exception:
            pass
        self._actualizar_botones_reproductor()

    def _play_pause(self):
        if not PYGAME_OK:
            return
        if self._pausado:
            pygame.mixer.music.unpause()
            self._pausado = False
            self._tiempo_inicio = time.time()
            self._actualizar_botones_reproductor()
            return
        if self._reproduciendo:
            if self._tiempo_inicio is not None:
                self._posicion_actual += time.time() - self._tiempo_inicio
            self._tiempo_inicio = None
            pygame.mixer.music.pause()
            self._pausado = True
            self._actualizar_botones_reproductor()
            return
        self._play_idx(self._indice_seleccionado())

    def _retroceder(self):
        if not self.canciones or not PYGAME_OK:
            return
        idx = self._idx_reproduccion if self._idx_reproduccion is not None else self._indice_seleccionado()
        self._play_idx(idx - 1)

    def _avanzar(self):
        if not self.canciones or not PYGAME_OK:
            return
        idx = self._idx_reproduccion if self._idx_reproduccion is not None else self._indice_seleccionado()
        self._play_idx(idx + 1)

    def _parar(self):
        if not PYGAME_OK:
            return
        pygame.mixer.music.stop()
        self._reproduciendo = False
        self._pausado = False
        self._posicion_actual = 0.0
        self._tiempo_inicio = None
        try:
            self.barra_progreso.set(0)
        except Exception:
            pass
        self.lbl_progreso.config(text=f"00:00 / {self._formatear_mmss(self._duracion_actual)}")
        self._actualizar_botones_reproductor()

    def _doble_click(self, evento):
        if self.copiando or not PYGAME_OK:
            return
        iid = self.tabla.identify_row(evento.y)
        if not iid:
            return
        try:
            idx = int(iid)
        except ValueError:
            return
        self._play_idx(idx)

    def _menu_contextual(self, evento):
        if self.copiando:
            return
        iid = self.tabla.identify_row(evento.y)
        if not iid:
            return
        try:
            idx = int(iid)
        except ValueError:
            return
        self.tabla.selection_set(iid)
        cara_actual = self.cara_actual
        cara_destino = "B" if cara_actual == "A" else "A"

        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Borrar", command=lambda: self._borrar_cancion_idx(idx))
        menu.add_command(label="Pasar a Cara " + cara_destino,
                         command=lambda: self._pasar_a_cara(idx, cara_destino))
        try:
            menu.tk_popup(evento.x_root, evento.y_root)
        finally:
            menu.grab_release()

    def _borrar_cancion_idx(self, idx):
        canciones = self.canciones
        if idx < 0 or idx >= len(canciones):
            return
        self._sincronizar_nombre_cinta()
        del canciones[idx]
        self.sin_guardar = True
        self._refrescar()

    def _pasar_a_cara(self, idx, cara_destino):
        canciones = self.canciones
        if idx < 0 or idx >= len(canciones):
            return
        self._sincronizar_nombre_cinta()
        cancion = canciones[idx]
        del canciones[idx]
        self.cintas[self.indice_cinta]["caras"][cara_destino].append(cancion)
        self.sin_guardar = True
        self._refrescar()

    def _formatear_mmss(self, segundos):
        segundos = max(0, int(segundos))
        m, s = divmod(segundos, 60)
        return f"{m:02d}:{s:02d}"

    def _on_seek_press(self, evento):
        self._buscando = True

    def _on_seek_drag(self, valor):
        try:
            v = float(valor)
        except (TypeError, ValueError):
            return
        self._posicion_actual = v
        total = self._duracion_actual if self._duracion_actual > 0 else 0
        self.lbl_progreso.config(text=f"{self._formatear_mmss(v)} / {self._formatear_mmss(total)}")

    def _on_seek_release(self, evento):
        self._buscando = False
        try:
            v = float(self.barra_progreso.get())
        except (TypeError, ValueError):
            return
        v = max(0.0, min(v, self._duracion_actual if self._duracion_actual > 0 else v))
        self._posicion_actual = v
        if PYGAME_OK and (self._reproduciendo or self._pausado):
            was_paused = self._pausado
            c = None
            if self._idx_reproduccion is not None and self.canciones:
                try:
                    c = self.canciones[self._idx_reproduccion]
                except IndexError:
                    c = None
            if c is not None:
                try:
                    pygame.mixer.music.stop()
                    pygame.mixer.music.load(c.ruta)
                    pygame.mixer.music.play(start=v)
                    if was_paused:
                        pygame.mixer.music.pause()
                        self._pausado = True
                        self._tiempo_inicio = None
                    else:
                        self._reproduciendo = True
                        self._pausado = False
                        self._tiempo_inicio = time.time()
                except Exception:
                    # Fallback a set_pos
                    try:
                        pygame.mixer.music.set_pos(v)
                    except Exception:
                        pass
                    if not was_paused:
                        self._tiempo_inicio = time.time()
        total = self._duracion_actual if self._duracion_actual > 0 else 0
        self.lbl_progreso.config(text=f"{self._formatear_mmss(v)} / {self._formatear_mmss(total)}")

    def _actualizar_botones_reproductor(self):
        if not hasattr(self, 'btn_play'):
            return
        if self._reproduciendo and not self._pausado:
            self.btn_play.config(text='\u23F8')
        else:
            self.btn_play.config(text='\u25B6')

    def _poll_reproductor(self):
        if PYGAME_OK and self._reproduciendo and not self._pausado:
            try:
                if not self._buscando:
                    if self._tiempo_inicio is not None:
                        pos = self._posicion_actual + (time.time() - self._tiempo_inicio)
                    else:
                        pos = self._posicion_actual
                    if self._duracion_actual > 0:
                        pos = min(pos, self._duracion_actual)
                    # Reiniciar el punto de partida para no acumular
                    self._posicion_actual = pos
                    self._tiempo_inicio = time.time()
                    try:
                        self.barra_progreso.set(pos)
                    except Exception:
                        pass
                    total = self._duracion_actual if self._duracion_actual > 0 else 0
                    self.lbl_progreso.config(text=f"{self._formatear_mmss(pos)} / {self._formatear_mmss(total)}")
                if not pygame.mixer.music.get_busy():
                    # Avanzar a la siguiente canción si no es la última
                    idx = self._idx_reproduccion
                    canciones = self.canciones
                    if idx is not None and canciones and idx + 1 < len(canciones):
                        self._play_idx(idx + 1)
                    else:
                        # Última canción: parar sin bucle
                        self._reproduciendo = False
                        self._pausado = False
                        self._actualizar_botones_reproductor()
            except Exception:
                pass
        self.after(500, self._poll_reproductor)

    def borrar_cara(self):
        """Borra solo la cara que se está editando (A o B)."""
        if self.copiando:
            return
        if not self.canciones:
            self.texto_estado.set(f"La Cara {self.cara_actual} ya está vacía.")
            return
        
        # Guardar el nombre de cinta actual antes de cambiar
        self.cintas[self.indice_cinta]["nombre"] = self.nombre_cinta.get().strip()
        
        cara = self.cara_actual
        self.canciones = []
        self.lista_registrada = False
        self.sin_guardar = True
        self._refrescar()
        self.texto_estado.set(f"Cara {cara} borrada.")

    def borrar_sobrante(self):
        """Borra las canciones que no caben en el máximo, en las dos caras."""
        if self.copiando:
            return
        limite = self.maximo.get() * 60
        sobrantes = {}
        for clave in ("A", "B"):
            acumulado = 0
            fuera = []
            for i, c in enumerate(self.caras[clave]):
                acumulado += c.segundos
                if acumulado > limite:
                    fuera.append(i)
            if fuera:
                sobrantes[clave] = fuera

        if not sobrantes:
            self.texto_estado.set(
                f"Nada que borrar: todo cabe en {self.maximo.get()} minutos por cara.")
            return

        total = sum(len(v) for v in sobrantes.values())
        detalle = ", ".join(f"Cara {clave}: {len(v)}" for clave, v in sobrantes.items())
        if not messagebox.askyesno(
                "Borrar sobrante",
                f"Se quitarán {total} canciones que no caben en el máximo de "
                f"{self.maximo.get()} minutos por cara ({detalle}).\n\n"
                "El resto quedará dentro del máximo.\n\n¿Continuar?"):
            return

        for clave, indices in sobrantes.items():
            quitar = set(indices)
            self.caras[clave] = [c for i, c in enumerate(self.caras[clave])
                                 if i not in quitar]
        self.sin_guardar = True
        self._refrescar()
        self.texto_estado.set(
            f"Borradas {total} canciones que no cabían ({detalle}).")

    def borrar_cinta(self):
        """Borra las dos caras de la cinta actual."""
        if self.copiando:
            return
        if not cinta_tiene_canciones(self.cinta):
            self.texto_estado.set("La cinta ya está vacía.")
            return
        if not messagebox.askyesno(
                "Borrar cinta",
                f"Se borrarán las dos caras de la cinta {self.indice_cinta + 1}."
                "\n\n¿Continuar?"):
            return
        self.cinta["caras"] = {"A": [], "B": []}
        self.cinta["copiada"] = {"A": False, "B": False}
        self.sin_guardar = True
        self._refrescar()
        self.texto_estado.set(f"Cinta {self.indice_cinta + 1} borrada.")

    # -- copiar a carpeta -------------------------------------------------------
    def copiar_a_carpeta(self):
        # Guardar el nombre de la cinta actual antes de copiar
        self.cintas[self.indice_cinta]["nombre"] = self.nombre_cinta.get().strip()

        # Cintas que tienen canciones
        cintas_a_copiar = [(i, c) for i, c in enumerate(self.cintas)
                           if cinta_tiene_canciones(c)]
        if not cintas_a_copiar:
            messagebox.showinfo(
                "Copiar a carpeta",
                "No hay cintas con canciones. Arrastra primero algunas canciones.")
            return

        # Comprobar el limite en todas las cintas
        limite = self.maximo.get() * 60
        pasadas = []
        for i, c in cintas_a_copiar:
            for clave in ("A", "B"):
                total = sum(x.segundos for x in c["caras"][clave])
                if total > limite:
                    pasadas.append(f"Cinta {i + 1} Cara {clave} ({formato_tiempo(total)})")
        if pasadas:
            detalle = ", ".join(pasadas)
            if not messagebox.askyesno(
                    "Copiar a carpeta",
                    f"Supera el máximo de {self.maximo.get()} minutos: {detalle}.\n"
                    "Las canciones en rojo no caben.\n\n"
                    "Copiar las cintas enteras igualmente?"):
                return

        destino = filedialog.askdirectory(
            title="Selecciona la carpeta donde copiar las cintas")
        if not destino:
            return
        destino = os.path.normpath(destino)

        # Prohibición: el destino no puede ser una carpeta donde ya están las canciones
        carpetas_destino = []
        for i, c in cintas_a_copiar:
            nombre_cinta = (c.get("nombre") or "").strip()
            carpeta_cinta = os.path.join(destino, nombre_carpeta_cinta(i, nombre_cinta))
            for clave in ("A", "B"):
                carpetas_destino.append(clave_ruta(os.path.join(carpeta_cinta, f"Cara {clave}")))
        prohibidas = {clave_ruta(destino)} | set(carpetas_destino)
        todas_canciones = [cancion for i, cint in cintas_a_copiar
                           for clave in ("A", "B") for cancion in cint["caras"][clave]]
        if any(clave_ruta(os.path.dirname(c.ruta)) in prohibidas for c in todas_canciones):
            messagebox.showerror(
                "Copiar a carpeta",
                "La carpeta de destino no puede ser una carpeta donde ya están las canciones.\n"
                "Elige otra.")
            return

        self.copiando = True
        self._ocupado(True)
        self.texto_estado.set("Copiando las cintas (Cara A y Cara B)...")
        cola = queue.Queue()
        datos = [(i, {k: list(c["caras"][k]) for k in ("A", "B")},
                  (c.get("nombre") or "").strip())
                 for i, c in cintas_a_copiar]
        indices = [i for i, _ in cintas_a_copiar]
        threading.Thread(target=self._trabajo_copia,
                         args=(cola, datos, destino), daemon=True).start()
        self.after(100, self._esperar_copia, cola, destino, indices)

    def _trabajo_copia(self, cola, datos, destino):
        try:
            resultado = {}
            for indice, caras, nombre_cinta in datos:
                resultado[indice] = copiar_cinta(caras, destino, indice, nombre_cinta)
            cola.put(("ok", resultado))
        except Exception as e:  # noqa: BLE001
            cola.put(("error", str(e)))

    def _esperar_copia(self, cola, destino, indices):
        try:
            res = cola.get_nowait()
        except queue.Empty:
            self.after(100, self._esperar_copia, cola, destino, indices)
            return

        self.copiando = False
        self._ocupado(False)
        if res[0] == "error":
            self.texto_estado.set("")
            messagebox.showerror("Copiar a carpeta", f"No se pudo copiar:\n{res[1]}")
            return

        resultado = res[1]
        total_copiadas, total_fallos, fallos = 0, 0, []
        for indice in indices:
            cinta = self.cintas[indice]
            r = resultado.get(indice, {})
            for clave in ("A", "B"):
                copiadas, errores = r.get(clave, ([], []))
                total_copiadas += len(copiadas)
                total_fallos += len(errores)
                fallos.extend((indice, clave, nombre, error) for nombre, error in errores)
                if copiadas:
                    cinta["copiada"][clave] = True
        self.sin_guardar = True

        self._refrescar()
        self.texto_estado.set(
            f"{len(indices)} cinta(s) copiada(s) a {destino} "
            f"({total_copiadas} canciones).")

        aviso = "Recuerda borrar esta carpeta cuando termines."
        if fallos:
            detalle = "\n".join(
                f"- Cinta {ind + 1} Cara {clave}: {nombre}: {error}"
                for ind, clave, nombre, error in fallos[:5])
            if len(fallos) > 5:
                detalle += f"\n... y {len(fallos) - 5} más"
            if total_copiadas:
                messagebox.showwarning(
                    "Copiar a carpeta",
                    f"No se pudieron copiar {total_fallos} canciones:\n{detalle}\n\n{aviso}")
            else:
                messagebox.showerror("Copiar a carpeta",
                                     f"No se pudo copiar ninguna canción:\n{detalle}")
        else:
            messagebox.showinfo("Copia terminada", aviso)

    # -- guardar / cargar cinta (.cinta) ----------------------------------------
    def guardar_cinta(self):
        self._sincronizar_nombre_cinta()
        if not cinta_tiene_canciones(self.cinta):
            messagebox.showinfo(
                "Guardar cinta",
                "La cinta está vacía. Arrastra primero algunas canciones.")
            return
        numero = self.indice_cinta + 1
        nombre_cinta = self.cinta.get("nombre", "")
        ruta = filedialog.asksaveasfilename(
            title="Guardar cinta",
            defaultextension=".cinta",
            initialfile=f"{nombre_cinta if nombre_cinta else 'cinta'} {numero}.cinta",
            filetypes=[("Cinta de Cassettotes", "*.cinta"),
                       ("Archivo JSON", "*.json")],
        )
        if not ruta:
            return
        try:
            guardar_cinta_json(ruta, self.cinta)
        except OSError as e:
            messagebox.showerror("Guardar cinta", f"No se pudo guardar el archivo:\n{e}")
            return
        self.texto_estado.set(
            f"Cinta {numero} guardada en {os.path.basename(ruta)} "
            "(las dos caras y todos sus datos).")

    def cargar_cinta(self):
        if self.copiando:
            return
        ruta = filedialog.askopenfilename(
            title="Cargar cinta",
            filetypes=[("Cinta de Cassettotes", "*.cinta"),
                       ("Archivo JSON", "*.json"),
                       ("Todos los archivos", "*.*")],
        )
        if not ruta:
            return
        try:
            cinta = cargar_cinta_json(ruta)
        except ValueError as e:
            messagebox.showerror("Cargar cinta", str(e))
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Cargar cinta", f"No se pudo leer el archivo:\n{e}")
            return

        if not cinta_tiene_canciones(self.cinta):
            self.cintas[self.indice_cinta] = cinta
            numero = self.indice_cinta + 1
        else:
            self.cintas.append(cinta)
            numero = len(self.cintas)
        self.sin_guardar = True
        self._ir_a_cinta(numero)
        self.texto_estado.set(
            f"Cinta cargada desde {os.path.basename(ruta)} en la posición {numero}.")

        faltan = [c for clave in ("A", "B") for c in cinta["caras"][clave]
                  if not os.path.exists(c.ruta)]
        if faltan:
            messagebox.showwarning(
                "Cargar cinta",
                f"{len(faltan)} de las canciones de la cinta ya no están en su "
                "ruta original. Se han cargado igualmente, pero no se podrán "
                "copiar hasta que sus archivos vuelvan a estar ahí.")

    def exportar_cinta_txt(self):
        self._sincronizar_nombre_cinta()
        if not cinta_tiene_canciones(self.cinta):
            messagebox.showinfo(
                "Exportar cinta a TXT",
                "La cinta está vacía. Arrastra primero algunas canciones.")
            return
        numero = self.indice_cinta + 1
        nombre_cinta = self.cinta.get("nombre", "")
        ruta = filedialog.asksaveasfilename(
            title="Exportar cinta a TXT",
            defaultextension=".txt",
            initialfile=f"{nombre_cinta if nombre_cinta else 'cinta'} {numero}.txt",
            filetypes=[("Archivo de texto", "*.txt")],
        )
        if not ruta:
            return
        try:
            guardar_cinta_txt(ruta, self.caras, self.cinta["maximo"], nombre_cinta)
        except OSError as e:
            messagebox.showerror("Exportar cinta a TXT",
                                 f"No se pudo guardar el archivo:\n{e}")
            return
        self.texto_estado.set(
            f"Cinta {numero} exportada a {os.path.basename(ruta)}.")

    # -- sesión -----------------------------------------------------------------
    def guardar_sesion(self):
        self._sincronizar_nombre_cinta()
        if Workbook is None:
            messagebox.showerror("Falta openpyxl", "Instala la librería con:  pip install openpyxl")
            return
        if not any(cinta_tiene_canciones(c) for c in self.cintas):
            messagebox.showinfo(
                "Guardar sesión",
                "No hay nada que guardar: todas las cintas están vacías.")
            return

        ruta = filedialog.asksaveasfilename(
            title="Guardar sesión",
            defaultextension=".xlsx",
            initialfile=self.nombre_sesion or "sesion.xlsx",
            filetypes=[("Hoja de cálculo de Excel", "*.xlsx")],
        )
        if not ruta:
            return
        try:
            guardar_sesion_xlsx(ruta, self.cintas)
        except PermissionError:
            messagebox.showerror(
                "Guardar sesión",
                "No se pudo guardar el archivo. Si lo tienes abierto en Excel, ciérralo "
                "e inténtalo de nuevo.")
            return
        except OSError as e:
            messagebox.showerror("Guardar sesión", f"No se pudo guardar el archivo:\n{e}")
            return

        self.sin_guardar = False
        self.nombre_sesion = os.path.basename(ruta)
        total = sum(len(c["caras"]["A"]) + len(c["caras"]["B"]) for c in self.cintas)
        self.texto_estado.set(
            f"Sesión guardada en {self.nombre_sesion} ({len(self.cintas)} cintas, "
            f"{total} canciones)")
        self._refrescar()

    def cargar_sesion(self):
        if Workbook is None:
            messagebox.showerror("Falta openpyxl", "Instala la librería con:  pip install openpyxl")
            return
        if self.copiando:
            return
        if any(cinta_tiene_canciones(c) for c in self.cintas) and not messagebox.askyesno(
                "Cargar sesión",
                "Se reemplazarán las cintas que hay ahora por las de la sesión.\n"
                "Si tienes cambios sin guardar, se perderán.\n\n"
                "¿Cargar la sesión?"):
            return
        ruta = filedialog.askopenfilename(
            title="Cargar sesión",
            filetypes=[("Hoja de cálculo de Excel", "*.xlsx")],
        )
        if not ruta:
            return
        try:
            cintas, omitidas = cargar_sesion_xlsx(ruta)
        except ValueError as e:
            messagebox.showerror("Cargar sesión", str(e))
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Cargar sesión", f"No se pudo leer el archivo:\n{e}")
            return

        self.cintas = cintas
        self.nombre_sesion = os.path.basename(ruta)
        self.cara.set("A")
        self.indice_cinta = 0
        self.sin_guardar = False
        self._ir_a_cinta(1, avisar=False)
        self.texto_estado.set(
            f"Sesión cargada: {len(cintas)} cintas. Usa las flechas de 'Cinta' "
            "para verlas.")
        if omitidas:
            messagebox.showwarning(
                "Cargar sesión", f"Se omitieron {omitidas} filas con datos incompletos.")

    def cerrar(self):
        hay = any(cinta_tiene_canciones(c) for c in self.cintas)
        if self.sin_guardar and hay and not messagebox.askyesno(
                "Salir",
                "Hay cintas con cambios que no se han guardado en una sesión.\n"
                "Si sales, esos cambios se perderán.\n\n"
                "¿Salir igualmente?",
                icon="warning"):
            return
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
