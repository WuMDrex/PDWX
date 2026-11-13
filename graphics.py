import os
import re
import math
import ezdxf
from ezdxf.addons import importer
from ezdxf.enums import TextEntityAlignment
from ezdxf.addons.drawing import RenderContext, Frontend, config
from ezdxf.addons.drawing.pyqt import PyQtBackend

# La app (RAIZ.py) usa PySide6: hay que importar el MISMO
# binding, porque si no los QRectF de los items de la escena
# (PySide6) y los de aqui (PyQt5) son de distinto tipo y
# united() falla: el ajuste de vista nunca se aplicaria.
try:
    from PySide6.QtCore import Qt, QRectF, QTimer
except ImportError:
    try:
        from PyQt5.QtCore import Qt, QRectF, QTimer
    except ImportError:
        from PyQt6.QtCore import Qt, QRectF, QTimer

# Constante interna que usa ezdxf para adjuntar la entidad DXF de origen a
# cada QGraphicsItem dibujado (la usamos para saber a qué capa pertenece
# cada item sin tener que adivinar por tamaño/posición).
from ezdxf.addons.drawing.pyqt import CorrespondingDXFEntity

# =========================================================================
# CONFIGURACIÓN Y VARIABLES GLOBALES
# =========================================================================
MOSTRAR_CAPAS_NO_IMPRIMIBLES = 0
# ezdxf dibuja en coordenadas Qt (Y hacia abajo), por eso el dibujo salía al revés.
# Espejos aplicados al visor (ponlos en False si algún día quieres desactivarlos):
ESPEJO_VERTICAL = True
ESPEJO_HORIZONTAL = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

rutas_posibles = [
    os.path.join(BASE_DIR, "Bloques", "LIBRERIA.dxf"),
    os.path.join(BASE_DIR, "LIBRERIA.dxf"),
    os.path.join(os.path.dirname(BASE_DIR), "Bloques", "LIBRERIA.dxf")
]

doc_libreria = None

for ruta in rutas_posibles:
    if os.path.exists(ruta):
        try:
            doc_libreria = ezdxf.readfile(ruta)
            break
        except Exception:
            pass

# =========================================================================
# NOTA IMPORTANTE SOBRE BLOQUES DINÁMICOS
# =========================================================================
# ezdxf NO evalúa estados de visibilidad de bloques dinámicos de AutoCAD
# (esa lógica vive en el grafo de evaluación propietario de AutoCAD).
# Por eso, para cada "estado de visibilidad" que existía en tus bloques
# dinámicos, debes guardar en LIBRERIA.dxf una copia ESTÁTICA (explotada
# y re-bloqueada) con un nombre propio. Las tablas de abajo son las que
# traducen "qué estado necesito" -> "qué nombre de bloque insertar".
#
# AJUSTA los nombres del lado derecho para que coincidan EXACTAMENTE con
# los nombres de bloque que uses en tu LIBRERIA.dxf.
# =========================================================================

# ---- POSTES: familia (según Material/clasificación) ----
FAMILIA_POSTE_BASE = {
    "concreto": "P.C",
    "madera": "P.M.",
    "metal": "P.METAL",
    "fibra": "P.F",
}

# ---- POSTES: sufijo de bloque estático según el estado "Visibility1" (altura) ----
# Estado "30 pies" / "35 pies" / "40 pies" / "45 pies" -> sufijo del nombre de bloque
SUFIJO_ALTURA_POSTE = {
    "30 pies": "30",
    "35 pies": "35",
    "40 pies": "40",
    "44 pies": "45",   # no existe estado propio de 44'; se usa el bloque de 45' -> AJUSTA si tienes uno específico
    "45 pies": "45",
}

def nombre_bloque_poste(familia, altura_texto):
    """
    familia: 'concreto' | 'madera' | 'metal' | 'fibra'
    altura_texto: '30 pies' | '35 pies' | '40 pies' | '44 pies' | '45 pies'
    Devuelve el nombre real de bloque estático, ej: 'P.C-35'.
    Si la librería no tiene la variante de altura, usa el bloque
    base (ej: 'P.C'), que es lo que dibujan la mayoría de las librerías.
    """
    base = FAMILIA_POSTE_BASE.get(familia, "")
    if not base:
        return ""
    sufijo = SUFIJO_ALTURA_POSTE.get(altura_texto, "")
    if sufijo and doc_libreria and f"{base}-{sufijo}" in doc_libreria.blocks:
        return f"{base}-{sufijo}"
    return base

# ---- ANCLAS: nombre de bloque estático según "TIPO DE RETENIDA" ----
MAPEO_ANCLAS = {
    "ANCLAS DOBLES": "ANCLAS",
    "ANCLAS DOBLES VERTICALES": "ANCLAS",
    "ANCLA SIMPLE": "ANCLAS",
    "ANCLA SIMPLE VERTICAL": "ANCLAS",
}

# ---- Resto de materiales -> nombre de bloque en LIBRERIA.dxf ----
MAPEO_MATERIALES = {
    "CIMENTACION": "CIMENTACION",
    "ESTRUCTURA": "ESTRUCTURAS",
    "ESTRUCTURA REMATE DOBLE": "ESTRUCTURAS",
    "ESTRUCTURA REMATE SIMPLE": "ESTRUCTURAS",
    "TRANSFORMADORES": "TRAFOS",
    "CORTACIRCUITO": "CORTA_CIRCUITO",
    "PARARRAYOS": "PARARRAYOS",
    "AMORTIGUADORES": "AMORTIGUADORES",
    "TIERRAS": "TIERRAS",
    "ANILLO DE TIERRA": "ANILLO_TIERRA"
}

MAPA_TENSION = {"MEDIA TENSION": "MT", "MEDIA + BAJA TENSION": "MT+BT", "BAJA TENSION": "BT"}
MAPA_FASE = {1: "1F", 2: "2F", 3: "3F", "1F": "1F", "2F": "2F", "3F": "3F"}

# Radio del array polar para bloques genericos. Los bloques reales
# de la libreria se extienden hacia arriba desde el origen del
# poste (~10-16 unidades), por eso el generico va al mismo radio.
RADIO_POLAR_FALLBACK = 10.0

# Bloques que otros bloques traen anidados. El importador de
# ezdxf no los incluye automaticamente y el render falla
# (DXFStructureError) si no existen en el documento destino.
BLOQUES_ANIDADOS = {
    "TRAFOS": ["1143"],
}


def _bloque_tiene_geometria(doc, nombre_bloque):
    """True si el bloque existe y tiene geometria visible (no solo
    atributos). Sirve para detectar bloques 'vacios' en la libreria,
    como ESTRUCTURAS, que existen pero no dibujan nada."""
    if not doc or not nombre_bloque or nombre_bloque not in doc.blocks:
        return False
    tipos_geometricos = (
        "LINE", "LWPOLYLINE", "POLYLINE", "CIRCLE", "ARC", "ELLIPSE",
        "SPLINE", "HATCH", "SOLID", "TRACE", "INSERT", "MESH", "IMAGE",
    )
    for entidad in doc.blocks.get(nombre_bloque):
        if entidad.dxftype() in tipos_geometricos:
            return True
    return False


def _importar_bloque(doc_libreria, doc_destino, nombre_bloque):
    """Importa un bloque (y sus bloques anidados) al documento
    destino y vacia el texto de los MTEXT que traen.

    El MTEXT de estos bloques tiene columnas y ezdxf no puede
    medirlo si la fuente no se resuelve en la maquina (crashea
    el dibujo). Como son solo etiquetas estaticas -los datos
    dinamicos van en los atributos- se dejan en blanco y la
    geometria del bloque se dibuja igual.
    """
    if nombre_bloque in doc_destino.blocks:
        return
    para_importar = [nombre_bloque]
    for anidado in BLOQUES_ANIDADOS.get(nombre_bloque, []):
        if anidado in doc_libreria.blocks and anidado not in doc_destino.blocks:
            para_importar.append(anidado)
    importer.Importer(doc_libreria, doc_destino).import_blocks(para_importar)
    for nombre in para_importar:
        bloque = doc_destino.blocks.get(nombre)
        if bloque is None:
            continue
        for entidad in list(bloque):
            if entidad.dxftype() == "MTEXT":
                entidad.dxf.text = ""


# =========================================================================
# BLOQUES GENÉRICOS (fallback cuando LIBRERIA.dxf no tiene el bloque)
# =========================================================================
_CONTADOR_FALLBACK = 0
_FORMAS_FALLBACK = ("circulo", "cuadrado", "triangulo", "rombo")
_COLORES_FALLBACK = (1, 2, 3, 4, 5, 6)  # rojo, amarillo, verde, cian, azul, magenta


def _siguiente_indice_fallback():
    """Devuelve el índice del próximo bloque genérico y avanza el contador."""
    global _CONTADOR_FALLBACK
    indice = _CONTADOR_FALLBACK
    _CONTADOR_FALLBACK += 1
    return indice


def _reset_fallback():
    """Reinicia el contador al empezar un dibujo, para que la alternancia
    de forma/color sea estable (no cambia entre redibujos)."""
    global _CONTADOR_FALLBACK
    _CONTADOR_FALLBACK = 0


def _rotar_punto(px, py, cx, cy, angulo_grados):
    """Rota un punto alrededor de (cx, cy) los grados indicados."""
    if not angulo_grados:
        return px, py
    rad = math.radians(angulo_grados)
    dx, dy = px - cx, py - cy
    return (cx + dx * math.cos(rad) - dy * math.sin(rad),
            cy + dx * math.sin(rad) + dy * math.cos(rad))


def dibujar_fallback(msp, cx, cy, rotacion, nombre_bloque, fallback_type, txt_dyn=""):
    """Bloque genérico que se dibuja cuando el bloque solicitado no existe
    (o está vacío) en LIBRERIA.dxf. Cada elemento alterna forma y color
    para distinguirse y se coloca en un array polar alrededor del poste
    (misma dirección que los bloques reales, que se extienden hacia
    arriba desde el origen); las estructuras siempre son un círculo con
    el nombre de la estructura encerrado dentro."""
    indice = _siguiente_indice_fallback()

    if fallback_type == "ESTRUCTURA":
        forma, color = "circulo", 3
        texto = str(txt_dyn or nombre_bloque or "ESTRUCTURA").strip()
    else:
        forma = _FORMAS_FALLBACK[indice % len(_FORMAS_FALLBACK)]
        color = _COLORES_FALLBACK[indice % len(_COLORES_FALLBACK)]
        texto = str(txt_dyn).strip()

    atributos = {'color': color}

    # El poste y la cimentación van en el centro; el resto en array polar.
    if fallback_type in ("POSTE", "CIMENTACION"):
        px, py = cx, cy
    else:
        # Los bloques reales se extienden hacia arriba (+Y) desde el
        # origen y luego se rotan: ángulo 0 = arriba del poste,
        # 90 = a la derecha, etc. El genérico sigue la misma dirección.
        rad = math.radians(rotacion)
        px = cx - RADIO_POLAR_FALLBACK * math.sin(rad)
        py = cy + RADIO_POLAR_FALLBACK * math.cos(rad)

    if forma == "circulo":
        msp.add_circle((px, py), radius=2, dxfattribs=atributos)
    elif forma == "cuadrado":
        puntos = [(px - 1.6, py - 1.6), (px + 1.6, py - 1.6),
                  (px + 1.6, py + 1.6), (px - 1.6, py + 1.6),
                  (px - 1.6, py - 1.6)]
        msp.add_lwpolyline([_rotar_punto(ax, ay, px, py, rotacion) for ax, ay in puntos], dxfattribs=atributos)
    elif forma == "triangulo":
        puntos = [(px, py + 2), (px - 1.8, py - 1.5), (px + 1.8, py - 1.5), (px, py + 2)]
        msp.add_lwpolyline([_rotar_punto(ax, ay, px, py, rotacion) for ax, ay in puntos], dxfattribs=atributos)
    else:  # rombo
        puntos = [(px, py + 2), (px + 1.6, py), (px, py - 2), (px - 1.6, py), (px, py + 2)]
        msp.add_lwpolyline([_rotar_punto(ax, ay, px, py, rotacion) for ax, ay in puntos], dxfattribs=atributos)

    if texto:
        # El nombre de la estructura va dentro del círculo; lo demás, debajo.
        tx, ty = (px, py) if (forma == "circulo" and fallback_type == "ESTRUCTURA") else (px, py - 3)
        t = msp.add_text(texto, dxfattribs={'height': 1.5, 'color': color})
        t.set_placement((tx, ty), align=TextEntityAlignment.MIDDLE_CENTER)


def insertar_bloque_dxf(doc_destino, msp, nombre_bloque, cx, cy, rotacion, atributos=None, fallback_type="MATERIAL", faltantes=None):
    bloque_exitoso = False
    if doc_libreria and nombre_bloque and _bloque_tiene_geometria(doc_libreria, nombre_bloque):
        try:
            _importar_bloque(doc_libreria, doc_destino, nombre_bloque)

            blk_ref = msp.add_blockref(nombre_bloque, insert=(cx, cy), dxfattribs={'rotation': rotacion})
            bloque_exitoso = True

            if atributos:
                blk_ref.add_auto_attribs(atributos)
            # Atributos siempre visibles y en su capa, para poder revisarlos en el
            # dibujo tanto en modo Nuevo como en modo Editar.
            if blk_ref.attribs:
                for attrib in blk_ref.attribs:
                    attrib.dxf.layer = "00-ATRIBUTOS INVISIBLES"
                    attrib.dxf.invisible = 0
        except Exception:
            bloque_exitoso = False

    if not bloque_exitoso:
        txt_dyn = atributos.get("CANTIDAD", "") if atributos else ""
        if not txt_dyn and atributos and "ESTRUCTURA" in atributos:
            txt_dyn = atributos.get("ESTRUCTURA", "")
        dibujar_fallback(msp, cx, cy, rotacion, nombre_bloque, fallback_type, txt_dyn)
        # Se registra para poder avisarle al usuario en la UI cuáles bloques
        # NO existen (aún) en LIBRERIA.dxf con ese nombre exacto.
        if faltantes is not None:
            etiqueta = nombre_bloque if nombre_bloque else f"(sin nombre / {fallback_type})"
            if etiqueta not in faltantes:
                faltantes.append(etiqueta)
    return bloque_exitoso


def extraer_datos_poste(texto_poste_largo):
    """
    Segmenta el texto de 'Tipo Poste:' en (familia, altura_texto, dan_texto).
    Ej: '35 pies 300daN Poste de Concreto' -> ('concreto', '35 pies', '300')
        '35 pies Clase 5 Poste de Madera'  -> ('madera', '35 pies', '5')
    """
    familia = ""
    altura_texto = ""
    dan_texto = ""

    if not texto_poste_largo:
        return familia, altura_texto, dan_texto

    t_lower = texto_poste_largo.lower()
    if "concreto" in t_lower:
        familia = "concreto"
    elif "madera" in t_lower:
        familia = "madera"
    elif "metal" in t_lower:
        familia = "metal"
    elif "fibra" in t_lower:
        familia = "fibra"

    for h in ["30 pies", "35 pies", "40 pies", "44 pies", "45 pies"]:
        if h in texto_poste_largo:
            altura_texto = h
            break

    if familia == "madera":
        match = re.search(r'Clase\s+(\d+)', texto_poste_largo, re.IGNORECASE)
        if match:
            dan_texto = match.group(1)
    else:
        match = re.search(r'(\d+)\s*daN', texto_poste_largo, re.IGNORECASE)
        if match:
            dan_texto = match.group(1)

    return familia, altura_texto, dan_texto


def _rect_geometria_visible(escena, capa_excluir="00-ATRIBUTOS INVISIBLES"):
    """
    Devuelve el QRectF que envuelve la geometria 'fisica' del dibujo (postes,
    estructuras, anclas, materiales), sin contar el texto de los atributos.
    Los atributos (capa 00-ATRIBUTOS INVISIBLES) se siguen dibujando y se ven
    igual, pero no participan en el calculo del encuadre/zoom: como estan
    posicionados dentro de cada bloque en un lugar que no siempre coincide
    con el icono visual, dejarlos fuera evita que el encuadre se abra de mas
    y el dibujo se vea chico dentro de un lienzo enorme.
    Si por algun motivo no queda geometria fuera de esa capa (dibujo vacio,
    solo error, etc.) se usa el rect de TODO lo dibujado como respaldo.
    """
    rect = QRectF()
    hubo_geometria = False
    for item in escena.items():
        entidad = item.data(CorrespondingDXFEntity)
        capa = getattr(getattr(entidad, "dxf", None), "layer", None)
        if capa == capa_excluir:
            continue
        rect = rect.united(item.sceneBoundingRect())
        hubo_geometria = True

    if not hubo_geometria:
        rect = escena.itemsBoundingRect()

    return rect


def _ajustar_vista(escena, vista_grafica):
    """Calcula el encuadre de la geometria visible y ajusta la camara,
    guardando el rect usado para que la vista pueda reajustarse sola si
    el panel cambia de tamano despues (ver VistaCAD.resizeEvent en RAIZ.py)."""
    escena.update()
    vista_grafica.update()

    rect_real = _rect_geometria_visible(escena)

    if rect_real.width() <= 0 or rect_real.height() <= 0:
        rect_real = QRectF(-50, -50, 100, 100)

    # Margen proporcional al tamano del dibujo, con un minimo fijo para
    # que las estructuras muy pequenas (ej. solo un poste) no queden
    # pegadas al borde del visor.
    margen_x = max(rect_real.width() * 0.15, 5)
    margen_y = max(rect_real.height() * 0.15, 5)
    rect_real.adjust(-margen_x, -margen_y, margen_x, margen_y)

    # El ajuste final acerca la vista 1.5 veces al contenido para que
    # el dibujo no se vea pequeno (se recorta algo del excedente; se
    # puede alejar con la rueda del raton o el boton "Ajustar").
    factor_acercamiento = 1.5
    ancho = rect_real.width() / factor_acercamiento
    alto = rect_real.height() / factor_acercamiento
    centro = rect_real.center()
    rect_real = QRectF(centro.x() - ancho / 2.0, centro.y() - alto / 2.0, ancho, alto)

    escena.setSceneRect(rect_real)
    # Se guarda para que VistaCAD pueda reencuadrar automaticamente al
    # redimensionar la ventana, sin tener que volver a dibujar todo.
    vista_grafica._dibujo_rect_actual = QRectF(rect_real)
    # El espejo se aplica ANTES del fitInView (que conserva el signo) y una sola vez
    vista_grafica.resetTransform()
    vista_grafica.scale(-1 if ESPEJO_HORIZONTAL else 1, -1 if ESPEJO_VERTICAL else 1)
    vista_grafica.fitInView(rect_real, Qt.KeepAspectRatio)

    # Un segundo ajuste diferido: si esta es la primera vez que se muestra
    # el panel, Qt puede no haber terminado de calcular el tamano real del
    # viewport en el momento de este fitInView. Este segundo llamado, ya
    # con el layout resuelto, corrige ese caso sin que se note parpadeo.
    def _reajustar_diferido():
        if vista_grafica._dibujo_rect_actual is not None:
            vista_grafica.fitInView(vista_grafica._dibujo_rect_actual, Qt.KeepAspectRatio)

    QTimer.singleShot(0, _reajustar_diferido)


def obtener_angulos_anclas(cantidad):
    """
    Las anclas van primero, separadas ~40 grados entre si:
    1 ancla -> [0]; 2 -> [0, 40]; 3 -> [0, 40, 80]; ...
    Aplica sobre el total combinado de anclas (mismo tipo o tipos distintos).
    """
    if cantidad <= 0:
        return []
    return [i * 40.0 for i in range(cantidad)]


def obtener_angulos_rotantes(cantidad, angulos_anclas):
    """
    Los elementos extras (transformador, pararrayos, tierras, etc.)
    se reparten de forma uniforme en el espacio de angulo que queda
    libre despues de las anclas.
    """
    if cantidad <= 0:
        return []
    if not angulos_anclas:
        # Sin anclas: repartidos en todo el círculo.
        return [i * (360.0 / cantidad) for i in range(cantidad)]
    inicio = max(angulos_anclas) + 20.0
    if cantidad == 1:
        return [inicio]
    fin = 340.0  # hueco antes de volver a la primera ancla
    espacio = fin - inicio
    if espacio < 15.0:
        # No queda espacio: se acomodan con paso fijo (puede superar 360).
        return [inicio + i * 45.0 for i in range(cantidad)]
    return [inicio + espacio * i / (cantidad - 1) for i in range(cantidad)]


def ensamblar_y_dibujar(escena, vista_grafica, nombre_poste, altura_poste, dan_poste, tiene_cimentacion,
                          textos_info, lista_anclas, elementos_rotantes, tipo_linea_texto, num_poste,
                          fases, estructura_tipo, modo_dia=False):
    """
    nombre_poste: puede llegar como FAMILIA ('concreto'|'madera'|'metal'|'fibra') o ya
                  como nombre de bloque base ('P.C', 'P.M.', etc.) - se resuelve igual.
    altura_poste: texto tal cual '35 pies', etc. (para elegir el bloque estatico correcto)
    lista_anclas: lista de strings con el MATERIAL de cada ancla, ej:
                  ["ANCLAS DOBLES", "ANCLAS DOBLES", "ANCLA SIMPLE"]
    """
    escena.clear()
    vista_grafica.resetTransform()
    _reset_fallback()

    doc_dxf = ezdxf.new('R2010')
    if doc_libreria:
        for layer in doc_libreria.layers:
            if layer.dxf.name not in doc_dxf.layers:
                nueva_capa = doc_dxf.layers.new(
                    name=layer.dxf.name,
                    dxfattribs={'color': layer.dxf.color, 'linetype': layer.dxf.linetype}
                )
                if hasattr(layer.dxf, 'plot') and not layer.dxf.plot and MOSTRAR_CAPAS_NO_IMPRIMIBLES == 0:
                    nueva_capa.off()

    if "00-ATRIBUTOS INVISIBLES" not in doc_dxf.layers:
        doc_dxf.layers.new(name="00-ATRIBUTOS INVISIBLES", dxfattribs={'color': 30})
    capa_attrs = doc_dxf.layers.get("00-ATRIBUTOS INVISIBLES")
    capa_attrs.on()
    capa_attrs.thaw()

    if "MI_ARIAL" not in doc_dxf.styles:
        doc_dxf.styles.new("MI_ARIAL", dxfattribs={"font": "arial.ttf"})

    msp = doc_dxf.modelspace()

    if not doc_libreria:
        t = msp.add_text("ERROR: LIBRERIA NO ENCONTRADA", dxfattribs={'height': 3, 'color': 1})
        t.set_placement((0, 5), align=TextEntityAlignment.MIDDLE_CENTER)

    tipo_tension = MAPA_TENSION.get(str(tipo_linea_texto).strip().upper(), "")
    fase_texto = MAPA_FASE.get(fases, fases if fases in ("1F", "2F", "3F") else "")

    cx, cy = 0, 0
    faltantes = []  # nombres de bloque que se pidieron pero no existen en LIBRERIA.dxf

    # ---------------------------------------------------------------
    # 1. POSTE (bloque estatico resuelto por familia + altura)
    # ---------------------------------------------------------------
    if nombre_poste:
        if nombre_poste in FAMILIA_POSTE_BASE:
            bloque_poste = nombre_bloque_poste(nombre_poste, altura_poste)
        else:
            # Compatibilidad: si llega ya como 'P.C', 'P.M.', etc., resolvemos el sufijo igual
            familia_inversa = {v: k for k, v in FAMILIA_POSTE_BASE.items()}
            familia = familia_inversa.get(nombre_poste)
            bloque_poste = nombre_bloque_poste(familia, altura_poste) if familia else nombre_poste

        insertar_bloque_dxf(
            doc_dxf, msp, bloque_poste, cx, cy, 0,
            {
                "NO_POSTE": f"No. Poste: {num_poste}" if num_poste else "",
                "TIPO_LINEA": tipo_tension,
                "daN": str(dan_poste) if dan_poste else "",
            },
            "POSTE", faltantes
        )

    # ---------------------------------------------------------------
    # 2. ESTRUCTURA
    # ---------------------------------------------------------------
    if estructura_tipo:
        insertar_bloque_dxf(
            doc_dxf, msp, "ESTRUCTURAS", cx, cy, 0,
            {
                "TIPO_TENSION": tipo_tension,
                "FASE": fase_texto,
                "ESTRUCTURA": estructura_tipo,
            },
            "ESTRUCTURA", faltantes
        )

    # ---------------------------------------------------------------
    # 3. CIMENTACION
    # ---------------------------------------------------------------
    if tiene_cimentacion:
        insertar_bloque_dxf(doc_dxf, msp, MAPEO_MATERIALES.get("CIMENTACION", "CIMENTACION"), cx, cy, 0, None, "CIMENTACION", faltantes)

    # ---------------------------------------------------------------
    # 4. ANCLAS (bloque segun tipo + angulos segun cantidad total)
    # ---------------------------------------------------------------
    angulos_anclas = obtener_angulos_anclas(len(lista_anclas))
    for angulo, tipo_ancla in zip(angulos_anclas, lista_anclas):
        bloque_ancla = MAPEO_ANCLAS.get(tipo_ancla, "ANCLAS")
        insertar_bloque_dxf(
            doc_dxf, msp, bloque_ancla, cx, cy, angulo,
            {"TIPO_TENSION": tipo_tension},
            "ANCLA", faltantes
        )

    # ---------------------------------------------------------------
    # 5. ELEMENTOS ROTANTES (transformador, pararrayos, tierras, etc.)
    # ---------------------------------------------------------------
    angulos_rotantes = obtener_angulos_rotantes(len(elementos_rotantes), angulos_anclas)
    for angulo_actual, elem in zip(angulos_rotantes, elementos_rotantes):
        insertar_bloque_dxf(
            doc_dxf, msp, MAPEO_MATERIALES.get(elem["id"], elem["id"]), cx, cy, angulo_actual,
            {"CANTIDAD": f"X{elem['texto']}" if str(elem['texto']).isdigit() else str(elem['texto'])},
            "MATERIAL", faltantes
        )

    # ---------------------------------------------------------------
    # 6. TEXTOS INFORMATIVOS (fases, notas, etc.)
    # ---------------------------------------------------------------
    if textos_info:
        y_texto = -15
        for txt in list(dict.fromkeys(textos_info)):
            if txt and txt.strip():
                t_info = msp.add_text(txt, dxfattribs={'height': 2.5, 'color': 3, 'style': 'MI_ARIAL', 'rotation': 0})
                t_info.set_placement((cx, cy + y_texto), align=TextEntityAlignment.MIDDLE_CENTER)
                y_texto -= 4

    # -----------------------------------------------------------------
    # RENDERIZADO (Fondo manual para evitar bugs de ezdxf)
    # -----------------------------------------------------------------
    ctx = RenderContext(doc_dxf)
    out = PyQtBackend(escena)

    cfg = config.Configuration(
        background_policy=config.BackgroundPolicy.CUSTOM,
        custom_bg_color="#00000000",
        color_policy=config.ColorPolicy.BLACK if modo_dia else config.ColorPolicy.COLOR
    )

    Frontend(ctx, out, config=cfg).draw_layout(msp, finalize=True)
    vista_grafica.setStyleSheet(f"background-color: {'#FFFFFF' if modo_dia else '#212830'}; border: none;")

    # -----------------------------------------------------------------
    # CAMARA: auto-ajuste al contenido real (ver _ajustar_vista arriba)
    # -----------------------------------------------------------------
    _ajustar_vista(escena, vista_grafica)

    return faltantes