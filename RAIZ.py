import sys
import os
import pandas as pd
import numpy as np
import ctypes
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QLineEdit,
    QComboBox, QRadioButton, QListWidget, QMessageBox,
    QFileDialog, QHBoxLayout, QVBoxLayout, QButtonGroup,
    QDialog, QTableWidget, QTableWidgetItem, QHeaderView, QSplashScreen,
    QSplitter, QGraphicsView, QGraphicsScene, QCheckBox
)
from PySide6.QtGui import QIcon, QKeySequence, QShortcut, QPixmap
from PySide6.QtCore import Qt, QTimer
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
import shutil
import time
import graphics

try:
    import xlwings as xw
except ImportError:
    print("Advertencia: xlwings no está instalado. Los filtros automáticos no funcionarán.")
    xw = None

from license_manager import check_license

if not check_license():
    print("Licencia inválida o expirada.")
    sys.exit()

print("Programa iniciado correctamente")

COLUMNAS_LINEA = [
    "MEDIA TENSION",
    "MEDIA + BAJA TENSION",
    "MEDIA + BAJA TENSION 2/0",
    "MEDIA + BAJA TENSION 1/0",
    "BAJA TENSION 2/0",
    "BAJA TENSION 1/0"
]

MATERIALES_COLUMNAS = [
    "ESTRUCTURA",
    "ESTRUCTURA REMATE DOBLE",
    "ESTRUCTURA REMATE SIMPLE",
    "ANCLAS DOBLES",
    "ANCLAS DOBLES VERTICALES",
    "ANCLA SIMPLE",
    "ANCLA SIMPLE VERTICAL",
    "TIERRAS",
    "PARARRAYOS",
    "CORTACIRCUITO",
    "ANILLO DE TIERRA",
    "TRANSFORMADORES",
    "AMORTIGUADORES",
    "RETENIDA CON APOYO SIMPLE",
    "RETENIDA CON APOYO DOBLE",
    "CONECTORES BT",
    "SUPLEMENTOS POR EXCAVACION",
    "CIMENTACION"
]

MATERIALES_TEXTO = ["ESTRUCTURA", "TRANSFORMADORES", "CIMENTACION"]

# Materiales y tipos creados por el usuario en esta sesión (para poder borrarlos)
MATERIALES_CREADOS = set()
TIPOS_ESTRUCTURA_CREADOS = set()
TIPOS_TRANSFORMADOR_CREADOS = set()

TIPOS_ESTRUCTURA = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "XI", "VI CV"]

TIPOS_TRANSFORMADOR = [
    "Transformador autoprotegido bajas pérdidas 25 kVA. (19.9 kV)",
    "Transformador autoprotegido bajas pérdidas 10 kVA. (19.9 kV)",
    "Transformador autoprotegido bajas pérdidas 50 kVA. (19.9 kV)"
]

TIPOS_POSTE_CONCRETO = [
    "35 pies 300 daN Poste de concreto",
    "35 pies 500 daN Poste de concreto",
    "35 pies 800 daN Poste de concreto",
    "40 pies 300 daN Poste de concreto",
    "40 pies 500 daN Poste de concreto",
    "40 pies 800 daN Poste de concreto",
    "44 pies 800 daN Poste de concreto",
    "30 pies 300 daN Poste de concreto",
    "30 pies 500 daN Poste de concreto"
]

TIPOS_POSTE_MADERA = [
    "40 pies Clase 5 Poste de madera",
    "40 pies Clase 3 Poste de madera",
    "35 pies Clase 5 Poste de madera",
    "35 pies Clase 3 Poste de madera",
    "30 pies Clase 5 Poste de madera"
]

TIPOS_POSTE_METAL = [
    "44 pies 300 daN Poste de metal",
    "44 pies 500 daN Poste de metal",
    "44 pies 800 daN Poste de metal",
    "40 pies 300 daN Poste de metal",
    "40 pies 500 daN Poste de metal",
    "40 pies 800 daN Poste de metal",
    "35 pies 300 daN Poste de metal",
    "35 pies 500 daN Poste de metal",
    "35 pies 800 daN Poste de metal",
    "30 pies 300 daN Poste de metal",
    "30 pies 500 daN Poste de metal"
]

TIPOS_POSTE_FIBRA = [
    "44 pies 300 daN Poste de Fibra",
    "44 pies 500 daN Poste de Fibra",
    "44 pies 800 daN Poste de Fibra",
    "40 pies 300 daN Poste de Fibra",
    "40 pies 500 daN Poste de Fibra",
    "40 pies 800 daN Poste de Fibra",
    "35 pies 300 daN Poste de Fibra",
    "35 pies 500 daN Poste de Fibra",
    "35 pies 800 daN Poste de Fibra",
    "30 pies 300 daN Poste de Fibra",
    "30 pies 500 daN Poste de Fibra",
    "30 pies 800 daN Poste de Fibra"
]

TIPOS_POSTE = TIPOS_POSTE_CONCRETO + TIPOS_POSTE_MADERA + TIPOS_POSTE_METAL + TIPOS_POSTE_FIBRA
TABLAS_FILTRAR = ["MT", "MT_BT", "BT"]


def _poste_sort_key(nombre):
    import re
    m = re.search(r'(\d+)\s*pies', nombre)
    altura = int(m.group(1)) if m else 0
    if "Clase" in nombre:
        c = re.search(r'Clase\s+(\d+)', nombre)
        clase = int(c.group(1)) if c else 0
        return (-altura, clase)
    d = re.search(r'(\d+)\s*daN', nombre)
    danaje = int(d.group(1)) if d else 0
    return (-altura, -danaje)


def ordenar_tipos_poste(lista):
    lista.sort(key=_poste_sort_key)


ordenar_tipos_poste(TIPOS_POSTE_CONCRETO)
ordenar_tipos_poste(TIPOS_POSTE_MADERA)
ordenar_tipos_poste(TIPOS_POSTE_METAL)
ordenar_tipos_poste(TIPOS_POSTE_FIBRA)
TIPOS_POSTE = TIPOS_POSTE_CONCRETO + TIPOS_POSTE_MADERA + TIPOS_POSTE_METAL + TIPOS_POSTE_FIBRA


# ==============================
# VISOR CAD
# ==============================
class VistaCAD(QGraphicsView):
    def __init__(self, escena):
        super().__init__(escena)
        self._dibujo_rect_actual = None
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorUnderMouse)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._dibujo_rect_actual is not None:
            self.fitInView(self._dibujo_rect_actual, Qt.KeepAspectRatio)

    def wheelEvent(self, event):
        factor = 1.15 if event.angleDelta().y() > 0 else 1 / 1.15
        self.scale(factor, factor)

    def zoom_in(self):
        self.scale(1.25, 1.25)

    def zoom_out(self):
        self.scale(1 / 1.25, 1 / 1.25)

    def ajustar_a_contenido(self):
        if self._dibujo_rect_actual is not None:
            self.fitInView(self._dibujo_rect_actual, Qt.KeepAspectRatio)


# ==============================
# NUEVAS VENTANAS DE CREACIÓN
# ==============================
class DialogoGestorPostes(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_form = parent
        self.setWindowTitle("Gestor de Postes")
        layout = QVBoxLayout(self)

        self.material_actual = "Concreto"
        self.btn_concreto = QPushButton("Concreto")
        self.btn_madera = QPushButton("Madera")
        self.btn_metal = QPushButton("Metal")
        self.btn_fibra = QPushButton("Fibra")
        for b in (self.btn_concreto, self.btn_madera, self.btn_metal, self.btn_fibra):
            b.setCheckable(True)
        self.btn_concreto.setChecked(True)
        fila_mat = QHBoxLayout()
        for b in (self.btn_concreto, self.btn_madera, self.btn_metal, self.btn_fibra):
            fila_mat.addWidget(b)
        layout.addLayout(fila_mat)

        fila1 = QHBoxLayout()
        fila1.addWidget(QLabel("Tipo Poste:"))
        self.cmb_tipo = QComboBox()
        self.cmb_tipo.setMinimumWidth(300)
        fila1.addWidget(self.cmb_tipo, 1)
        layout.addLayout(fila1)

        self.lbl_alt = QLabel("Altura:")
        self.txt_alt = QLineEdit()
        self.txt_alt.setPlaceholderText("Ej: 40")
        from PySide6.QtGui import QIntValidator
        self.txt_alt.setValidator(QIntValidator(1, 200, self))
        fila2 = QHBoxLayout()
        fila2.addWidget(self.lbl_alt)
        fila2.addWidget(self.txt_alt)
        fila2.addWidget(QLabel("pies"))
        layout.addLayout(fila2)

        self.lbl_danaje = QLabel("Danaje:")
        self.txt_danaje = QLineEdit()
        self.txt_danaje.setPlaceholderText("Ej: 800")
        from PySide6.QtGui import QIntValidator
        self.txt_danaje.setValidator(QIntValidator(0, 10000, self))
        fila3 = QHBoxLayout()
        fila3.addWidget(self.lbl_danaje)
        fila3.addWidget(self.txt_danaje)
        self.lbl_unidad = QLabel("daN")
        fila3.addWidget(self.lbl_unidad)
        layout.addLayout(fila3)

        fila_btn = QHBoxLayout()
        self.btn_agregar = QPushButton("Agregar")
        self.btn_borrar = QPushButton("Borrar")
        fila_btn.addWidget(self.btn_agregar)
        fila_btn.addWidget(self.btn_borrar)
        layout.addLayout(fila_btn)

        self.btn_regresar = QPushButton("Regresar")
        layout.addWidget(self.btn_regresar)

        self.btn_regresar.clicked.connect(self.reject)
        self.btn_agregar.clicked.connect(self.agregar)
        self.btn_borrar.clicked.connect(self.borrar)
        self.cmb_tipo.currentTextChanged.connect(self.autocompletar)
        self.txt_alt.textChanged.connect(self.actualizar_estado)
        self.txt_danaje.textChanged.connect(self.actualizar_estado)
        for b, m in ((self.btn_concreto, "Concreto"), (self.btn_madera, "Madera"), (self.btn_metal, "Metal"), (self.btn_fibra, "Fibra")):
            b.clicked.connect(lambda checked=False, mm=m: self.cambiar_material(mm))

        self.setStyleSheet(parent.styleSheet() if parent is not None else "")
        estilo_btn = (
            "QPushButton { background-color: #3498db; color: white; border: none; padding: 4px 14px; }"
            "QPushButton:hover { background-color: #2980b9; }"
            "QPushButton:disabled { background-color: #95a5a6; }"
        )
        estilo_borrar = (
            "QPushButton { background-color: #e74c3c; color: white; border: none; padding: 4px 14px; }"
            "QPushButton:hover { background-color: #c0392b; }"
            "QPushButton:disabled { background-color: #95a5a6; }"
        )
        estilo_clas = "QPushButton { background-color: #95a5a6; color: white; border: none; padding: 3px 8px; font-size: 7pt; } QPushButton:hover { background-color: #7f8c8d; } QPushButton:checked { background-color: #3498db; }"
        for b in (self.btn_concreto, self.btn_madera, self.btn_metal, self.btn_fibra):
            b.setStyleSheet(estilo_clas)
        self.btn_agregar.setStyleSheet(estilo_btn)
        self.btn_regresar.setStyleSheet(estilo_btn)
        self.btn_borrar.setStyleSheet(estilo_borrar)

        self.cargar_tipos()
        self.actualizar_estado()

    def cambiar_material(self, m):
        self.material_actual = m
        self.btn_concreto.setChecked(m == "Concreto")
        self.btn_madera.setChecked(m == "Madera")
        self.btn_metal.setChecked(m == "Metal")
        self.btn_fibra.setChecked(m == "Fibra")
        if m == "Madera":
            self.lbl_danaje.setText("Clase:")
            self.lbl_unidad.setText("")
            self.txt_danaje.setPlaceholderText("Ej: 3")
        else:
            self.lbl_danaje.setText("Danaje:")
            self.lbl_unidad.setText("daN")
            self.txt_danaje.setPlaceholderText("Ej: 800")
        self.cargar_tipos()
        self.actualizar_estado()

    def _tipos_actuales(self):
        if self.material_actual == "Concreto":
            return TIPOS_POSTE_CONCRETO
        if self.material_actual == "Madera":
            return TIPOS_POSTE_MADERA
        if self.material_actual == "Metal":
            return TIPOS_POSTE_METAL
        return TIPOS_POSTE_FIBRA

    def cargar_tipos(self):
        self.cmb_tipo.blockSignals(True)
        self.cmb_tipo.clear()
        self.cmb_tipo.addItem("")
        self.cmb_tipo.addItems(self._tipos_actuales())
        self.cmb_tipo.blockSignals(False)

    def autocompletar(self):
        import re
        texto = self.cmb_tipo.currentText()
        if not texto:
            return
        m = re.search(r'(\d+)\s*pies', texto)
        if m:
            self.txt_alt.setText(m.group(1))
        if self.material_actual == "Madera":
            m2 = re.search(r'Clase\s+(\d+)', texto)
            if m2:
                self.txt_danaje.setText(m2.group(1))
        else:
            m2 = re.search(r'(\d+)\s*daN', texto)
            if m2:
                self.txt_danaje.setText(m2.group(1))

    def actualizar_estado(self):
        alt = self.txt_alt.text().strip()
        dan = self.txt_danaje.text().strip()
        nombre = self._armar_nombre()
        existe = nombre in TIPOS_POSTE if nombre else False
        self.btn_agregar.setEnabled(bool(alt and dan and alt.isdigit() and dan.isdigit() and not existe))
        self.btn_borrar.setEnabled(bool(self.cmb_tipo.currentText()))

    def _armar_nombre(self):
        alt = self.txt_alt.text().strip()
        dan = self.txt_danaje.text().strip()
        if not alt or not dan:
            return ""
        if self.material_actual == "Madera":
            return f"{alt} pies Clase {dan} Poste de madera"
        nombre_mat = self.material_actual.lower()
        if self.material_actual == "Fibra":
            nombre_mat = "Fibra"
        return f"{alt} pies {dan} daN Poste de {nombre_mat if self.material_actual != 'Fibra' else 'Fibra'}"

    def agregar(self):
        nombre = self._armar_nombre()
        if not nombre or nombre in TIPOS_POSTE:
            return
        lista = self._tipos_actuales()
        lista.append(nombre)
        TIPOS_POSTE.append(nombre)
        ordenar_tipos_poste(lista)
        ordenar_tipos_poste(TIPOS_POSTE)
        self.parent_form.cambiar_clasificacion_poste(self.material_actual)
        self.parent_form.cmb_tipo_poste.setCurrentText(nombre)
        self.cargar_tipos()
        self.cmb_tipo.setCurrentText(nombre)
        self.actualizar_estado()

    def borrar(self):
        tipo = self.cmb_tipo.currentText()
        if not tipo:
            return
        # verificar que no haya registros con este poste
        df = self.parent_form.df
        usado = False
        if df is not None:
            for idx, row in df.iterrows():
                if row.get("_hijo_vii", False):
                    continue
                val = str(row.get("TIPO DE POSTE", "")).strip()
                cim = str(row.get("CIMENTACION", "")).strip()
                if val == tipo or val == f"E {tipo}" or cim == f"C {tipo}" or cim == f"C E {tipo}" or cim.lower() == f"c {tipo.lower()}" or cim.lower() == f"c e {tipo.lower()}":
                    usado = True
                    break
        if usado:
            QMessageBox.warning(self, "En uso", "Hay registros con este poste o su cimentación, no se puede eliminar.")
            return
        resp = QMessageBox.question(self, "Confirmar", f"¿Eliminar \"{tipo}\"?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if resp != QMessageBox.StandardButton.Yes:
            return
        lista = self._tipos_actuales()
        if tipo in lista:
            lista.remove(tipo)
        if tipo in TIPOS_POSTE:
            TIPOS_POSTE.remove(tipo)
        self.parent_form.cambiar_clasificacion_poste(self.material_actual)
        self.cargar_tipos()
        self.txt_alt.clear()
        self.txt_danaje.clear()
        self.actualizar_estado()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_F3:
            self.reject()
            return
        super().keyPressEvent(event)


class DialogoGestorMateriales(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_form = parent
        self.setWindowTitle("Gestor de Materiales")
        layout = QVBoxLayout(self)

        fila1 = QHBoxLayout()
        fila1.addWidget(QLabel("Material:"))
        self.cmb_material_gestor = QComboBox()
        self.cmb_material_gestor.addItem("")
        for m in MATERIALES_COLUMNAS:
            if m != "CIMENTACION":
                self.cmb_material_gestor.addItem(m)
        fila1.addWidget(self.cmb_material_gestor, 1)
        layout.addLayout(fila1)

        fila2 = QHBoxLayout()
        self.lbl_tipo = QLabel("Tipo:")
        self.cmb_tipo = QComboBox()
        self.lbl_tipo.hide()
        self.cmb_tipo.hide()
        fila2.addWidget(self.lbl_tipo)
        fila2.addWidget(self.cmb_tipo, 1)
        layout.addLayout(fila2)

        layout.addWidget(QLabel("Nuevo:"))
        self.txt_nuevo = QLineEdit()
        self.txt_nuevo.setPlaceholderText("Escribe el nombre del material")
        layout.addWidget(self.txt_nuevo)

        fila_btn = QHBoxLayout()
        self.btn_agregar = QPushButton("Agregar")
        self.btn_borrar = QPushButton("Borrar")
        fila_btn.addWidget(self.btn_agregar)
        fila_btn.addWidget(self.btn_borrar)
        layout.addLayout(fila_btn)

        self.btn_regresar = QPushButton("Regresar")
        layout.addWidget(self.btn_regresar)

        self.btn_regresar.clicked.connect(self.reject)
        self.btn_agregar.clicked.connect(self.agregar)
        self.btn_borrar.clicked.connect(self.borrar)
        self.cmb_material_gestor.currentTextChanged.connect(self.actualizar_estado)
        self.cmb_tipo.currentTextChanged.connect(self.actualizar_estado)
        self.txt_nuevo.textChanged.connect(self.actualizar_estado)

        self.setStyleSheet(parent.styleSheet() if parent is not None else "")

        estilo_btn = (
            "QPushButton { background-color: #3498db; color: white; border: none; padding: 4px 14px; }"
            "QPushButton:hover { background-color: #2980b9; }"
            "QPushButton:disabled { background-color: #95a5a6; }"
        )
        estilo_borrar = (
            "QPushButton { background-color: #e74c3c; color: white; border: none; padding: 4px 14px; }"
            "QPushButton:hover { background-color: #c0392b; }"
            "QPushButton:disabled { background-color: #95a5a6; }"
        )
        self.btn_agregar.setStyleSheet(estilo_btn)
        self.btn_regresar.setStyleSheet(estilo_btn)
        self.btn_borrar.setStyleSheet(estilo_borrar)

        self.actualizar_estado()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_F2:
            self.reject()
            return
        super().keyPressEvent(event)

    def _es_creado_usuario(self, mat, tipo):
        if mat == "ESTRUCTURA":
            return tipo in TIPOS_ESTRUCTURA_CREADOS
        if mat == "TRANSFORMADORES":
            return tipo in TIPOS_TRANSFORMADOR_CREADOS
        return mat in MATERIALES_CREADOS

    def actualizar_estado(self):
        mat = self.cmb_material_gestor.currentText()
        texto = self.txt_nuevo.text().strip()

        if mat in ("ESTRUCTURA", "TRANSFORMADORES"):
            self.lbl_tipo.show()
            self.cmb_tipo.show()
            self.cmb_tipo.blockSignals(True)
            self.cmb_tipo.clear()
            if mat == "ESTRUCTURA":
                self.cmb_tipo.addItems(TIPOS_ESTRUCTURA)
                self.txt_nuevo.setPlaceholderText("Escribe un nuevo tipo de estructura")
            else:
                self.cmb_tipo.addItems(TIPOS_TRANSFORMADOR)
                self.txt_nuevo.setPlaceholderText("Escribe el nuevo voltaje del transformador")
            self.cmb_tipo.blockSignals(False)
        else:
            self.lbl_tipo.hide()
            self.cmb_tipo.hide()
            self.txt_nuevo.setPlaceholderText("Escribe el nombre del material")

        tipo_sel = self.cmb_tipo.currentText()

        # Borrar solo si el elemento fue creado por el usuario
        puede_borrar = False
        if mat in ("ESTRUCTURA", "TRANSFORMADORES"):
            if tipo_sel and self._es_creado_usuario(mat, tipo_sel):
                puede_borrar = True
        elif mat:
            if self._es_creado_usuario(mat, None):
                puede_borrar = True
        self.btn_borrar.setEnabled(puede_borrar)

        # Agregar solo si hay texto nuevo y no existe
        puede_agregar = False
        if mat == "":
            if texto and texto.upper() not in [m.upper() for m in MATERIALES_COLUMNAS]:
                puede_agregar = True
        elif mat == "ESTRUCTURA":
            if texto and texto.upper() not in [t.upper() for t in TIPOS_ESTRUCTURA]:
                puede_agregar = True
        elif mat == "TRANSFORMADORES":
            if texto and texto.isdigit():
                puede_agregar = True
        self.btn_agregar.setEnabled(puede_agregar)

    def agregar(self):
        global MATERIALES_COLUMNAS, TIPOS_ESTRUCTURA, TIPOS_TRANSFORMADOR
        mat = self.cmb_material_gestor.currentText()
        texto = self.txt_nuevo.text().strip()
        if not texto:
            return

        if mat == "":
            nombre = texto.upper()
            if nombre not in MATERIALES_COLUMNAS:
                MATERIALES_COLUMNAS.append(nombre)
                MATERIALES_CREADOS.add(nombre)
                if self.parent_form.df is not None:
                    self.parent_form.df[nombre] = 0
        elif mat == "ESTRUCTURA":
            nombre = texto.upper()
            if nombre not in TIPOS_ESTRUCTURA:
                TIPOS_ESTRUCTURA.append(nombre)
                TIPOS_ESTRUCTURA_CREADOS.add(nombre)
        elif mat == "TRANSFORMADORES":
            if texto.isdigit():
                vol = self.parent_form.cmb_potencia.currentText() if self.parent_form is not None else ""
                sufijo = "(7.6 kV)" if "7.6" in vol else "(19.9 kV)"
                nombre = f"Transformador autoprotegido bajas pérdidas {texto} kVA. {sufijo}"
                if nombre not in TIPOS_TRANSFORMADOR:
                    TIPOS_TRANSFORMADOR.append(nombre)
                    TIPOS_TRANSFORMADOR_CREADOS.add(nombre)

        self.parent_form.cambios_catalogo = True
        self.parent_form.filtrar_materiales()
        self.parent_form.cambiar_tipo_input()
        self.txt_nuevo.clear()
        self.actualizar_estado()

    def borrar(self):
        mat = self.cmb_material_gestor.currentText()
        if mat in ("ESTRUCTURA", "TRANSFORMADORES"):
            tipo = self.cmb_tipo.currentText()
            if not tipo:
                return
            if not self._es_creado_usuario(mat, tipo):
                QMessageBox.information(self, "Protegido", "Este elemento es inicial y no se puede borrar.")
                return
            en_uso = self._tipo_en_uso(mat, tipo)
            if en_uso:
                resp = QMessageBox.question(self, "Confirmar", "Hay registros con este tipo, ¿deseas eliminarlo?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if resp != QMessageBox.StandardButton.Yes:
                    return
                self._limpiar_tipo(mat, tipo)
            self._eliminar_tipo(mat, tipo)
        else:
            if not self._es_creado_usuario(mat, None):
                QMessageBox.information(self, "Protegido", "Este material es inicial y no se puede borrar.")
                return
            en_uso = self._material_en_uso(mat)
            if en_uso:
                resp = QMessageBox.question(self, "Confirmar", "Hay registros con este material, ¿deseas eliminarlo?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if resp != QMessageBox.StandardButton.Yes:
                    return
                self._limpiar_material(mat)
            self._eliminar_material(mat)

        self.parent_form.cambios_catalogo = True
        self.parent_form.filtrar_materiales()
        self.parent_form.cambiar_tipo_input()
        self.cmb_material_gestor.setCurrentIndex(0)
        self.txt_nuevo.clear()
        self.actualizar_estado()

    def _material_en_uso(self, mat):
        df = self.parent_form.df
        if df is None or mat not in df.columns:
            return False
        for idx, row in df.iterrows():
            if row.get("_hijo_vii", False):
                continue
            if self.parent_form._tiene_valor(row.get(mat, 0)):
                return True
        return False

    def _tipo_en_uso(self, mat, tipo):
        df = self.parent_form.df
        if df is None:
            return False
        col = "ESTRUCTURA" if mat == "ESTRUCTURA" else "TRANSFORMADORES"
        if col not in df.columns:
            return False
        for idx, row in df.iterrows():
            if row.get("_hijo_vii", False):
                continue
            if str(row.get(col, "")).strip() == tipo:
                return True
        return False

    def _limpiar_material(self, mat):
        df = self.parent_form.df
        if df is not None and mat in df.columns:
            df.drop(columns=[mat], inplace=True)

    def _limpiar_tipo(self, mat, tipo):
        df = self.parent_form.df
        if df is None:
            return
        col = "ESTRUCTURA" if mat == "ESTRUCTURA" else "TRANSFORMADORES"
        if col in df.columns:
            for idx, row in df.iterrows():
                if str(row.get(col, "")).strip() == tipo:
                    df.at[idx, col] = ""

    def _eliminar_material(self, mat):
        global MATERIALES_COLUMNAS
        if mat in MATERIALES_COLUMNAS:
            MATERIALES_COLUMNAS.remove(mat)
        MATERIALES_CREADOS.discard(mat)

    def _eliminar_tipo(self, mat, tipo):
        global TIPOS_ESTRUCTURA, TIPOS_TRANSFORMADOR
        if mat == "ESTRUCTURA":
            if tipo in TIPOS_ESTRUCTURA:
                TIPOS_ESTRUCTURA.remove(tipo)
            TIPOS_ESTRUCTURA_CREADOS.discard(tipo)
        else:
            if tipo in TIPOS_TRANSFORMADOR:
                TIPOS_TRANSFORMADOR.remove(tipo)
            TIPOS_TRANSFORMADOR_CREADOS.discard(tipo)


# ==============================
# FORMULARIO PRINCIPAL
# ==============================
class Formulario(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("POSTE A POSTE – WuMDrex")
        self.ANCHO_BASE = 580
        self.ANCHO_CON_CAD = 640
        self.resize(self.ANCHO_BASE, 760)
        self.setMinimumSize(560, 720)

        try:
            myappid = 'wumdrex.posteaposte.app.1'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

        ruta_logo = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__))), "logo.ico")
        if os.path.exists(ruta_logo):
            self.setWindowIcon(QIcon(ruta_logo))

        self.df = None
        self.archivo_excel = None
        self.registros_filtrados = pd.DataFrame()
        self.indice_actual = 0
        self.modo_edicion = False
        self.registro_actual_idx = None
        self.nombre_hoja = None
        self.actualizando = False
        self.hay_cambios_pendientes = False
        self.postes_unicos = []
        self.poste_actual_index = -1
        self.ultimo_registro_creado = {}
        self.usando_machote = False
        self.registros_sin_guardar = 0
        
        self.modo_oscuro = True
        
        self.es_duplicado_actual = False
        self.es_pf_actual = False
        self.es_hijo_vii_actual = False
        self.cambios_catalogo = False

        self.init_ui()
        self.establecer_modo_inicial()

    def keyPressEvent(self, event):
        focused = QApplication.focusWidget()
        if isinstance(focused, QLineEdit):
            if event.key() == Qt.Key_Up:
                self.focusNextPrevChild(False)
                event.accept()
                return
            elif event.key() == Qt.Key_Down:
                self.focusNextPrevChild(True)
                event.accept()
                return
        super().keyPressEvent(event)

    def closeEvent(self, event):
        necesita_guardar = self.hay_cambios_pendientes or self.registros_sin_guardar > 0 or self.cambios_catalogo
        if necesita_guardar:
            msg = "Tienes cambios o registros nuevos sin guardar en el Excel.\n\n¿Deseas guardarlos antes de salir?"
            resp = QMessageBox.question(self, "Salir", msg, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel)
            
            if resp == QMessageBox.StandardButton.Yes:
                if self.hay_cambios_pendientes:
                    self.actualizar_dataframe_completo()
                self.guardar_excel()
                event.accept()
            elif resp == QMessageBox.StandardButton.No:
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

    def salir_aplicacion(self):
        self.close()

    def formatear_valor_entero(self, valor):
        try:
            num = float(valor)
            if num.is_integer():
                return str(int(num))
            else:
                return str(num)
        except:
            return str(valor).replace("nan", "").strip()

    def _tipo_poste_guardado(self):
        texto = self.cmb_tipo_poste.currentText()
        if self.btn_existente.isChecked() and texto:
            return "E " + texto
        return texto

    def init_ui(self):
        main = QVBoxLayout()
        main.setContentsMargins(10, 10, 10, 10)
        main.setSpacing(8)

        top = QHBoxLayout()
        top.setSpacing(6)

        self.btn_editar = QPushButton("Editar")
        self.btn_nuevo = QPushButton("Nuevo")
        self.btn_aceptar = QPushButton("Aceptar")
        self.btn_excel = QPushButton("Excel")
        self.btn_tema = QPushButton("☀️") 
        self.btn_tema.setFixedWidth(40)

        top.addWidget(self.btn_editar)
        top.addWidget(self.btn_nuevo)
        top.addWidget(self.btn_aceptar)
        top.addWidget(self.btn_excel)
        top.addWidget(self.btn_tema)
        top.addStretch()

        self.detalles_group = QWidget()
        self.detalles_group.setObjectName("DetallesGroup")
        detalles_layout = QVBoxLayout(self.detalles_group)
        detalles_layout.setContentsMargins(10, 8, 10, 8)
        detalles_layout.setSpacing(5)

        self.titulo_detalles = QLabel("Detalles del Poste")
        self.titulo_detalles.setStyleSheet("font-size: 14pt; font-weight: bold; padding-bottom: 2px; background-color: transparent; border: none;")
        detalles_layout.addWidget(self.titulo_detalles)

        linea = QWidget()
        linea.setFixedHeight(2)
        linea.setStyleSheet("background-color: #3498db;")
        detalles_layout.addWidget(linea)

        nav_layout = QHBoxLayout()
        reg_label = QLabel("Registro:")
        reg_label.setMinimumWidth(55)
        self.cmb_registro = QComboBox()
        self.cmb_registro.setMinimumWidth(110)

        self.btn_borrar_registro = QPushButton("Borrar Registro")
        lbl_potencia = QLabel("Voltaje:")
        self.cmb_potencia = QComboBox()
        self.cmb_potencia.addItems(["", "Linea 7.6 kV", "Linea 19.9 kV"])
        self.cmb_potencia.setMinimumWidth(110)

        nav_layout.addWidget(reg_label)
        nav_layout.addWidget(self.cmb_registro)
        nav_layout.addWidget(self.btn_borrar_registro)
        nav_layout.addWidget(lbl_potencia)
        nav_layout.addWidget(self.cmb_potencia)
        nav_layout.addStretch()
        detalles_layout.addLayout(nav_layout)

        poste_layout = QHBoxLayout()
        poste_label = QLabel("No. Poste:")
        poste_label.setMinimumWidth(55)
        self.txt_poste = QLineEdit()
        self.txt_poste.setMaximumWidth(75)

        self.btn_prev = QPushButton("< Anterior")
        self.btn_next = QPushButton("Siguiente >")
        self.btn_ultimo = QPushButton(">> Último")
        self.btn_ultimo.setVisible(False)

        poste_layout.addWidget(poste_label)
        poste_layout.addWidget(self.txt_poste)
        poste_layout.addWidget(self.btn_prev)
        poste_layout.addWidget(self.btn_next)
        poste_layout.addWidget(self.btn_ultimo)
        poste_layout.addStretch()
        detalles_layout.addLayout(poste_layout)

        pp_layout = QHBoxLayout()
        po_label = QLabel("Po:")
        po_label.setMinimumWidth(55)
        self.txt_po = QLineEdit()
        self.txt_po.setMaximumWidth(85)
        pf_label = QLabel("Pf:")
        pf_label.setMinimumWidth(30)
        self.txt_pf = QLineEdit()
        self.txt_pf.setMaximumWidth(85)

        pp_layout.addWidget(po_label)
        pp_layout.addWidget(self.txt_po)
        pp_layout.addWidget(pf_label)
        pp_layout.addWidget(self.txt_pf)
        pp_layout.addStretch()
        detalles_layout.addLayout(pp_layout)

        clasificacion_layout = QHBoxLayout()
        clasificacion_label = QLabel("Material:")
        clasificacion_layout.addWidget(clasificacion_label)

        self.btn_concreto = QPushButton("Concreto")
        self.btn_madera = QPushButton("Madera")
        self.btn_metal = QPushButton("Metal")
        self.btn_fibra = QPushButton("Fibra")

        for btn in [self.btn_concreto, self.btn_madera, self.btn_metal, self.btn_fibra]:
            btn.setCheckable(True)
            clasificacion_layout.addWidget(btn)
        self.btn_concreto.setChecked(True)

        self.btn_existente = QPushButton("Existente")
        self.btn_existente.setCheckable(True)
        clasificacion_layout.addWidget(self.btn_existente)
        estilo_clas = "QPushButton { background-color: #95a5a6; color: white; border: none; padding: 3px 8px; font-size: 7pt; } QPushButton:hover { background-color: #7f8c8d; } QPushButton:checked { background-color: #3498db; }"
        self.btn_existente.setStyleSheet(estilo_clas)

        clasificacion_layout.addStretch()
        detalles_layout.addLayout(clasificacion_layout)

        combo_poste_layout = QHBoxLayout()
        tipo_poste_label = QLabel("Tipo Poste:")
        tipo_poste_label.setMinimumWidth(55)
        self.cmb_tipo_poste = QComboBox()
        self.cmb_tipo_poste.addItems(["", *TIPOS_POSTE_CONCRETO])
        self.cmb_tipo_poste.setMinimumWidth(300)

        lbl_cimentado = QLabel("Cimentado:")
        self.rb_cim_si = QRadioButton("SI")
        self.rb_cim_no = QRadioButton("NO")
        self.rb_cim_no.setChecked(True)
        
        self.cim_group = QButtonGroup(self)
        self.cim_group.addButton(self.rb_cim_si)
        self.cim_group.addButton(self.rb_cim_no)

        combo_poste_layout.addWidget(tipo_poste_label)
        combo_poste_layout.addWidget(self.cmb_tipo_poste)
        combo_poste_layout.addWidget(lbl_cimentado)
        combo_poste_layout.addWidget(self.rb_cim_si)
        combo_poste_layout.addWidget(self.rb_cim_no)
        combo_poste_layout.addStretch()
        detalles_layout.addLayout(combo_poste_layout)

        linea_layout = QHBoxLayout()
        tipo_linea_label = QLabel("Tipo Línea en Poste:")
        tipo_linea_label.setMinimumWidth(55)
        self.cmb_tipo_linea = QComboBox()
        self.cmb_tipo_linea.addItems(["", "MEDIA TENSION", "MEDIA + BAJA TENSION", "BAJA TENSION"])
        self.cmb_tipo_linea.setMinimumWidth(115)

        linea_llega_label = QLabel("Llegan (m):")
        self.cmb_linea_llega = QComboBox()
        self.cmb_linea_llega.addItems(["", "MEDIA TENSION", "MEDIA + BAJA TENSION", "BAJA TENSION"])
        self.cmb_linea_llega.setMinimumWidth(115)

        self.txt_metros = QLineEdit()
        self.txt_metros.setPlaceholderText("Metros")
        self.txt_metros.setMaximumWidth(70)

        linea_layout.addWidget(tipo_linea_label)
        linea_layout.addWidget(self.cmb_tipo_linea)
        linea_layout.addWidget(linea_llega_label)
        linea_layout.addWidget(self.cmb_linea_llega)
        linea_layout.addWidget(self.txt_metros)
        linea_layout.addStretch()
        detalles_layout.addLayout(linea_layout)

        fases_layout = QHBoxLayout()
        fases_label = QLabel("Fases:")
        fases_label.setMinimumWidth(55)
        
        self.rb_1f = QRadioButton("  1F  ")
        self.rb_2f = QRadioButton("  2F  ")
        self.rb_3f = QRadioButton("  3F  ")
        self.rb_1f.setChecked(True) 
        
        self.fases_group = QButtonGroup(self)
        self.fases_group.addButton(self.rb_1f)
        self.fases_group.addButton(self.rb_2f)
        self.fases_group.addButton(self.rb_3f)

        self.lbl_tipo_bt = QLabel("Tipo:")
        self.lbl_tipo_bt.setVisible(False)
        self.rb_bt_1_0 = QRadioButton(" 1/0 ")
        self.rb_bt_2_0 = QRadioButton(" 2/0 ")
        self.tipo_bt_group = QButtonGroup(self)
        self.tipo_bt_group.addButton(self.rb_bt_1_0)
        self.tipo_bt_group.addButton(self.rb_bt_2_0)
        self.rb_bt_1_0.setVisible(False)
        self.rb_bt_2_0.setVisible(False)

        fases_layout.addWidget(fases_label)
        fases_layout.addWidget(self.rb_1f)
        fases_layout.addWidget(self.rb_2f)
        fases_layout.addWidget(self.rb_3f)
        fases_layout.addWidget(self.lbl_tipo_bt)
        fases_layout.addWidget(self.rb_bt_1_0)
        fases_layout.addWidget(self.rb_bt_2_0)
        fases_layout.addStretch()
        detalles_layout.addLayout(fases_layout)

        self.material_group = QWidget()
        self.material_group.setObjectName("MaterialGroup")
        material_layout = QVBoxLayout(self.material_group)
        material_layout.setContentsMargins(10, 4, 10, 4) 
        material_layout.setSpacing(4)

        titulo_material = QLabel("Agregar Material")
        titulo_material.setStyleSheet("font-size: 10pt; font-weight: bold; padding-bottom: 2px; background-color: transparent; border: none;")
        material_layout.addWidget(titulo_material)

        linea2 = QWidget()
        linea2.setFixedHeight(2)
        linea2.setStyleSheet("background-color: #27ae60;")
        material_layout.addWidget(linea2)

        mat_controls = QHBoxLayout()
        mat_label = QLabel("Material:")
        mat_label.setMinimumWidth(55)
        self.cmb_material = QComboBox()
        self.cmb_material.setMinimumWidth(145)
        
        self.lbl_cantidad = QLabel("Cantidad:")
        self.txt_cantidad = QLineEdit()
        self.txt_cantidad.setMaximumWidth(70)

        self.cmb_tipo_especial = QComboBox()
        self.cmb_tipo_especial.setMinimumWidth(190)
        self.cmb_tipo_especial.hide()

        self.btn_add_mat = QPushButton("Agregar")
        
        mat_controls.addWidget(mat_label)
        mat_controls.addWidget(self.cmb_material)
        mat_controls.addWidget(self.lbl_cantidad)
        mat_controls.addWidget(self.txt_cantidad)
        mat_controls.addWidget(self.cmb_tipo_especial)
        mat_controls.addWidget(self.btn_add_mat)
        mat_controls.addStretch()
        material_layout.addLayout(mat_controls)

        quitar_layout = QHBoxLayout()
        self.btn_quitar_mat = QPushButton("Eliminar")
        self.btn_quitar_mat.setEnabled(False)
        
        self.btn_m_anterior = QPushButton("M. Anterior")
        self.btn_m_anterior.setStyleSheet("QPushButton { background-color: #f39c12; color: white; border: none; padding: 4px 14px; font-weight: bold; } QPushButton:hover { background-color: #e67e22; }")
        self.btn_m_anterior.setVisible(False)
        
        quitar_layout.addWidget(self.btn_quitar_mat)
        quitar_layout.addWidget(self.btn_m_anterior)
        quitar_layout.addStretch()
        material_layout.addLayout(quitar_layout)

        visor_layout = QHBoxLayout()
        lista_layout = QVBoxLayout()
        lista_layout.setContentsMargins(0, 0, 0, 0)
        
        lista_titulo = QLabel("Lista de Materiales del Poste:")
        lista_titulo.setStyleSheet("font-weight: 600; margin-top: 3px; background-color: transparent; border: none;")
        lista_layout.addWidget(lista_titulo)

        self.lst_materiales = QListWidget()
        self.lst_materiales.setSelectionMode(QListWidget.ExtendedSelection)
        self.lst_materiales.setMinimumHeight(160) 
        self.lst_materiales.setMaximumHeight(240)
        lista_layout.addWidget(self.lst_materiales)
        visor_layout.addLayout(lista_layout, 6)
        
        from PySide6.QtWidgets import QGraphicsView, QGraphicsScene
        from PySide6.QtGui import QPainter
        
        self.panel_cad = QWidget()
        cad_layout = QVBoxLayout(self.panel_cad)
        cad_layout.setContentsMargins(5, 0, 0, 0)
        lbl_canvas_row = QHBoxLayout()
        lbl_canvas = QLabel("Vista Previa (AutoCAD)")
        lbl_canvas.setAlignment(Qt.AlignCenter)
        lbl_canvas.setStyleSheet("font-weight: bold; margin-top: 3px; background-color: transparent; border: none;")
        lbl_canvas_row.addWidget(lbl_canvas)

        self.btn_zoom_menos = QPushButton("-")
        self.btn_zoom_mas = QPushButton("+")
        self.btn_zoom_ajustar = QPushButton("Ajustar")
        for b in (self.btn_zoom_menos, self.btn_zoom_mas, self.btn_zoom_ajustar):
            b.setFixedHeight(22)
            b.setStyleSheet("QPushButton { padding: 0px 8px; font-size: 8pt; }")
        self.btn_zoom_menos.setFixedWidth(26)
        self.btn_zoom_mas.setFixedWidth(26)
        lbl_canvas_row.addWidget(self.btn_zoom_menos)
        lbl_canvas_row.addWidget(self.btn_zoom_mas)
        lbl_canvas_row.addWidget(self.btn_zoom_ajustar)
        cad_layout.addLayout(lbl_canvas_row)
        
        self.escena = QGraphicsScene()
        self.vista_grafica = VistaCAD(self.escena)
        self.vista_grafica.setRenderHint(QPainter.Antialiasing)
        self.vista_grafica.setStyleSheet("background-color: #212830; border: 2px solid #555; border-radius: 4px;")
        
        cad_layout.addWidget(self.vista_grafica)
        self.lbl_bloques_faltantes = QLabel("")
        self.lbl_bloques_faltantes.setWordWrap(True)
        self.lbl_bloques_faltantes.setStyleSheet("color: #e74c3c; font-size: 7pt; background-color: transparent; border: none;")
        cad_layout.addWidget(self.lbl_bloques_faltantes)

        self.btn_zoom_menos.clicked.connect(self.vista_grafica.zoom_out)
        self.btn_zoom_mas.clicked.connect(self.vista_grafica.zoom_in)
        self.btn_zoom_ajustar.clicked.connect(self.vista_grafica.ajustar_a_contenido)
        visor_layout.addWidget(self.panel_cad, 4)
        self.panel_cad.setVisible(False)
        self._cad_visible = False
        material_layout.addLayout(visor_layout)

        bottom = QHBoxLayout()
        self.btn_guardar = QPushButton("Guardar al Excel")
        self.btn_guardar.setMinimumHeight(32)
        
        self.btn_salir = QPushButton("Salir")
        self.btn_salir.setMinimumHeight(32)

        bottom.addWidget(self.btn_guardar)
        bottom.addWidget(self.btn_salir)

        main.addLayout(top)
        main.addWidget(self.detalles_group)
        main.addWidget(self.material_group)
        main.addLayout(bottom)
        self.setLayout(main)

        self.aplicar_tema()

        self.txt_poste.textChanged.connect(self.actualizar_dibujo)
        self.cmb_tipo_linea.currentTextChanged.connect(self.actualizar_dibujo)
        self.rb_cim_si.toggled.connect(self.actualizar_dibujo)
        self.rb_1f.toggled.connect(self.actualizar_dibujo)
        self.rb_2f.toggled.connect(self.actualizar_dibujo)
        self.rb_3f.toggled.connect(self.actualizar_dibujo)
        self.cmb_tipo_poste.currentTextChanged.connect(self.actualizar_dibujo)

        QShortcut(QKeySequence("Ctrl+S"), self).activated.connect(self.atajo_guardar)
        QShortcut(QKeySequence("Ctrl+1"), self).activated.connect(lambda: self.aplicar_combo_materiales("1"))
        QShortcut(QKeySequence("Ctrl+2"), self).activated.connect(lambda: self.aplicar_combo_materiales("2"))
        QShortcut(QKeySequence("Ctrl+3"), self).activated.connect(lambda: self.aplicar_combo_materiales("3"))
        QShortcut(QKeySequence("Ctrl+4"), self).activated.connect(lambda: self.aplicar_combo_materiales("4"))
        QShortcut(QKeySequence("Ctrl+5"), self).activated.connect(lambda: self.aplicar_combo_materiales("5"))
        QShortcut(QKeySequence("Ctrl+6"), self).activated.connect(lambda: self.aplicar_combo_materiales("6"))
        QShortcut(QKeySequence("Ctrl+7"), self).activated.connect(lambda: self.aplicar_combo_materiales("7"))
        QShortcut(QKeySequence("Ctrl+8"), self).activated.connect(lambda: self.aplicar_combo_materiales("8"))
        QShortcut(QKeySequence("Ctrl+9"), self).activated.connect(lambda: self.aplicar_combo_materiales("9"))
        QShortcut(QKeySequence("Ctrl+0"), self).activated.connect(lambda: self.aplicar_combo_materiales("0"))
        
        QShortcut(QKeySequence("Ctrl+Q"), self).activated.connect(self.btn_add_mat.click)
        QShortcut(QKeySequence("Alt+M"), self).activated.connect(self.atajo_madera)
        QShortcut(QKeySequence("Alt+C"), self).activated.connect(self.atajo_concreto)
        QShortcut(QKeySequence("Alt+D"), self).activated.connect(self.atajo_metros)
        QShortcut(QKeySequence("Alt+T"), self).activated.connect(self.cmb_tipo_poste.setFocus)
        QShortcut(QKeySequence("Alt+E"), self).activated.connect(self.btn_editar.click)
        QShortcut(QKeySequence("Alt+N"), self).activated.connect(self.btn_nuevo.click)
        QShortcut(QKeySequence("Ctrl+M"), self).activated.connect(self.cmb_material.setFocus)
        QShortcut(QKeySequence("Alt+P"), self).activated.connect(self.atajo_poste)
        QShortcut(QKeySequence("Alt+V"), self).activated.connect(self.convertir_ancla_vertical)
        
        QShortcut(QKeySequence("Ctrl+G"), self).activated.connect(lambda: self.aplicar_combo_materiales("G"))
        QShortcut(QKeySequence("Ctrl+F"), self).activated.connect(lambda: self.aplicar_combo_materiales("F"))
        QShortcut(QKeySequence("Ctrl+T"), self).activated.connect(lambda: self.aplicar_combo_materiales("T"))
        QShortcut(QKeySequence("Ctrl+D"), self).activated.connect(lambda: self.aplicar_combo_materiales("D"))
        QShortcut(QKeySequence("Ctrl+W"), self).activated.connect(lambda: self.aplicar_combo_materiales("W"))
        QShortcut(QKeySequence("Ctrl+B"), self).activated.connect(lambda: self.aplicar_combo_materiales("B"))
        
        QShortcut(QKeySequence("Ctrl+Y"), self).activated.connect(self.atajo_toggle_cim)
        QShortcut(QKeySequence("Alt+Q"), self).activated.connect(self.btn_m_anterior.click)

        QShortcut(QKeySequence("Alt+L"), self).activated.connect(self.atajo_lista_materiales)
        QShortcut(QKeySequence("Alt+H"), self).activated.connect(lambda: self.cmb_tipo_linea.setCurrentText("MEDIA TENSION") if self.cmb_tipo_linea.isEnabled() else None)
        QShortcut(QKeySequence("Alt+J"), self).activated.connect(lambda: self.cmb_tipo_linea.setCurrentText("MEDIA + BAJA TENSION") if self.cmb_tipo_linea.isEnabled() else None)
        QShortcut(QKeySequence("Alt+K"), self).activated.connect(lambda: self.cmb_tipo_linea.setCurrentText("BAJA TENSION") if self.cmb_tipo_linea.isEnabled() else None)
        QShortcut(QKeySequence("Alt+B"), self).activated.connect(self.toggle_tema)
        QShortcut(QKeySequence("F1"), self).activated.connect(self.mostrar_ayuda)
        QShortcut(QKeySequence("F2"), self).activated.connect(self.alternar_gestor_materiales)
        QShortcut(QKeySequence("F3"), self).activated.connect(self.alternar_gestor_postes)
        QShortcut(QKeySequence("F5"), self).activated.connect(self.alternar_visor_cad)
        QShortcut(QKeySequence("Ctrl+Right"), self).activated.connect(self.navegar_siguiente)
        QShortcut(QKeySequence("Ctrl+Left"), self).activated.connect(self.navegar_anterior)
        QShortcut(QKeySequence("Ctrl+Shift+Right"), self).activated.connect(self.ir_a_ultimo)
        QShortcut(QKeySequence("Ctrl+Shift+Left"), self).activated.connect(self.ir_a_primero)

        shortcut_del = QShortcut(QKeySequence("Del"), self.lst_materiales)
        shortcut_del.setContext(Qt.WidgetShortcut)
        shortcut_del.activated.connect(self.quitar_material)

        self.btn_tema.clicked.connect(self.toggle_tema)
        self.btn_excel.clicked.connect(self.mostrar_menu_excel)
        self.btn_editar.clicked.connect(self.click_editar) 
        self.btn_nuevo.clicked.connect(self.establecer_modo_nuevo)
        self.btn_aceptar.clicked.connect(self.aceptar_nuevo)
        self.btn_borrar_registro.clicked.connect(self.borrar_registro)
        self.txt_poste.textChanged.connect(self.buscar_por_poste)
        self.txt_poste.textChanged.connect(self.actualizar_titulo)
        self.cmb_registro.currentIndexChanged.connect(self.cargar_registro)
        self.btn_prev.clicked.connect(self.navegar_anterior)
        self.btn_next.clicked.connect(self.navegar_siguiente)
        self.btn_ultimo.clicked.connect(self.ir_a_ultimo)
        self.cmb_material.currentTextChanged.connect(self.cambiar_tipo_input)
        self.btn_add_mat.clicked.connect(self.agregar_material)
        self.btn_quitar_mat.clicked.connect(self.quitar_material)
        self.btn_m_anterior.clicked.connect(self.cargar_materiales_anterior) 
        
        
        self.lst_materiales.itemSelectionChanged.connect(lambda: self.btn_quitar_mat.setEnabled(len(self.lst_materiales.selectedItems()) > 0))
        self.lst_materiales.itemClicked.connect(self.seleccionar_material)
        self.btn_guardar.clicked.connect(self.guardar_accion)
        self.btn_salir.clicked.connect(self.salir_aplicacion)
        self.cmb_potencia.currentTextChanged.connect(self.cambiar_voltaje)

        self.btn_concreto.clicked.connect(lambda: self.cambiar_clasificacion_poste("Concreto"))
        self.btn_madera.clicked.connect(lambda: self.cambiar_clasificacion_poste("Madera"))
        self.btn_metal.clicked.connect(lambda: self.cambiar_clasificacion_poste("Metal"))
        self.btn_fibra.clicked.connect(lambda: self.cambiar_clasificacion_poste("Fibra"))

        self.txt_metros.textChanged.connect(self.validar_metros)
        self.txt_cantidad.textChanged.connect(self.validar_cantidad)

        self.txt_po.textChanged.connect(self.marcar_cambio)
        self.txt_pf.textChanged.connect(self.marcar_cambio)
        self.cmb_tipo_poste.currentTextChanged.connect(self.marcar_cambio)
        self.cmb_tipo_linea.currentTextChanged.connect(self.marcar_cambio)
        self.cmb_linea_llega.currentTextChanged.connect(self.marcar_cambio)
        self.txt_metros.textChanged.connect(self.marcar_cambio)
        self.rb_1f.toggled.connect(self.marcar_cambio)
        self.rb_2f.toggled.connect(self.marcar_cambio)
        self.rb_3f.toggled.connect(self.marcar_cambio)
        self.rb_1f.toggled.connect(self.cambiar_tipo_input)
        self.rb_2f.toggled.connect(self.cambiar_tipo_input)
        self.rb_3f.toggled.connect(self.cambiar_tipo_input)
        self.rb_bt_1_0.toggled.connect(self.marcar_cambio)
        self.rb_bt_2_0.toggled.connect(self.marcar_cambio)
        self.rb_cim_si.toggled.connect(self.marcar_cambio)
        self.rb_cim_no.toggled.connect(self.marcar_cambio)

        self.cmb_linea_llega.currentTextChanged.connect(self.controlar_tipo_bt)
        
        self.cmb_tipo_linea.currentTextChanged.connect(self.filtrar_linea_llega)
        self.cmb_tipo_linea.currentTextChanged.connect(self.filtrar_materiales)
        self.cmb_linea_llega.currentTextChanged.connect(self.controlar_estado_metros)
        
        self.txt_poste.textChanged.connect(self.validar_reglas_ui)
        self.txt_po.textChanged.connect(self.validar_reglas_ui)
        self.txt_pf.textChanged.connect(self.validar_reglas_ui)

        self.cmb_tipo_poste.currentTextChanged.connect(self.validar_estado_cimentacion)

        self.filtrar_materiales()
        self.controlar_estado_metros()

    def abrir_gestor_postes(self):
        dlg = DialogoGestorPostes(self)
        self._gestor_postes = dlg
        dlg.exec()
        self._gestor_postes = None

    def alternar_gestor_postes(self):
        if getattr(self, '_gestor_postes', None) is not None:
            self._gestor_postes.reject()
        else:
            self.abrir_gestor_postes()

    def abrir_gestor_materiales(self):
        dlg = DialogoGestorMateriales(self)
        self._gestor_materiales = dlg
        dlg.exec()
        self._gestor_materiales = None

    def alternar_gestor_materiales(self):
        if getattr(self, '_gestor_materiales', None) is not None:
            self._gestor_materiales.reject()
        else:
            self.abrir_gestor_materiales()

    def mostrar_ayuda(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Atajos de Teclado")
        dialog.resize(550, 650)
        layout = QVBoxLayout(dialog)
        
        atajos = [
            ("Ctrl + S", "Guardar al Excel / Aceptar"),
            ("Ctrl + 1", "Materiales(Estructura = 'I')"),
            ("Ctrl + 2", "Materiales(Estructura = 'II', Anclas Dobles = 1)"),
            ("Ctrl + 3", "Materiales(Estructura = 'III')"),
            ("Ctrl + 4", "Materiales(Estructura = 'IV', Anclas Dobles = 2)"),
            ("Ctrl + 5", "Materiales(Estructura = 'V', Anclas Dobles = 3)"),
            ("Ctrl + 6", "Materiales(Estructura = 'VI', Anclas Dobles = 1)"),
            ("Ctrl + 7", "Materiales(Estructura = 'VII', Anclas Dobles = 6)"),
            ("Ctrl + 8", "Materiales(Estructura = 'VIII')"),
            ("Ctrl + 9", "Materiales(Estructura = 'IX', Anclas Dobles = 1)"),
            ("Ctrl + 0", "Materiales(Estructura = 'XI', Anclas Dobles = 3)"),
            ("Ctrl + F", "Agregar: Tierra(1), Ancla Simple(1), Remate Simple(1)"),
            ("Ctrl + G", "Agregar: Suplementos por Excavación (1)"),
            ("Ctrl + D", "Agregar: Remate Doble = 1"),
            ("Ctrl + T", "Agregar: Tierras = 1"),
            ("Ctrl + W", "Agregar: Ancla Simple = 1"),
            ("Ctrl + B", "Agregar: Estructura Remate Simple = 1"),
            ("Ctrl + Y", "Alternar si el poste está Cimentado (SI/NO)"),
            ("Ctrl + Q", "Botón 'Agregar' material"),
            ("Ctrl + M", "Enfocar la lista desplegable de 'Material'"),
            ("Alt + Q", "Botón 'M. Anterior' (Material Anterior)"),
            ("Alt + M", "Marcar material de poste: 'Madera' y enfocar Tipo"),
            ("Alt + C", "Marcar material de poste: 'Concreto' y enfocar Tipo"),
            ("Alt + P", "Enfocar y seleccionar el texto en 'No. Poste'"),
            ("Alt + V", "Convertir anclas simples/dobles en verticales (y viceversa)"),
            ("Alt + D", "Enfocar y seleccionar el texto en 'Metros'"),
            ("Alt + T", "Enfocar lista desplegable 'Tipo Poste'"),
            ("Alt + E", "Ingresar directamente al Modo Editar"),
            ("Alt + N", "Ingresar directamente al Modo Nuevo"),
            ("Alt + L", "Enfocar Lista de Materiales y seleccionar el primer elemento"),
            ("Alt + H", "Seleccionar 'MEDIA TENSION' en Tipo Línea"),
            ("Alt + J", "Seleccionar 'MEDIA + BAJA TENSION' en Tipo Línea"),
            ("Alt + K", "Seleccionar 'BAJA TENSION' en Tipo Línea"),
            ("Alt + B", "Alternar entre Modo Día y Modo Noche"),
            ("Shift + Clic", "Seleccionar múltiples materiales en la lista"),
            ("Suprimir (Del)", "Eliminar todos los materiales seleccionados"),
            ("Arriba / Abajo", "Navegar arriba/abajo entre cuadros de texto"),
            ("F1", "Mostrar esta ayuda"),
            ("F2", "Abrir/cerrar gestor de materiales, estructuras y transformadores"),
            ("F3", "Abrir/cerrar gestor de postes"),
            ("F5", "Mostrar/Ocultar Vista Previa (AutoCAD)"),
            ("Ctrl + Flecha derecha", "Siguiente registro/poste"),
            ("Ctrl + Flecha izquierda", "Registro/poste anterior"),
            ("Ctrl + Shift + Flecha derecha", "Ir al último registro"),
            ("Ctrl + Shift + Flecha izquierda", "Ir al primer registro")
        ]
        
        tabla = QTableWidget(len(atajos), 2)
        tabla.setHorizontalHeaderLabels(["Atajo", "Descripción"])
        tabla.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        tabla.setEditTriggers(QTableWidget.NoEditTriggers)
        
        for i, (atajo, desc) in enumerate(atajos):
            tabla.setItem(i, 0, QTableWidgetItem(atajo))
            tabla.setItem(i, 1, QTableWidgetItem(desc))
            
        layout.addWidget(tabla)
        dialog.exec()

    def actualizar_titulo(self):
        poste = self.txt_poste.text().strip()
        if poste:
            self.titulo_detalles.setText(f"Detalles del Poste ({poste})")
        else:
            self.titulo_detalles.setText("Detalles del Poste")

    def atajo_guardar(self):
        if self.modo_edicion:
            if self.btn_guardar.isEnabled():
                self.btn_guardar.click()
        else:
            if self.btn_aceptar.isEnabled():
                self.btn_aceptar.click()

    def atajo_lista_materiales(self):
        if self.lst_materiales.isEnabled() and self.lst_materiales.count() > 0:
            self.lst_materiales.setFocus()
            self.lst_materiales.setCurrentRow(0)

    def atajo_madera(self):
        self.btn_madera.click()
        self.cmb_tipo_poste.setFocus()

    def atajo_concreto(self):
        self.btn_concreto.click()
        self.cmb_tipo_poste.setFocus()

    def atajo_metros(self):
        if self.txt_metros.isEnabled():
            self.txt_metros.setFocus()
            self.txt_metros.selectAll()

    def atajo_poste(self):
        if self.txt_poste.isEnabled():
            self.txt_poste.setFocus()
            self.txt_poste.selectAll()

    def atajo_toggle_cim(self):
        if self.rb_cim_si.isEnabled():
            if self.rb_cim_si.isChecked():
                self.rb_cim_no.setChecked(True)
            else:
                self.rb_cim_si.setChecked(True)

    def aplicar_combo_materiales(self, tipo):
        if not self.material_group.isEnabled() or not self.cmb_material.isEnabled():
            return
            
        mats = []
        if tipo == "F":
            mats = [("ANCLA SIMPLE", "1"), ("TIERRAS", "1"), ("ESTRUCTURA REMATE SIMPLE", "1")]
        elif tipo == "G":
            mats = [("SUPLEMENTOS POR EXCAVACION", "1")]
        elif tipo == "T":
            mats = [("TIERRAS", "1")]
        elif tipo == "D":
            mats = [("ESTRUCTURA REMATE DOBLE", "1")]
        elif tipo == "W":
            mats = [("ANCLA SIMPLE", "1")]
        elif tipo == "B":
            mats = [("ESTRUCTURA REMATE SIMPLE", "1")]
        elif tipo == "1":
            mats = [("ESTRUCTURA", "I")]
        elif tipo == "2":
            mats = [("ESTRUCTURA", "II"), ("ANCLAS DOBLES", "1")]
        elif tipo == "3":
            mats = [("ESTRUCTURA", "III")]
        elif tipo == "4":
            mats = [("ESTRUCTURA", "IV"), ("ANCLAS DOBLES", "2")]
        elif tipo == "5":
            mats = [("ESTRUCTURA", "V"), ("ANCLAS DOBLES", "3")]
        elif tipo == "6":
            mats = [("ESTRUCTURA", "VI"), ("ANCLAS DOBLES", "1")]
        elif tipo == "7":
            mats = [("ESTRUCTURA", "VII"), ("ANCLAS DOBLES", "6")]
        elif tipo == "8":
            mats = [("ESTRUCTURA", "VIII")]
        elif tipo == "9":
            mats = [("ESTRUCTURA", "IX"), ("ANCLAS DOBLES", "1")]
        elif tipo == "0":
            mats = [("ESTRUCTURA", "XI"), ("ANCLAS DOBLES", "3")]
            
        if self.rb_cim_si.isChecked():
            mats = [(m, v) for m, v in mats if m not in ["ANCLAS DOBLES", "ANCLAS DOBLES VERTICALES", "ANCLA SIMPLE", "ANCLA SIMPLE VERTICAL"]]
            
        for mat, val in mats:
            if self.cmb_material.findText(mat) == -1:
                continue
                
            found = False
            for i in range(self.lst_materiales.count()):
                if self.lst_materiales.item(i).text().startswith(mat + " – "):
                    self.lst_materiales.item(i).setText(f"{mat} – {val}")
                    found = True
                    break
            if not found:
                self.lst_materiales.addItem(f"{mat} – {val}")
                
        self.btn_m_anterior.setVisible(False)
        self.actualizar_conectores_bt()
        self.marcar_cambio()
        if hasattr(self, 'actualizar_dibujo'):
            self.actualizar_dibujo()

    def toggle_tema(self):
        self.modo_oscuro = not self.modo_oscuro
        self.btn_tema.setText("☀️" if self.modo_oscuro else "🌙")
        self.aplicar_tema()

    def aplicar_tema(self):
        if self.modo_oscuro:
            bg_color = "#2c3e50"
            text_color = "#ecf0f1"
            input_bg = "#34495e"
            border_color = "#7f8c8d"
            rb_bg = "#34495e"
            rb_hover = "#7f8c8d"
            group_bg = "#34495e"
            group_border = "#7f8c8d"
            list_sel = "#2980b9"
        else:
            bg_color = "#f5f6fa"
            text_color = "#2c3e50"
            input_bg = "white"
            border_color = "#bdc3c7"
            rb_bg = "#d6eaf8"
            rb_hover = "#aed6f1"
            group_bg = "white"
            group_border = "#e0e0e0"
            list_sel = "#3498db"

        css = f"""
            QWidget {{
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 8pt;
                background-color: {bg_color};
                color: {text_color};
            }}
            QLabel {{
                color: {text_color};
                font-weight: 300;
                background-color: transparent;
            }}
            QLineEdit, QComboBox {{
                padding: 3px 6px;
                border: 2px solid {border_color};
                border-radius: 4px;
                background-color: {input_bg};
                color: {text_color};
                min-height: 18px;
            }}
            QLineEdit:focus, QComboBox:focus {{
                border: 2px solid #3498db;
            }}
            QPushButton {{
                padding: 4px 10px;
                border-radius: 4px;
                font-weight: 300;
                min-height: 24px;
            }}
            QListWidget {{
                border: 2px solid {border_color};
                border-radius: 4px;
                padding: 3px;
                background-color: {input_bg};
                color: {text_color};
            }}
            QListWidget::item {{
                padding: 0px;
                border-bottom: 1px solid {border_color};
            }}
            QListWidget::item:selected {{
                background-color: {list_sel};
                color: white;
                border-radius: 3px;
            }}
            QRadioButton {{
                spacing: 2px;
                font-size: 9pt;
                color: {text_color};
                padding: 4px 10px;
                border-radius: 4px;
                background-color: {rb_bg};
                border: 1px solid {border_color};
            }}
            QRadioButton:checked {{
                background-color: #3498db;
                color: white;
                border: 2px solid #21618c;
                font-weight: bold;
            }}
            QRadioButton:hover {{
                background-color: {rb_hover};
            }}
            QRadioButton::indicator {{
                width: 0px;
                height: 0px;
                border: none;
                background: transparent;
            }}
            #DetallesGroup, #MaterialGroup {{
                background-color: {group_bg};
                border-radius: 8px;
                border: 1px solid {group_border};
            }}
        """
        self.setStyleSheet(css)

        estilo_base_btn = "QPushButton { color: white; border: none; } QPushButton:disabled { background-color: #95a5a6; }"
        self.btn_editar.setStyleSheet(estilo_base_btn + "QPushButton { background-color: #3498db; } QPushButton:hover { background-color: #2980b9; }")
        self.btn_nuevo.setStyleSheet(estilo_base_btn + "QPushButton { background-color: #27ae60; } QPushButton:hover { background-color: #229954; }")
        self.btn_aceptar.setStyleSheet(estilo_base_btn + "QPushButton { background-color: #16a085; } QPushButton:hover { background-color: #138d75; }")
        self.btn_excel.setStyleSheet(estilo_base_btn + "QPushButton { background-color: #16a085; } QPushButton:hover { background-color: #138d75; }")
        
        estilo_nav = "QPushButton { background-color: #3498db; color: white; border: none; padding: 4px 10px; } QPushButton:hover { background-color: #2980b9; } QPushButton:disabled { background-color: #95a5a6; }"
        self.btn_prev.setStyleSheet(estilo_nav)
        self.btn_next.setStyleSheet(estilo_nav)
        self.btn_ultimo.setStyleSheet(estilo_nav)

        self.btn_borrar_registro.setStyleSheet("QPushButton { background-color: #e74c3c; color: white; border: none; padding: 4px 10px; } QPushButton:hover { background-color: #c0392b; } QPushButton:disabled { background-color: #95a5a6; }")
        self.btn_add_mat.setStyleSheet("QPushButton { background-color: #27ae60; color: white; border: none; padding: 4px 14px; } QPushButton:hover { background-color: #229954; } QPushButton:disabled { background-color: #95a5a6; }")
        self.btn_quitar_mat.setStyleSheet("QPushButton { background-color: #e74c3c; color: white; border: none; padding: 4px 14px; } QPushButton:hover { background-color: #c0392b; } :disabled { background-color: #bdc3c7; }")
        self.btn_guardar.setStyleSheet("QPushButton { background-color: #3498db; color: white; border: none; font-size: 9pt; font-weight: bold; } QPushButton:hover { background-color: #2980b9; } :disabled { background-color: #bdc3c7; }")
        self.btn_salir.setStyleSheet("QPushButton { background-color: #e74c3c; color: white; border: none; font-size: 9pt; font-weight: bold; } QPushButton:hover { background-color: #c0392b; }")
        
        estilo_clas = "QPushButton { background-color: #95a5a6; color: white; border: none; padding: 3px 8px; font-size: 7pt; } QPushButton:hover { background-color: #7f8c8d; } QPushButton:checked { background-color: #3498db; }"
        for btn in [self.btn_concreto, self.btn_madera, self.btn_metal, self.btn_fibra]:
            btn.setStyleSheet(estilo_clas)

    def establecer_modo_inicial(self):
        self.modo_edicion = False
        self.registro_actual_idx = None
        
        self.btn_editar.setEnabled(False)
        self.btn_nuevo.setEnabled(False)
        self.btn_aceptar.setEnabled(False)
        self.btn_guardar.setEnabled(False)
        self.btn_borrar_registro.setEnabled(False)
        
        self.btn_excel.setEnabled(True)
        self.btn_salir.setEnabled(True)

        self.detalles_group.setEnabled(False)
        self.material_group.setEnabled(False)

    def controlar_tipo_bt(self):
        linea_llega = self.cmb_linea_llega.currentText()
        es_bt_o_mbt = linea_llega in ["BAJA TENSION", "MEDIA + BAJA TENSION"]
        self.lbl_tipo_bt.setVisible(es_bt_o_mbt)
        self.rb_bt_1_0.setVisible(es_bt_o_mbt)
        self.rb_bt_2_0.setVisible(es_bt_o_mbt)
        
        if not es_bt_o_mbt:
            self.tipo_bt_group.setExclusive(False)
            self.rb_bt_1_0.setChecked(False)
            self.rb_bt_2_0.setChecked(False)
            self.tipo_bt_group.setExclusive(True)
        else:
            if not self.actualizando and not self.rb_bt_1_0.isChecked() and not self.rb_bt_2_0.isChecked():
                self.rb_bt_2_0.setChecked(True)

    def filtrar_linea_llega(self):
        tipo_linea = self.cmb_tipo_linea.currentText()
        actual = self.cmb_linea_llega.currentText()
        
        self.cmb_linea_llega.blockSignals(True)
        self.cmb_linea_llega.clear()
        
        opciones = [""]
        if tipo_linea == "MEDIA TENSION":
            opciones.append("MEDIA TENSION")
            self.cmb_linea_llega.setEnabled(True)
        elif tipo_linea == "BAJA TENSION":
            opciones.append("BAJA TENSION")
            self.cmb_linea_llega.setEnabled(True)
        elif tipo_linea == "MEDIA + BAJA TENSION":
            opciones.extend(["MEDIA TENSION", "MEDIA + BAJA TENSION", "BAJA TENSION"])
            self.cmb_linea_llega.setEnabled(True)
        else:
            self.cmb_linea_llega.setEnabled(False)
            
        self.cmb_linea_llega.addItems(opciones)
        if actual in opciones:
            self.cmb_linea_llega.setCurrentText(actual)
            
        self.cmb_linea_llega.blockSignals(False)
        self.controlar_estado_metros()

    def filtrar_materiales(self):
        tipo_linea = self.cmb_tipo_linea.currentText()
        mat_actual = self.cmb_material.currentText()
        
        self.cmb_material.blockSignals(True)
        self.cmb_material.clear()
        
        items_validos = [""]
        for m in MATERIALES_COLUMNAS:
            if m == "CIMENTACION": continue
            
            if tipo_linea == "MEDIA TENSION":
                if m in ["ESTRUCTURA REMATE DOBLE", "ESTRUCTURA REMATE SIMPLE", "CONECTORES BT"]:
                    continue
            elif tipo_linea == "BAJA TENSION":
                if m in ["ESTRUCTURA", "ANILLO DE TIERRA", "TRANSFORMADORES", "AMORTIGUADORES", "CONECTORES BT"]:
                    continue
            elif tipo_linea == "MEDIA + BAJA TENSION":
                if m in ["CONECTORES BT"]:
                    continue
            
            if getattr(self, 'es_duplicado_actual', False):
                if m not in ["ESTRUCTURA"]:
                    continue
            
            items_validos.append(m)
            
        self.cmb_material.addItems(items_validos)
        if mat_actual in items_validos:
            self.cmb_material.setCurrentText(mat_actual)
            
        self.cmb_material.blockSignals(False)
        
        bloquear_todo = False
        if getattr(self, 'es_pf_actual', False) and not getattr(self, 'es_duplicado_actual', False):
            bloquear_todo = True
            
        if getattr(self, 'es_hijo_vii_actual', False):
            bloquear_todo = True
            
        if bloquear_todo:
            self.cmb_material.setEnabled(False)
            self.txt_cantidad.setEnabled(False)
            self.cmb_tipo_especial.setEnabled(False)
            self.btn_add_mat.setEnabled(False)
        else:
            self.cmb_material.setEnabled(True)
            self.txt_cantidad.setEnabled(True)
            self.cmb_tipo_especial.setEnabled(True)
            self.btn_add_mat.setEnabled(True)
            self.cambiar_tipo_input() 

    def controlar_estado_metros(self):
        if not self.cmb_linea_llega.currentText():
            self.txt_metros.clear()
            self.txt_metros.setEnabled(False)
        else:
            self.txt_metros.setEnabled(True)

    def validar_estado_cimentacion(self):
        tiene_poste = bool(self.cmb_tipo_poste.currentText().strip())
        self.rb_cim_si.setEnabled(tiene_poste)
        self.rb_cim_no.setEnabled(tiene_poste)
        if not tiene_poste:
            self.rb_cim_no.setChecked(True)

    def validar_reglas_ui(self):
        if self.actualizando: return  
        
        if getattr(self, 'es_hijo_vii_actual', False):
            self.filtrar_materiales()
            return
            
        poste = self.txt_poste.text().strip()
        po = self.txt_po.text().strip()
        pf = self.txt_pf.text().strip()

        es_duplicado = False
        es_poste_igual_pf_existente = False
        linea_copiada = ""

        if self.df is not None and not self.df.empty:
            if "_hijo_vii" in self.df.columns:
                mask_normales = self.df["_hijo_vii"] != True
            else:
                mask_normales = pd.Series(True, index=self.df.index)
            
            def get_prev_match(mask):
                valid_idx = self.df[mask].index.tolist()
                if self.modo_edicion and self.registro_actual_idx is not None:
                    valid_idx = [i for i in valid_idx if i < self.registro_actual_idx]
                if valid_idx:
                    return valid_idx[0]
                return None

            if poste and po and pf:
                mask_dup = (
                    mask_normales &
                    (self.df["NO POSTE"].apply(lambda x: self.formatear_valor_entero(x)) == poste) & 
                    (self.df["Po"].apply(lambda x: self.formatear_valor_entero(x)) == po) & 
                    (self.df["Pf"].apply(lambda x: self.formatear_valor_entero(x)) == pf)
                )
                match_idx = get_prev_match(mask_dup)
                if match_idx is not None:
                    es_duplicado = True
                    linea_copiada = str(self.df.at[match_idx, "LINEA"]).replace("nan", "")
            
            if poste and pf and not es_duplicado:
                mask_pf = (
                    mask_normales &
                    (self.df["NO POSTE"].apply(lambda x: self.formatear_valor_entero(x)) == poste) & 
                    (self.df["Pf"].apply(lambda x: self.formatear_valor_entero(x)) == pf)
                )
                match_idx = get_prev_match(mask_pf)
                if match_idx is not None:
                    es_poste_igual_pf_existente = True
                    linea_copiada = str(self.df.at[match_idx, "LINEA"]).replace("nan", "")

        self.actualizando = True
        
        self.es_duplicado_actual = es_duplicado
        self.es_pf_actual = es_poste_igual_pf_existente
        
        if es_duplicado or es_poste_igual_pf_existente:
            self.btn_m_anterior.setVisible(False)
        
        if es_duplicado:
            self.cmb_tipo_poste.blockSignals(True)
            self.cmb_tipo_poste.setCurrentIndex(0)
            self.cmb_tipo_poste.blockSignals(False)
            self.cmb_tipo_poste.setEnabled(False)
            
            self.rb_cim_no.setChecked(True)
            self.validar_estado_cimentacion()
            
            self.cmb_tipo_linea.blockSignals(True)
            if linea_copiada in ["MEDIA TENSION", "MEDIA + BAJA TENSION", "BAJA TENSION"]:
                self.cmb_tipo_linea.setCurrentText(linea_copiada)
            else:
                self.cmb_tipo_linea.setCurrentIndex(0)
            self.cmb_tipo_linea.blockSignals(False)
            self.cmb_tipo_linea.setEnabled(False)
            
            self.cmb_linea_llega.blockSignals(True)
            self.cmb_linea_llega.setCurrentIndex(0)
            self.cmb_linea_llega.blockSignals(False)
            self.cmb_linea_llega.setEnabled(False)
            
            self.txt_metros.blockSignals(True)
            self.txt_metros.clear()
            self.txt_metros.blockSignals(False)
            self.txt_metros.setEnabled(False)
            
            self.controlar_tipo_bt()
        
        elif es_poste_igual_pf_existente:
            self.cmb_tipo_poste.blockSignals(True)
            self.cmb_tipo_poste.setCurrentIndex(0)
            self.cmb_tipo_poste.blockSignals(False)
            self.cmb_tipo_poste.setEnabled(False)
            
            self.rb_cim_no.setChecked(True)
            self.validar_estado_cimentacion()
            
            self.cmb_tipo_linea.blockSignals(True)
            if linea_copiada in ["MEDIA TENSION", "MEDIA + BAJA TENSION", "BAJA TENSION"]:
                self.cmb_tipo_linea.setCurrentText(linea_copiada)
            else:
                self.cmb_tipo_linea.setCurrentIndex(0)
            self.cmb_tipo_linea.blockSignals(False)
            self.cmb_tipo_linea.setEnabled(False)
            
            self.cmb_linea_llega.setEnabled(True)
            self.controlar_estado_metros()
        
        else:
            self.cmb_tipo_poste.setEnabled(True)
            self.validar_estado_cimentacion()
            self.cmb_tipo_linea.setEnabled(True) 
            if not self.cmb_tipo_linea.currentText():
                self.cmb_linea_llega.setEnabled(False)
            else:
                self.cmb_linea_llega.setEnabled(True)
            self.controlar_estado_metros()
            
        self.filtrar_materiales()
        self.actualizando = False

    def cambiar_voltaje(self):
        seleccion = self.cmb_potencia.currentText()
        if not seleccion:
            return

        if "7.6" in seleccion:
            viejo = "(19.9 kV)"
            nuevo = "(7.6 kV)"
        else:
            viejo = "(7.6 kV)"
            nuevo = "(19.9 kV)"

        if self.df is not None and not self.df.empty:
            for col in self.df.columns:
                if self.df[col].dtype == object:
                    self.df[col] = self.df[col].fillna("")
                    self.df[col] = self.df[col].astype(str).str.replace(viejo, nuevo, regex=False)
                    self.df[col] = self.df[col].replace(["nan", "None", "NaN"], "")

        global TIPOS_TRANSFORMADOR
        TIPOS_TRANSFORMADOR = [t.replace(viejo, nuevo) for t in TIPOS_TRANSFORMADOR]

        if self.cmb_material.currentText() == "TRANSFORMADORES":
            seleccion_actual = self.cmb_tipo_especial.currentText()
            self.cambiar_tipo_input()
            nueva_seleccion = seleccion_actual.replace(viejo, nuevo)
            self.cmb_tipo_especial.setCurrentText(nueva_seleccion)

        if self.modo_edicion and self.registro_actual_idx is not None:
            self.cargar_registro()

        if self.modo_edicion:
            self.btn_guardar.setText("Guardar Cambios (Pendientes)")
        self.hay_cambios_pendientes = True

    def mostrar_menu_excel(self):
        if self.hay_cambios_pendientes or self.registros_sin_guardar > 0 or self.cambios_catalogo:
            resp = QMessageBox.question(
                self,
                "Guardar Cambios",
                "¿Deseas guardar los cambios actuales antes de cambiar de archivo?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
            )
            if resp == QMessageBox.StandardButton.Yes:
                if self.hay_cambios_pendientes:
                    self.confirmar_cambios_dataframe()
                self.guardar_excel()
            elif resp == QMessageBox.StandardButton.Cancel:
                return

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Configuración de Archivo")
        msg_box.setText("¿Qué deseas hacer con el archivo Excel?")
        
        btn_abrir = msg_box.addButton("Abrir Existente", QMessageBox.ActionRole)
        btn_nuevo = msg_box.addButton("Empezar de 0", QMessageBox.ActionRole)
        btn_cancelar = msg_box.addButton("Cancelar", QMessageBox.RejectRole)
        
        msg_box.exec()
        
        if msg_box.clickedButton() == btn_abrir:
            self.cargar_excel()
        elif msg_box.clickedButton() == btn_nuevo:
            self.iniciar_desde_machote()

    def estandarizar_columnas(self, df):
        renames = {}
        for col in df.columns:
            col_u = str(col).upper().strip()
            if col_u == "PO": renames[col] = "Po"
            elif col_u == "PF": renames[col] = "Pf"
            elif col_u == "NO POSTE": renames[col] = "NO POSTE"
            elif col_u == "TIPO DE POSTE": renames[col] = "TIPO DE POSTE"
            elif col_u == "LINEA": renames[col] = "LINEA"
            elif col_u == "FASES": renames[col] = "FASES"
            else:
                for ec in COLUMNAS_LINEA + MATERIALES_COLUMNAS:
                    if col_u == ec.upper():
                        renames[col] = ec
                        break
        df.rename(columns=renames, inplace=True)
        todas_columnas = ["Po", "Pf", "NO POSTE", "TIPO DE POSTE", "LINEA", "FASES"] + COLUMNAS_LINEA + MATERIALES_COLUMNAS
        for col_nec in todas_columnas:
            if col_nec not in df.columns:
                df[col_nec] = "" if col_nec in MATERIALES_TEXTO else 0
        return df

    def iniciar_desde_machote(self):
        try:
            if getattr(sys, 'frozen', False):
                ruta_base = sys._MEIPASS
            else:
                ruta_base = os.path.dirname(os.path.abspath(__file__))
            ruta_machote = os.path.join(ruta_base, "Poste a Poste - Machote.xlsx")
            if not os.path.exists(ruta_machote):
                QMessageBox.critical(self, "Error", f"No se encontró el machote:\n{ruta_machote}")
                return
            wb = load_workbook(ruta_machote, data_only=True)
            for ws in wb.worksheets:
                for tbl in ws.tables.values():
                    if tbl.name == "P_P":
                        data = ws[tbl.ref]
                        rows = [[c.value for c in r] for r in data]
                        self.df = pd.DataFrame(columns=rows[0])
                        self.df = self.estandarizar_columnas(self.df)
                        self.df["_hijo_vii"] = False
                        
                        self.usando_machote = True
                        self.archivo_excel = None
                        self.nombre_hoja = ws.title
                        self.ultimo_registro_creado = {}
                        self.limpiar_campos()
                        self.cmb_potencia.setCurrentIndex(0)
                        QMessageBox.information(self, "Nuevo Documento", "Documento nuevo iniciado.")
                        wb.close()
                        self.establecer_modo_nuevo()
                        return
            wb.close()
            QMessageBox.critical(self, "Error", "No se encontró la tabla P_P")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error cargando machote: {str(e)}")

    def validar_metros(self):
        if self.actualizando: return
        texto = self.txt_metros.text()
        solo_numeros = ''.join(filter(str.isdigit, texto))
        if texto != solo_numeros:
            self.txt_metros.blockSignals(True)
            self.txt_metros.setText(solo_numeros)
            self.txt_metros.blockSignals(False)

    def validar_cantidad(self):
        if self.actualizando: return
        texto = self.txt_cantidad.text()
        solo_numeros = ''.join(filter(str.isdigit, texto))
        if texto != solo_numeros:
            self.txt_cantidad.blockSignals(True)
            self.txt_cantidad.setText(solo_numeros)
            self.txt_cantidad.blockSignals(False)

    def cambiar_clasificacion_poste(self, clasificacion):
        self.btn_concreto.setChecked(False)
        self.btn_madera.setChecked(False)
        self.btn_metal.setChecked(False)
        self.btn_fibra.setChecked(False)
        if clasificacion == "Concreto":
            self.btn_concreto.setChecked(True); tipos = TIPOS_POSTE_CONCRETO
        elif clasificacion == "Madera":
            self.btn_madera.setChecked(True); tipos = TIPOS_POSTE_MADERA
        elif clasificacion == "Metal":
            self.btn_metal.setChecked(True); tipos = TIPOS_POSTE_METAL
        elif clasificacion == "Fibra":
            self.btn_fibra.setChecked(True); tipos = TIPOS_POSTE_FIBRA
        else: return
        valor_actual = self.cmb_tipo_poste.currentText()
        self.cmb_tipo_poste.clear()
        self.cmb_tipo_poste.addItems(["", *tipos])
        if valor_actual in tipos:
            self.cmb_tipo_poste.setCurrentText(valor_actual)

    def borrar_registro(self):
        if not self.modo_edicion or self.registro_actual_idx is None:
            QMessageBox.warning(self, "Advertencia", "No hay registro seleccionado")
            return
            
        row = self.df.loc[self.registro_actual_idx]
        if row.get("_hijo_vii", False):
            return 
            
        if QMessageBox.question(self, "Confirmar", f"¿Eliminar registro {self.registro_actual_idx}?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            try:
                idx = self.cmb_registro.currentIndex()
                p, po, pf = row["NO POSTE"], row["Po"], row["Pf"]
                
                mask_hijo = (self.df["NO POSTE"] == p) & (self.df["Po"] == po) & (self.df["Pf"] == pf) & (self.df.get("_hijo_vii", False) == True)
                if mask_hijo.any():
                    self.df = self.df[~mask_hijo]
                
                self.df = self.df.drop(self.registro_actual_idx).reset_index(drop=True)
                self.actualizar_postes_unicos()
                self.buscar_por_poste()
                if self.cmb_registro.count() > 0:
                    self.cmb_registro.setCurrentIndex(max(0, idx - 1))
                else:
                    self.ir_poste_anterior() if self.poste_actual_index > 0 else self.limpiar_campos()
                QMessageBox.information(self, "Éxito", "Eliminado correctamente")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error: {str(e)}")

    def marcar_cambio(self):
        if not self.modo_edicion or self.actualizando or self.registro_actual_idx is None: return
        if self.df.loc[self.registro_actual_idx].get("_hijo_vii", False): return
        
        if not self.hay_cambios_pendientes:
            self.hay_cambios_pendientes = True
            self.txt_poste.setEnabled(False)
            self.cmb_registro.setEnabled(False)
            self.btn_prev.setEnabled(False)
            self.btn_next.setEnabled(False)
            self.btn_ultimo.setEnabled(False)
            self.btn_guardar.setText("Guardar Cambios")

    def _falta_remate_obligatorio(self):
        """
        True solo si el registro actual DEBE llevar un remate (doble/simple) y no lo tiene.
        No se exige remate cuando:
          - la línea es solo MEDIA TENSION,
          - el remate ya está en la lista de materiales del registro actual,
          - es un registro repetido (duplicado / mismo Pf), o
          - OTRO registro del mismo No. Poste ya tiene el remate (en repeticiones solo
            cambia la cantidad de metros, la estructura o la fase).
        """
        if self.cmb_tipo_linea.currentText() not in ["MEDIA + BAJA TENSION", "BAJA TENSION"]:
            return False
        if getattr(self, 'es_duplicado_actual', False) or getattr(self, 'es_pf_actual', False):
            return False

        for i in range(self.lst_materiales.count()):
            t = self.lst_materiales.item(i).text()
            if "ESTRUCTURA REMATE DOBLE" in t or "ESTRUCTURA REMATE SIMPLE" in t:
                return False

        try:
            if self.df is not None and not self.df.empty:
                p = self.formatear_valor_entero(self.txt_poste.text().strip())
                excluir = self.registro_actual_idx if self.modo_edicion else None
                for j, fila in self.df.iterrows():
                    if excluir is not None and j == excluir:
                        continue
                    if fila.get("_hijo_vii", False) == True:
                        continue
                    if self.formatear_valor_entero(fila.get("NO POSTE", "")) != p:
                        continue
                    for col in ("ESTRUCTURA REMATE DOBLE", "ESTRUCTURA REMATE SIMPLE"):
                        if col in self.df.columns:
                            v = fila.get(col, 0)
                            if pd.notna(v) and self._tiene_valor(v):
                                return False
        except Exception:
            pass
        return True

    def _campos_obligatorios_faltantes(self):
        faltan = []
        if not self.txt_poste.text().strip(): faltan.append("No. Poste")
        if not self.txt_po.text().strip(): faltan.append("Po")
        if not self.txt_pf.text().strip(): faltan.append("Pf")
        if not self.cmb_tipo_linea.currentText().strip(): faltan.append("Tipo de Línea")
        return faltan

    def confirmar_cambios_dataframe(self):
        if self._falta_remate_obligatorio():
            QMessageBox.warning(self, "Advertencia", "Falta agregar ESTRUCTURA REMATE DOBLE o ESTRUCTURA REMATE SIMPLE.")
            return

        faltan = self._campos_obligatorios_faltantes()
        if faltan:
            QMessageBox.warning(self, "Advertencia", "Faltan campos obligatorios:\n- " + "\n- ".join(faltan))
            return

        self.actualizar_dataframe_completo()
        self.hay_cambios_pendientes = False
        self.txt_poste.setEnabled(True)
        self.cmb_registro.setEnabled(True)
        self.btn_prev.setEnabled(True)
        self.btn_next.setEnabled(True)
        self.btn_ultimo.setEnabled(self.df is not None and len(self.df) > 2)
        self.btn_guardar.setText("Guardar al Excel")
        if hasattr(self, 'actualizar_dibujo'):
            self.actualizar_dibujo()

    def click_editar(self):
        self.establecer_modo_editar()
        self.ir_a_ultimo()

    def establecer_modo_editar(self):
        self.modo_edicion = True
        self.detalles_group.setEnabled(True)
        self.material_group.setEnabled(True)
        
        self.btn_editar.setEnabled(False)
        self.btn_nuevo.setEnabled(True)
        self.btn_aceptar.setEnabled(False)
        self.btn_borrar_registro.setEnabled(True)
        self.btn_guardar.setEnabled(True)
        self.cmb_potencia.setVisible(True)
        mostrar_ultimo = self.df is not None and len(self.df) > 2
        self.btn_ultimo.setVisible(mostrar_ultimo)
        self.btn_ultimo.setEnabled(mostrar_ultimo)
        self.btn_m_anterior.setVisible(False)
        if not self.hay_cambios_pendientes:
            self.txt_poste.setEnabled(True)
            self.cmb_registro.setEnabled(True)
            self.btn_prev.setEnabled(True)
            self.btn_next.setEnabled(True)

    def establecer_modo_nuevo(self):
        self.modo_edicion = False
        self.detalles_group.setEnabled(True)
        self.material_group.setEnabled(True)
        
        if self.df is not None and not self.df.empty:
            self.btn_editar.setEnabled(True)
        else:
            self.btn_editar.setEnabled(False)
            
        self.btn_nuevo.setEnabled(False)
        self.btn_aceptar.setEnabled(True)
        self.btn_borrar_registro.setEnabled(False)
        self.btn_guardar.setEnabled(False)
        self.cmb_registro.setEnabled(False)
        self.btn_prev.setEnabled(False)
        self.btn_next.setEnabled(False)
        self.btn_ultimo.setVisible(False)
        self.cmb_potencia.setVisible(True)
        
        self.preparar_nuevo_registro()
        
        if self.df is not None and not self.df.empty:
            self.btn_m_anterior.setVisible(True)
            self.btn_m_anterior.setEnabled(True)
        else:
            self.btn_m_anterior.setVisible(False)
        
        df_vacio = self.df is None or self.df.empty
        if df_vacio or not self.cmb_potencia.currentText().strip():
            QTimer.singleShot(100, lambda: (self.cmb_potencia.setFocus(), self.cmb_potencia.showPopup()))
        else:
            QTimer.singleShot(100, lambda: self.txt_poste.setFocus())

    def preparar_nuevo_registro(self):
        self.actualizando = True
        
        self.cmb_tipo_poste.setEnabled(True)
        self.cmb_tipo_linea.setEnabled(True)
        
        self.es_hijo_vii_actual = False
        self.es_duplicado_actual = False
        self.es_pf_actual = False
        
        if self.ultimo_registro_creado:
            try:
                up = int(self.ultimo_registro_creado.get("NO POSTE", "0"))
                self.txt_poste.setText(str(up + 1)); self.txt_po.setText(str(up)); self.txt_pf.setText(str(up + 1))
            except: self.limpiar_campos()
            
            tipo_poste = self.ultimo_registro_creado.get("TIPO DE POSTE", "")
            self.cmb_tipo_poste.setCurrentText(tipo_poste) if tipo_poste else self.cmb_tipo_poste.setCurrentIndex(0)
            self.btn_existente.setChecked(bool(self.ultimo_registro_creado.get("ES_EXISTENTE", False)))
            
            self.rb_cim_si.setChecked(self.ultimo_registro_creado.get("CIMENTACION", "") != "")
            self.rb_cim_no.setChecked(self.ultimo_registro_creado.get("CIMENTACION", "") == "")
            self.validar_estado_cimentacion()

            tl = self.ultimo_registro_creado.get("TIPO_LINEA", "")
            self.cmb_tipo_linea.setCurrentText(tl) if tl in ["MEDIA TENSION", "MEDIA + BAJA TENSION", "BAJA TENSION"] else self.cmb_tipo_linea.setCurrentIndex(0)
            self.filtrar_linea_llega()
            
            ll = self.ultimo_registro_creado.get("LINEA_LLEGA", "")
            self.cmb_linea_llega.setCurrentText(ll) if ll else self.cmb_linea_llega.setCurrentIndex(0)
            
            self.txt_metros.setText(self.ultimo_registro_creado.get("metros", ""))
            
            f = self.ultimo_registro_creado.get("FASES", "")
            self.rb_1f.setChecked(f=="1F"); self.rb_2f.setChecked(f=="2F"); self.rb_3f.setChecked(f=="3F")
            
            tbt = self.ultimo_registro_creado.get("TIPO_BT", "")
            if ll in ["BAJA TENSION", "MEDIA + BAJA TENSION"]:
                self.rb_bt_1_0.setChecked(tbt=="1/0"); self.rb_bt_2_0.setChecked(tbt=="2/0")
            
        else: 
            self.limpiar_campos()
        
        self.lst_materiales.clear()
        self.cmb_material.setCurrentIndex(0)
        self.txt_cantidad.clear()
        self.actualizando = False
        
        self.validar_reglas_ui()

        if not getattr(self, 'modo_edicion', True):
            if self.df is not None and not self.df.empty:
                self.btn_m_anterior.setVisible(True)
                self.btn_m_anterior.setEnabled(True)
            else:
                self.btn_m_anterior.setVisible(False)

        if hasattr(self, 'actualizar_dibujo'):
            self.actualizar_dibujo()

    def limpiar_campos(self):
        self.actualizando = True
        
        self.cmb_tipo_poste.setEnabled(True)
        self.cmb_tipo_linea.setEnabled(True) 
        self.cmb_linea_llega.setEnabled(True)
        
        self.es_hijo_vii_actual = False
        self.es_duplicado_actual = False
        self.es_pf_actual = False
        
        self.txt_poste.clear(); self.txt_po.clear(); self.txt_pf.clear()
        self.cmb_tipo_poste.setCurrentIndex(0)
        self.rb_cim_no.setChecked(True)
        self.validar_estado_cimentacion()
        
        self.cmb_tipo_linea.setCurrentIndex(0)
        self.filtrar_linea_llega()
        self.cmb_linea_llega.setCurrentIndex(0)
        
        self.controlar_estado_metros()
        
        self.rb_1f.setChecked(True)
        
        self.tipo_bt_group.setExclusive(False)
        self.rb_bt_1_0.setChecked(False); self.rb_bt_2_0.setChecked(False)
        self.tipo_bt_group.setExclusive(True)
        
        self.lst_materiales.clear()
        self.cmb_material.setCurrentIndex(0); self.txt_cantidad.clear()
        self.btn_m_anterior.setVisible(False)
        self.actualizando = False

        if hasattr(self, 'actualizar_dibujo'):
            self.actualizar_dibujo()

    def cargar_materiales_anterior(self):
        if self.df is None or self.df.empty:
            return

        idx_ultimo = len(self.df) - 1
        row_anterior = self.df.iloc[idx_ultimo]
            
        self.lst_materiales.clear()
        
        for col in MATERIALES_COLUMNAS:
            if col in self.df.columns and col not in ["CIMENTACION"]:
                val = row_anterior.get(col, 0)
                if pd.notna(val) and self._tiene_valor(val):
                    vf = self.formatear_valor_entero(val) if col not in MATERIALES_TEXTO else val
                    self.lst_materiales.addItem(f"{col} – {vf}")
        
        self.btn_m_anterior.setVisible(False)
        self.marcar_cambio()

        if hasattr(self, 'actualizar_dibujo'):
            self.actualizar_dibujo()    

    def actualizar_dataframe_completo(self):
        if not self.modo_edicion or self.registro_actual_idx is None: return
        
        def set_df_value(idx, col, value):
            try:
                self.df.at[idx, col] = value
            except (ValueError, TypeError):
                self.df[col] = self.df[col].astype(object)
                self.df.at[idx, col] = value

        try:
            idx = self.registro_actual_idx
            cols = ["Po", "Pf", "NO POSTE", "TIPO DE POSTE", "LINEA", "FASES"]
            for col in cols: 
                if col in self.df.columns and self.df[col].dtype != 'object': self.df[col] = self.df[col].astype(object)

            set_df_value(idx, "Po", self.txt_po.text().strip())
            set_df_value(idx, "Pf", self.txt_pf.text().strip())
            set_df_value(idx, "TIPO DE POSTE", self._tipo_poste_guardado())
            
            if "CIMENTACION" not in self.df.columns:
                self.df["CIMENTACION"] = ""
                self.df["CIMENTACION"] = self.df["CIMENTACION"].astype(object)
            
            cim_str = "C " + self._tipo_poste_guardado() if self.rb_cim_si.isChecked() else ""
            set_df_value(idx, "CIMENTACION", cim_str)

            set_df_value(idx, "LINEA", self.cmb_tipo_linea.currentText())
            fases_str = "1F" if self.rb_1f.isChecked() else ("2F" if self.rb_2f.isChecked() else ("3F" if self.rb_3f.isChecked() else ""))
            set_df_value(idx, "FASES", fases_str)

            for col in COLUMNAS_LINEA:
                if col in self.df.columns: set_df_value(idx, col, 0)

            linea = self.cmb_linea_llega.currentText()
            mts = self.txt_metros.text().strip()
            if linea and mts:
                col_dest = linea
                if linea in ["BAJA TENSION", "MEDIA + BAJA TENSION"]:
                    col_dest += " 1/0" if self.rb_bt_1_0.isChecked() else " 2/0"
                
                if col_dest not in self.df.columns:
                    self.df[col_dest] = 0
                
                try: set_df_value(idx, col_dest, int(mts))
                except: set_df_value(idx, col_dest, mts)

            for col in MATERIALES_COLUMNAS:
                if col == "CIMENTACION": continue 
                if col in self.df.columns:
                    val_reset = "" if col in MATERIALES_TEXTO else 0
                    set_df_value(idx, col, val_reset)

            has_vii = False
            for i in range(self.lst_materiales.count()):
                item = self.lst_materiales.item(i).text()
                parts = item.split(" – ")
                
                if len(parts) == 2:
                    if parts[0].strip() == "ESTRUCTURA" and parts[1].strip() == "VII":
                        has_vii = True
                        
                if "CIMENTACION" in item: continue
                
                if len(parts) == 2:
                    mat = parts[0].strip()
                    val = parts[1].strip()
                    if mat not in self.df.columns:
                        self.df[mat] = 0 
                    if mat in MATERIALES_TEXTO:
                        set_df_value(idx, mat, val)
                    else:
                        try: set_df_value(idx, mat, int(val))
                        except: set_df_value(idx, mat, val)
            
            p, po, pf = str(self.df.at[idx, "NO POSTE"]), str(self.df.at[idx, "Po"]), str(self.df.at[idx, "Pf"])
            mask_hijo = (self.df["NO POSTE"].astype(str) == p) & (self.df["Po"].astype(str) == po) & (self.df["Pf"].astype(str) == pf) & (self.df.get("_hijo_vii", False) == True)
            
            if has_vii:
                if not mask_hijo.any():
                    child_reg = {col: ("" if col in MATERIALES_TEXTO else 0) for col in self.df.columns if col != "_hijo_vii"}
                    child_reg["NO POSTE"] = self.df.at[idx, "NO POSTE"]
                    child_reg["Po"] = self.df.at[idx, "Po"]
                    child_reg["Pf"] = self.df.at[idx, "Pf"]
                    child_reg["TIPO DE POSTE"] = self.df.at[idx, "TIPO DE POSTE"]
                    child_reg["CIMENTACION"] = self.df.at[idx, "CIMENTACION"]
                    child_reg["LINEA"] = self.df.at[idx, "LINEA"]
                    child_reg["FASES"] = self.df.at[idx, "FASES"]
                    child_reg["_hijo_vii"] = True
                    self.df = pd.concat([self.df, pd.DataFrame([child_reg])], ignore_index=True)
                else:
                    child_idx = self.df[mask_hijo].index[0]
                    self.df.at[child_idx, "TIPO DE POSTE"] = self.df.at[idx, "TIPO DE POSTE"]
                    self.df.at[child_idx, "CIMENTACION"] = self.df.at[idx, "CIMENTACION"]
                    self.df.at[child_idx, "LINEA"] = self.df.at[idx, "LINEA"]
                    self.df.at[child_idx, "FASES"] = self.df.at[idx, "FASES"]
            else:
                if mask_hijo.any():
                    self.df = self.df[~mask_hijo].reset_index(drop=True)
            
        except Exception as e: QMessageBox.critical(self, "Error", f"Error al actualizar: {str(e)}")

    def aceptar_nuevo(self):
        if self.df is None: return QMessageBox.warning(self, "Advertencia", "Carga un Excel")
        if not self.txt_poste.text().strip(): return QMessageBox.warning(self, "Advertencia", "Falta No. Poste")

        if self._falta_remate_obligatorio():
            QMessageBox.warning(self, "Advertencia", "Falta agregar ESTRUCTURA REMATE DOBLE o ESTRUCTURA REMATE SIMPLE.")
            return

        if not self.cmb_potencia.currentText().strip():
            QMessageBox.warning(self, "Advertencia", "Seleccione un Voltaje antes de guardar.")
            return

        faltan = self._campos_obligatorios_faltantes()
        if faltan:
            QMessageBox.warning(self, "Advertencia", "Faltan campos obligatorios:\n- " + "\n- ".join(faltan))
            return

        reg = {}
        for col in self.df.columns:
            if col in MATERIALES_TEXTO: reg[col] = ""
            elif col in COLUMNAS_LINEA or col in MATERIALES_COLUMNAS: reg[col] = 0
            else: reg[col] = ""

        reg["NO POSTE"] = self.txt_poste.text().strip()
        reg["Po"] = self.txt_po.text().strip()
        reg["Pf"] = self.txt_pf.text().strip()
        reg["TIPO DE POSTE"] = self._tipo_poste_guardado()
        
        cim_str = "C " + self._tipo_poste_guardado() if self.rb_cim_si.isChecked() else ""
        reg["CIMENTACION"] = cim_str

        reg["LINEA"] = self.cmb_tipo_linea.currentText()
        fases_str = "1F" if self.rb_1f.isChecked() else ("2F" if self.rb_2f.isChecked() else ("3F" if self.rb_3f.isChecked() else ""))
        reg["FASES"] = fases_str

        if self.cmb_linea_llega.currentText() and self.txt_metros.text().strip():
            l = self.cmb_linea_llega.currentText()
            cd = l
            if l in ["BAJA TENSION", "MEDIA + BAJA TENSION"]: cd += " 1/0" if self.rb_bt_1_0.isChecked() else " 2/0"
            try: reg[cd] = int(self.txt_metros.text().strip())
            except: pass

        voltaje = self.cmb_potencia.currentText()
        if voltaje and ("7.6" in voltaje or "19.9" in voltaje):
            if "7.6" in voltaje:
                viejo_v = "(19.9 kV)"; nuevo_v = "(7.6 kV)"
            else:
                viejo_v = "(7.6 kV)"; nuevo_v = "(19.9 kV)"
            for i in range(self.lst_materiales.count()):
                item = self.lst_materiales.item(i)
                texto = item.text()
                if texto.startswith("TRANSFORMADORES"):
                    item.setText(texto.replace(viejo_v, nuevo_v))

        mats = []
        has_vii = False
        for i in range(self.lst_materiales.count()):
            item = self.lst_materiales.item(i).text()
            parts = item.split(" – ")
                
            if len(parts) == 2:
                 if parts[0].strip() == "ESTRUCTURA" and parts[1].strip() == "VII":
                    has_vii = True
                        
            if "CIMENTACION" in item: continue
                
            if len(parts) == 2:
                m = parts[0].strip()
                v = parts[1].strip()
                if m in reg:
                    if m in MATERIALES_TEXTO: reg[m] = v
                    else:
                        try: reg[m] = int(v)
                        except: reg[m] = v

        tbt = ""
        if self.cmb_linea_llega.currentText() in ["BAJA TENSION", "MEDIA + BAJA TENSION"]:
            tbt = "1/0" if self.rb_bt_1_0.isChecked() else "2/0"

        self.ultimo_registro_creado = {
            "NO POSTE": self.txt_poste.text().strip(),
            "Po": self.txt_po.text().strip(),
            "Pf": self.txt_pf.text().strip(),
            "TIPO DE POSTE": self.cmb_tipo_poste.currentText(),
            "ES_EXISTENTE": self.btn_existente.isChecked(),
            "CIMENTACION": reg["CIMENTACION"],
            "TIPO_LINEA": self.cmb_tipo_linea.currentText(),
            "LINEA_LLEGA": self.cmb_linea_llega.currentText(),
            "FASES": reg["FASES"],
            "metros": self.txt_metros.text().strip(),
            "TIPO_BT": tbt,
            "materiales": mats 
        }

        self.df = pd.concat([self.df, pd.DataFrame([reg])], ignore_index=True)
        
        if has_vii:
            child_reg = {c: ("" if c in MATERIALES_TEXTO else 0) for c in self.df.columns}
            child_reg["NO POSTE"] = reg["NO POSTE"]
            child_reg["Po"] = reg["Po"]
            child_reg["Pf"] = reg["Pf"]
            child_reg["TIPO DE POSTE"] = reg["TIPO DE POSTE"]
            child_reg["CIMENTACION"] = reg["CIMENTACION"]
            child_reg["LINEA"] = reg["LINEA"]
            child_reg["FASES"] = reg["FASES"]
            child_reg["_hijo_vii"] = True
            self.df = pd.concat([self.df, pd.DataFrame([child_reg])], ignore_index=True)

        self.actualizar_postes_unicos()
        self.registros_sin_guardar += 1
        
        self.btn_editar.setEnabled(True) 
        
        QMessageBox.information(self, "Éxito", "Agregado")
        
        if self.registros_sin_guardar >= 20:
            if QMessageBox.question(self, "Autoguardado", "20 nuevos. ¿Guardar?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes: self.guardar_excel(mantener_modo_nuevo=True)
        
        self.preparar_nuevo_registro()
        if hasattr(self, 'actualizar_dibujo'):
            self.actualizar_dibujo()
        QTimer.singleShot(100, lambda: (self.txt_metros.setEnabled(True), self.txt_metros.setFocus(), self.txt_metros.selectAll()))

    def cargar_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Abrir", "", "Excel (*.xlsx *.xlsm)")
        if not path: return
        try:
            self.cmb_potencia.blockSignals(True)
            self.cmb_potencia.setCurrentText("Linea 19.9 kV")
            self.cmb_potencia.blockSignals(False)
            
            global TIPOS_TRANSFORMADOR
            TIPOS_TRANSFORMADOR = [t.replace("(7.6 kV)", "(19.9 kV)") for t in TIPOS_TRANSFORMADOR]

            wb = load_workbook(path, data_only=True)
            df_cargado = None
            hoja_cargada = None
            
            for ws in wb.worksheets:
                for tbl in ws.tables.values():
                    if tbl.name == "P_P":
                        data = ws[tbl.ref]
                        rows = [[c.value for c in r] for r in data]
                        df_cargado = pd.DataFrame(rows[1:], columns=rows[0])
                        
                        for col in df_cargado.columns:
                            col_u = str(col).upper().strip()
                            if col_u not in ["PO", "PF", "NO POSTE", "TIPO DE POSTE", "LINEA", "FASES", "_HIJO_VII"] and col_u not in COLUMNAS_LINEA and col_u not in MATERIALES_COLUMNAS:
                                MATERIALES_COLUMNAS.append(col_u)
                                if df_cargado[col].dtype == object:
                                    MATERIALES_TEXTO.append(col_u)
                        
                        df_cargado = self.estandarizar_columnas(df_cargado)
                        hoja_cargada = ws.title
                        break
                if df_cargado is not None: break
            wb.close()

            if df_cargado is None:
                QMessageBox.critical(self, "Error", "No se encontró la tabla P_P en el archivo.")
                return

            self.df = df_cargado
            self.df["_hijo_vii"] = False
            self.archivo_excel = path
            self.nombre_hoja = hoja_cargada
            self.usando_machote = False
            
            if "TRANSFORMADORES" in self.df.columns:
                for val in self.df["TRANSFORMADORES"].dropna():
                    if "7.6 kV" in str(val):
                        self.cmb_potencia.blockSignals(True)
                        self.cmb_potencia.setCurrentText("Linea 7.6 kV")
                        self.cmb_potencia.blockSignals(False)
                        TIPOS_TRANSFORMADOR = [t.replace("(19.9 kV)", "(7.6 kV)") for t in TIPOS_TRANSFORMADOR]
                        break
            
            for i in range(len(self.df)):
                if str(self.df.at[i, "ESTRUCTURA"]).strip().upper() == "VII":
                    p, po, pf = str(self.df.at[i, "NO POSTE"]).strip(), str(self.df.at[i, "Po"]).strip(), str(self.df.at[i, "Pf"]).strip()
                    for j in range(i+1, len(self.df)):
                        if str(self.df.at[j, "NO POSTE"]).strip() == p and str(self.df.at[j, "Po"]).strip() == po and str(self.df.at[j, "Pf"]).strip() == pf:
                            has_mats = False
                            for m in MATERIALES_COLUMNAS:
                                if m != "CIMENTACION":
                                    val = self.df.at[j, m]
                                    if pd.notna(val) and str(val).strip().lower() not in ["", "0", "nan", "none", "0.0"]:
                                        has_mats = True
                                        break
                            if not has_mats:
                                self.df.at[j, "_hijo_vii"] = True
                                break

            self.cargar_tipos_poste()
            self.actualizar_postes_unicos()
            self.cargar_catalogos()
            self.registros_sin_guardar = 0

            if not self.df.empty:
                self.actualizando = True
                self.establecer_modo_editar()
                
                if self.postes_unicos: 
                    self.txt_poste.setText(self.formatear_valor_entero(self.postes_unicos[0]))
                    self.buscar_por_poste()
                    if self.cmb_registro.count() > 0: 
                        self.cmb_registro.setCurrentIndex(0)
                        
                if len(self.df) > 2: 
                    self.btn_ultimo.setVisible(True)
                    self.btn_ultimo.setEnabled(True)
                
                self.actualizando = False
                self.cargar_registro()
                
                self.hay_cambios_pendientes = False
                self.btn_guardar.setText("Guardar al Excel")
            else: 
                self.btn_editar.setEnabled(False)
                self.establecer_modo_nuevo()
            
            QMessageBox.information(self, "OK", "Tabla P_P cargada correctamente.")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error cargando archivo: {str(e)}")

    def cargar_tipos_poste(self):
        if self.df is not None and "TIPO DE POSTE" in self.df.columns:
            tu = self.df["TIPO DE POSTE"].dropna().unique().tolist()
            
            global TIPOS_POSTE, TIPOS_POSTE_CONCRETO, TIPOS_POSTE_MADERA, TIPOS_POSTE_METAL, TIPOS_POSTE_FIBRA
            
            for p in tu:
                p_str = str(p).strip()
                if p_str.startswith("E "):
                    p_str = p_str[2:]
                if p_str and p_str not in TIPOS_POSTE:
                    TIPOS_POSTE.append(p_str)
                    
                    pl = p_str.lower()
                    if "madera" in pl:
                        TIPOS_POSTE_MADERA.append(p_str)
                    elif "metal" in pl:
                        TIPOS_POSTE_METAL.append(p_str)
                    elif "fibr" in pl:
                        TIPOS_POSTE_FIBRA.append(p_str)
                    else:
                        TIPOS_POSTE_CONCRETO.append(p_str)

            ordenar_tipos_poste(TIPOS_POSTE_CONCRETO)
            ordenar_tipos_poste(TIPOS_POSTE_MADERA)
            ordenar_tipos_poste(TIPOS_POSTE_METAL)
            ordenar_tipos_poste(TIPOS_POSTE_FIBRA)
            ordenar_tipos_poste(TIPOS_POSTE)

    def cargar_catalogos(self):
        global TIPOS_ESTRUCTURA, TIPOS_TRANSFORMADOR
        if self.df is None:
            return
        if "ESTRUCTURA" in self.df.columns:
            for val in self.df["ESTRUCTURA"].dropna().unique():
                v = str(val).strip()
                if v and v.lower() not in ("nan", "") and v not in TIPOS_ESTRUCTURA:
                    TIPOS_ESTRUCTURA.append(v)
                    TIPOS_ESTRUCTURA_CREADOS.add(v)
        if "TRANSFORMADORES" in self.df.columns:
            for val in self.df["TRANSFORMADORES"].dropna().unique():
                v = str(val).strip()
                if v and v.lower() not in ("nan", "") and v not in TIPOS_TRANSFORMADOR:
                    TIPOS_TRANSFORMADOR.append(v)
                    TIPOS_TRANSFORMADOR_CREADOS.add(v)

    def actualizar_postes_unicos(self):
        if self.df is not None and "NO POSTE" in self.df.columns:
            raw = self.df["NO POSTE"].dropna().unique().tolist()
            try: self.postes_unicos = sorted(raw, key=lambda x: float(x) if str(x).replace('.', '', 1).isdigit() else str(x))
            except: self.postes_unicos = sorted(raw, key=lambda x: str(x))

    def buscar_por_poste(self):
        if self.df is None or not self.modo_edicion: return
        p = self.txt_poste.text().strip()
        if not p: self.cmb_registro.clear(); self.registros_filtrados = pd.DataFrame(); return
        
        self.registros_filtrados = self.df[self.df["NO POSTE"].apply(lambda x: self.formatear_valor_entero(x)) == p].copy()
        
        self.cmb_registro.clear()
        if self.registros_filtrados.empty: self.poste_actual_index = -1; return
        try:
            for idx, po in enumerate(self.postes_unicos):
                if self.formatear_valor_entero(po) == p: self.poste_actual_index = idx; break
        except: self.poste_actual_index = -1
        for i in self.registros_filtrados.index: self.cmb_registro.addItem(f"Registro {i}", i)
        if self.cmb_registro.count() > 0: self.cmb_registro.setCurrentIndex(0)

    def cargar_registro(self):
        if self.registros_filtrados.empty: return
        idx = self.cmb_registro.currentData()
        if idx is None: return
        self.actualizando = True
        self.registro_actual_idx = idx
        row = self.df.loc[idx]

        curr_p = self.formatear_valor_entero(row.get("NO POSTE", ""))
        curr_po = self.formatear_valor_entero(row.get("Po", ""))
        curr_pf = self.formatear_valor_entero(row.get("Pf", ""))
        
        mask_normales = self.df.get("_hijo_vii", False) == False
        
        es_secundario_duplicado = False
        es_secundario_poste_pf = False
        linea_copiada = ""
        
        try:
            matches_dup = self.df[
                mask_normales &
                (self.df["NO POSTE"].apply(lambda x: self.formatear_valor_entero(x)) == curr_p) & 
                (self.df["Po"].apply(lambda x: self.formatear_valor_entero(x)) == curr_po) & 
                (self.df["Pf"].apply(lambda x: self.formatear_valor_entero(x)) == curr_pf)
            ].index.tolist()
            
            matches_dup = [i for i in matches_dup if i < idx]
            if matches_dup and not row.get("_hijo_vii", False):
                es_secundario_duplicado = True
                linea_copiada = str(self.df.at[matches_dup[0], "LINEA"]).replace("nan", "")
        except: pass

        if not es_secundario_duplicado:
            try:
                matches_pf = self.df[
                    mask_normales &
                    (self.df["NO POSTE"].apply(lambda x: self.formatear_valor_entero(x)) == curr_p) & 
                    (self.df["Pf"].apply(lambda x: self.formatear_valor_entero(x)) == curr_pf)
                ].index.tolist()
                
                matches_pf = [i for i in matches_pf if i < idx]
                if matches_pf and not row.get("_hijo_vii", False):
                    es_secundario_poste_pf = True
                    linea_copiada = str(self.df.at[matches_pf[0], "LINEA"]).replace("nan", "")
            except: pass

        self.es_duplicado_actual = es_secundario_duplicado
        self.es_pf_actual = es_secundario_poste_pf
        self.es_hijo_vii_actual = row.get("_hijo_vii", False)

        self.txt_po.setText(curr_po)
        self.txt_pf.setText(curr_pf)
        
        if es_secundario_duplicado or es_secundario_poste_pf:
            self.cmb_tipo_poste.setCurrentIndex(0)
            self.rb_cim_no.setChecked(True)
            if linea_copiada in ["MEDIA TENSION", "MEDIA + BAJA TENSION", "BAJA TENSION"]:
                self.cmb_tipo_linea.setCurrentText(linea_copiada)
            else:
                self.cmb_tipo_linea.setCurrentIndex(0)
        else:
            tp = str(row.get("TIPO DE POSTE", "")).replace("nan", "")
            es_existente = tp.startswith("E ")
            self.btn_existente.setChecked(es_existente)
            tp_sin_e = tp[2:] if es_existente else tp
            if tp_sin_e in TIPOS_POSTE_CONCRETO: self.cambiar_clasificacion_poste("Concreto")
            elif tp_sin_e in TIPOS_POSTE_MADERA: self.cambiar_clasificacion_poste("Madera")
            elif tp_sin_e in TIPOS_POSTE_METAL: self.cambiar_clasificacion_poste("Metal")
            elif tp_sin_e in TIPOS_POSTE_FIBRA: self.cambiar_clasificacion_poste("Fibra")
            
            if tp_sin_e:
                self.cmb_tipo_poste.setCurrentText(tp_sin_e)

            cim = str(row.get("CIMENTACION", "")).replace("nan", "").strip()
            if cim and cim.lower() not in ["0", "0.0", "none", "false", "nan"]:
                self.rb_cim_si.setChecked(True)
            else:
                self.rb_cim_no.setChecked(True)

            tl = str(row.get("LINEA", "")).replace("nan", "")
            self.cmb_tipo_linea.setCurrentText(tl) if tl in ["MEDIA TENSION", "MEDIA + BAJA TENSION", "BAJA TENSION"] else self.cmb_tipo_linea.setCurrentIndex(0)

        self.filtrar_linea_llega()

        if es_secundario_duplicado:
            self.cmb_linea_llega.setCurrentIndex(0)
            self.txt_metros.clear()
            self.tipo_bt_group.setExclusive(False)
            self.rb_bt_1_0.setChecked(False)
            self.rb_bt_2_0.setChecked(False)
            self.tipo_bt_group.setExclusive(True)
        else:
            self.cmb_linea_llega.setCurrentText("")
            self.txt_metros.clear()
            self.tipo_bt_group.setExclusive(False)
            self.rb_bt_1_0.setChecked(False)
            self.rb_bt_2_0.setChecked(False)
            self.tipo_bt_group.setExclusive(True)

            for col in COLUMNAS_LINEA:
                if col in row.index:
                    val = row.get(col, 0)
                    if pd.notna(val) and self._tiene_valor(val):
                        if col in ["BAJA TENSION 1/0", "BAJA TENSION 2/0"]:
                            self.cmb_linea_llega.setCurrentText("BAJA TENSION")
                            if "1/0" in col: self.rb_bt_1_0.setChecked(True)
                            else: self.rb_bt_2_0.setChecked(True)
                        elif col in ["MEDIA + BAJA TENSION 1/0", "MEDIA + BAJA TENSION 2/0"]:
                             self.cmb_linea_llega.setCurrentText("MEDIA + BAJA TENSION")
                             if "1/0" in col: self.rb_bt_1_0.setChecked(True)
                             else: self.rb_bt_2_0.setChecked(True)
                        else: self.cmb_linea_llega.setCurrentText(col)
                        self.txt_metros.setText(self.formatear_valor_entero(val))
                        break

        self.rb_1f.setChecked(str(row.get("FASES", "")) == "1F")
        self.rb_2f.setChecked(str(row.get("FASES", "")) == "2F")
        self.rb_3f.setChecked(str(row.get("FASES", "")) == "3F")

        self.lst_materiales.clear()
        for col in MATERIALES_COLUMNAS:
            if col == "CIMENTACION": continue
            if es_secundario_duplicado and col != "ESTRUCTURA": continue
            if es_secundario_poste_pf: continue
            
            if col in row.index:
                val = row.get(col, 0)
                if pd.notna(val) and self._tiene_valor(val):
                    vf = self.formatear_valor_entero(val) if col not in MATERIALES_TEXTO else val
                    self.lst_materiales.addItem(f"{col} – {vf}")
        
        if self.es_hijo_vii_actual:
            self.btn_guardar.setEnabled(False)
            self.btn_borrar_registro.setEnabled(False)
            self.txt_poste.setEnabled(False)
            self.txt_po.setEnabled(False)
            self.txt_pf.setEnabled(False)
            self.cmb_tipo_poste.setEnabled(False)
            self.rb_cim_si.setEnabled(False)
            self.rb_cim_no.setEnabled(False)
            self.cmb_tipo_linea.setEnabled(False)
            self.cmb_linea_llega.setEnabled(False)
            self.txt_metros.setEnabled(False)
            self.rb_1f.setEnabled(False)
            self.rb_2f.setEnabled(False)
            self.rb_3f.setEnabled(False)
            self.material_group.setEnabled(False)
        else:
            self.btn_guardar.setEnabled(True)
            self.btn_borrar_registro.setEnabled(True)
            self.txt_poste.setEnabled(True)
            self.txt_po.setEnabled(True)
            self.txt_pf.setEnabled(True)
            self.rb_1f.setEnabled(True)
            self.rb_2f.setEnabled(True)
            self.rb_3f.setEnabled(True)
            self.material_group.setEnabled(True)
            
            if es_secundario_duplicado:
                self.cmb_tipo_poste.setEnabled(False)
                self.rb_cim_si.setEnabled(False)
                self.rb_cim_no.setEnabled(False)
                self.cmb_tipo_linea.setEnabled(False)
                self.cmb_linea_llega.setEnabled(False)
                self.txt_metros.setEnabled(False)
            elif es_secundario_poste_pf:
                self.cmb_tipo_poste.setEnabled(False)
                self.rb_cim_si.setEnabled(False)
                self.rb_cim_no.setEnabled(False)
                self.cmb_tipo_linea.setEnabled(False) 
                if not self.cmb_tipo_linea.currentText():
                    self.cmb_linea_llega.setEnabled(False)
                else:
                    self.cmb_linea_llega.setEnabled(True)
            else:
                self.cmb_tipo_poste.setEnabled(True)
                self.validar_estado_cimentacion()
                self.cmb_tipo_linea.setEnabled(True)
                if not self.cmb_tipo_linea.currentText():
                    self.cmb_linea_llega.setEnabled(False)
                else:
                    self.cmb_linea_llega.setEnabled(True)

            if self.cmb_linea_llega.isEnabled():
                if not self.cmb_linea_llega.currentText():
                    self.txt_metros.setEnabled(False)
                else:
                    self.txt_metros.setEnabled(True)
            
            self.filtrar_materiales()

        self.actualizando = False

        if hasattr(self, 'actualizar_dibujo'):
            self.actualizar_dibujo()

    def _tiene_valor(self, val):
        if pd.isna(val): return False
        s = str(val).strip().lower()
        if not s or s=="nan" or s=="none": return False
        try: 
            if float(s) == 0: return False
        except: pass
        return True

    def navegar_anterior(self):
        if self.txt_poste.text().strip():
            if self.cmb_registro.count()>1 and self.cmb_registro.currentIndex()>0: self.cmb_registro.setCurrentIndex(self.cmb_registro.currentIndex()-1)
            else: self.ir_poste_anterior()
        elif self.postes_unicos:
            self.actualizando = True; self.txt_poste.setText(self.formatear_valor_entero(self.postes_unicos[-1])); self.actualizando = False
            self.buscar_por_poste()
            if self.cmb_registro.count()>0: self.cmb_registro.setCurrentIndex(self.cmb_registro.count()-1)

    def navegar_siguiente(self):
        if self.txt_poste.text().strip():
            if self.cmb_registro.count()>1 and self.cmb_registro.currentIndex()<self.cmb_registro.count()-1: self.cmb_registro.setCurrentIndex(self.cmb_registro.currentIndex()+1)
            else: self.ir_poste_siguiente()
        elif self.postes_unicos:
            self.actualizando = True; self.txt_poste.setText(self.formatear_valor_entero(self.postes_unicos[0])); self.actualizando = False
            self.buscar_por_poste()
            if self.cmb_registro.count()>0: self.cmb_registro.setCurrentIndex(0)

    def ir_a_ultimo(self):
        if not self.postes_unicos: return
        self.actualizando = True; self.txt_poste.setText(self.formatear_valor_entero(self.postes_unicos[-1])); self.actualizando = False
        self.buscar_por_poste()
        if self.cmb_registro.count()>0: self.cmb_registro.setCurrentIndex(self.cmb_registro.count()-1)

    def ir_a_primero(self):
        if not self.postes_unicos: return
        self.actualizando = True; self.txt_poste.setText(self.formatear_valor_entero(self.postes_unicos[0])); self.actualizando = False
        self.buscar_por_poste()
        if self.cmb_registro.count()>0: self.cmb_registro.setCurrentIndex(0)

    def ir_poste_anterior(self):
        if self.poste_actual_index > 0:
            self.actualizando = True; self.txt_poste.setText(self.formatear_valor_entero(self.postes_unicos[self.poste_actual_index-1])); self.actualizando = False
            self.buscar_por_poste()
            if self.cmb_registro.count()>0: self.cmb_registro.setCurrentIndex(self.cmb_registro.count()-1)

    def ir_poste_siguiente(self):
        if 0 <= self.poste_actual_index < len(self.postes_unicos)-1:
            self.actualizando = True; self.txt_poste.setText(self.formatear_valor_entero(self.postes_unicos[self.poste_actual_index+1])); self.actualizando = False
            self.buscar_por_poste()
            if self.cmb_registro.count()>0: self.cmb_registro.setCurrentIndex(0)

    def actualizar_conectores_bt(self):
        if self.actualizando: return
        count_doble = 0
        count_simple = 0
        idx_con = -1
        
        for i in range(self.lst_materiales.count()):
            t = self.lst_materiales.item(i).text()
            parts = t.split(" – ")
            if len(parts) == 2:
                mat = parts[0].strip()
                val = parts[1].strip()
                if mat == "ESTRUCTURA REMATE DOBLE":
                    try: count_doble = int(val)
                    except: pass
                elif mat == "ESTRUCTURA REMATE SIMPLE":
                    try: count_simple = int(val)
                    except: pass
                elif mat == "CONECTORES BT":
                    idx_con = i
                    
        if count_doble == 1 and count_simple > 0:
            if idx_con != -1:
                self.lst_materiales.item(idx_con).setText(f"CONECTORES BT – {count_simple}")
            else:
                self.lst_materiales.addItem(f"CONECTORES BT – {count_simple}")
        else:
            if idx_con != -1:
                self.lst_materiales.takeItem(idx_con)

    def sincronizar_anillo_tierra(self, agregando_trafo=False, quitando_trafo=False):
        if self.actualizando: return
        tiene_trafo = False
        tiene_anillo = False
        idx_anillo = -1
        for i in range(self.lst_materiales.count()):
            t = self.lst_materiales.item(i).text()
            if t.startswith("TRANSFORMADORES – "):
                tiene_trafo = True
            elif t.startswith("ANILLO DE TIERRA – "):
                tiene_anillo = True
                idx_anillo = i
        if agregando_trafo and tiene_trafo and not tiene_anillo:
            self.lst_materiales.addItem("ANILLO DE TIERRA – 1")
        elif quitando_trafo and not tiene_trafo and tiene_anillo:
            self.lst_materiales.takeItem(idx_anillo)

    def cambiar_tipo_input(self):
        mat = self.cmb_material.currentText()
        exists = False
        for i in range(self.lst_materiales.count()):
            if self.lst_materiales.item(i).text().startswith(mat + " – "):
                exists = True
                break
        self.btn_add_mat.setText("Modificar" if exists else "Agregar")

        if mat in ["ESTRUCTURA", "TRANSFORMADORES"]:
            self.lbl_cantidad.setText("Tipo")
            self.txt_cantidad.hide()
            self.cmb_tipo_especial.clear()
            items = TIPOS_ESTRUCTURA if mat == "ESTRUCTURA" else TIPOS_TRANSFORMADOR
            self.cmb_tipo_especial.addItems(["", *items])
            self.cmb_tipo_especial.show()
        else:
            self.lbl_cantidad.setText("Cantidad")
            self.cmb_tipo_especial.hide()
            self.txt_cantidad.show()
            if not exists:
                if mat == "ANCLAS DOBLES":
                    qty = "1"
                    for i in range(self.lst_materiales.count()):
                        txt = self.lst_materiales.item(i).text()
                        if "ESTRUCTURA" in txt:
                            if "– IV" in txt: qty = "2"
                            elif "– V" in txt or "– XI" in txt: qty = "3"
                            elif "– VI" in txt or "– VII" in txt or "– VIII" in txt: qty = "1"
                    self.txt_cantidad.setText(qty)
                elif mat in ("PARARRAYOS", "AMORTIGUADORES"):
                    fase = "1F"
                    if self.rb_2f.isChecked(): fase = "2F"
                    elif self.rb_3f.isChecked(): fase = "3F"
                    if mat == "PARARRAYOS":
                        self.txt_cantidad.setText({"1F": "1", "2F": "2", "3F": "3"}[fase])
                    else:
                        self.txt_cantidad.setText({"1F": "2", "2F": "3", "3F": "4"}[fase])
                else:
                    self.txt_cantidad.setText("1")

    def convertir_ancla_vertical(self):
        if self.actualizando: return
        def get_item(name):
            for i in range(self.lst_materiales.count()):
                t = self.lst_materiales.item(i).text()
                if t.startswith(name + " – "):
                    return i, t.split(" – ")[1].strip()
            return None, None
        def set_qty(name, qty):
            i, _ = get_item(name)
            if i is not None:
                if qty <= 0:
                    self.lst_materiales.takeItem(i)
                else:
                    self.lst_materiales.item(i).setText(f"{name} – {qty}")
            else:
                if qty > 0:
                    self.lst_materiales.addItem(f"{name} – {qty}")
        i, v = get_item("ANCLAS DOBLES")
        if v and v.isdigit() and int(v) > 0:
            set_qty("ANCLAS DOBLES", int(v)-1)
            _, v2 = get_item("ANCLAS DOBLES VERTICALES")
            set_qty("ANCLAS DOBLES VERTICALES", (int(v2) if v2 and v2.isdigit() else 0) + 1)
        else:
            i, v = get_item("ANCLA SIMPLE")
            if v and v.isdigit() and int(v) > 0:
                set_qty("ANCLA SIMPLE", int(v)-1)
                _, v2 = get_item("ANCLA SIMPLE VERTICAL")
                set_qty("ANCLA SIMPLE VERTICAL", (int(v2) if v2 and v2.isdigit() else 0) + 1)
            else:
                i, v = get_item("ANCLAS DOBLES VERTICALES")
                if v and v.isdigit() and int(v) > 0:
                    set_qty("ANCLAS DOBLES VERTICALES", int(v)-1)
                    _, v2 = get_item("ANCLAS DOBLES")
                    set_qty("ANCLAS DOBLES", (int(v2) if v2 and v2.isdigit() else 0) + 1)
                else:
                    i, v = get_item("ANCLA SIMPLE VERTICAL")
                    if v and v.isdigit() and int(v) > 0:
                        set_qty("ANCLA SIMPLE VERTICAL", int(v)-1)
                        _, v2 = get_item("ANCLA SIMPLE")
                        set_qty("ANCLA SIMPLE", (int(v2) if v2 and v2.isdigit() else 0) + 1)
        self.marcar_cambio()
        if hasattr(self, 'actualizar_dibujo'):
            self.actualizar_dibujo()

    def agregar_material(self):
        mat = self.cmb_material.currentText()
        if not mat: return
        val = self.cmb_tipo_especial.currentText() if mat in ["ESTRUCTURA", "TRANSFORMADORES"] else self.txt_cantidad.text().strip()
        if not val: return

        for i in range(self.lst_materiales.count()):
            if self.lst_materiales.item(i).text().startswith(mat + " – "):
                self.lst_materiales.item(i).setText(f"{mat} – {val}")
                self.btn_m_anterior.setVisible(False) 
                self.actualizar_conectores_bt()
                self.sincronizar_anillo_tierra(agregando_trafo=(mat=="TRANSFORMADORES"))
                self.marcar_cambio()
                if hasattr(self, 'actualizar_dibujo'):
                    self.actualizar_dibujo()
                return
        self.lst_materiales.addItem(f"{mat} – {val}")
        self.btn_m_anterior.setVisible(False) 
        self.actualizar_conectores_bt()
        self.sincronizar_anillo_tierra(agregando_trafo=(mat=="TRANSFORMADORES"))
        self.marcar_cambio()
        
        if hasattr(self, 'actualizar_dibujo'):
            self.actualizar_dibujo()

    def quitar_material(self):
        items = self.lst_materiales.selectedItems()
        if not items:
            return
        quitando_trafo = False
        for item in items:
            if item.text().startswith("TRANSFORMADORES – "):
                quitando_trafo = True
            self.lst_materiales.takeItem(self.lst_materiales.row(item))
        self.btn_quitar_mat.setEnabled(False)
        self.btn_m_anterior.setVisible(False) 
        self.actualizar_conectores_bt()
        self.sincronizar_anillo_tierra(quitando_trafo=quitando_trafo)
        self.marcar_cambio()
        self.cambiar_tipo_input()

        if hasattr(self, 'actualizar_dibujo'):
            self.actualizar_dibujo()

    def seleccionar_material(self, item):
        self.btn_quitar_mat.setEnabled(True)
        parts = item.text().split(" – ")
        if len(parts) == 2:
            m, v = parts
            self.cmb_material.setCurrentText(m)
            if m in ["ESTRUCTURA", "TRANSFORMADORES"]: self.cmb_tipo_especial.setCurrentText(v)
            else: self.txt_cantidad.setText(v)

    def guardar_accion(self):
        if self.hay_cambios_pendientes: 
            self.confirmar_cambios_dataframe()
        else: 
            self.guardar_excel()

    def aplicar_filtros_con_excel_real(self, ruta):
        if xw is None: return
        print("--- Iniciando Excel ---")
        app = None
        try:
            ruta_abs = os.path.abspath(ruta)
            app = xw.App(visible=False)
            wb = app.books.open(ruta_abs)
            wb.app.calculate()
            for s in wb.sheets:
                for t in s.api.ListObjects:
                    if t.Name in TABLAS_FILTRAR:
                        print(f"Filtrando {t.Name}")
                        hr = t.HeaderRowRange
                        idx = -1
                        for i in range(1, hr.Columns.Count + 1):
                            if str(hr.Cells(1, i).Value).strip().upper() == "CANTIDAD":
                                idx = i; break
                        if idx != -1: t.Range.AutoFilter(Field=idx, Criteria1="<>0")
            wb.save(); wb.close()
            print("--- Fin Excel ---")
        except Exception as e: print(f"Error xlwings: {e}"); 
        finally: 
            if 'wb' in locals(): 
                try: wb.close() 
                except: pass
            if app: app.quit()

    def guardar_excel(self, mantener_modo_nuevo=False):
        if self.df is None: QMessageBox.warning(self, "Advertencia", "Sin datos"); return
        
        if self.archivo_excel and not self.usando_machote and not os.path.exists(self.archivo_excel):
            QMessageBox.warning(self, "Error", "Excel Base no encontrado. Por favor guarda el archivo como uno nuevo.")
            self.usando_machote = True
            self.archivo_excel = None

        if self.usando_machote or self.archivo_excel is None:
            fecha_actual = datetime.now().strftime("%d.%m.%Y")
            nombre_default = f"Poste a Poste - LUGAR - {fecha_actual}.xlsx"
            path, _ = QFileDialog.getSaveFileName(self, "Guardar", nombre_default, "Excel (*.xlsx)")
            if not path: return
            try:
                if getattr(sys, 'frozen', False): base = sys._MEIPASS
                else: base = os.path.dirname(os.path.abspath(__file__))
                shutil.copy(os.path.join(base, "Poste a Poste - Machote.xlsx"), path)
                self.archivo_excel = path; self.usando_machote = False
            except Exception as e: QMessageBox.critical(self, "Error", str(e)); return

        try:
            wb = load_workbook(self.archivo_excel, data_only=False)
            ws = wb[self.nombre_hoja]
            found = False
            for tbl in ws.tables.values():
                if tbl.name == "P_P":
                    found = True
                    r = tbl.ref
                    ini, fin = r.split(":")
                    cl_ini, fi = coordinate_from_string(ini)
                    ci = column_index_from_string(cl_ini)
                    cl_fin, max_row_old = coordinate_from_string(fin)
                    c_fin = column_index_from_string(cl_fin)
                    
                    fd = fi + 1 
                    
                    df_to_save = self.df.copy()
                    
                    excel_cols = {}
                    for c in range(ci, c_fin + 1):
                        val = ws.cell(row=fi, column=c).value
                        if val:
                            val_str = str(val).strip().upper()
                            if val_str == "PO": excel_cols[c] = "Po"
                            elif val_str == "PF": excel_cols[c] = "Pf"
                            else: excel_cols[c] = str(val).strip()

                    # -- DETECTAR Y CREAR NUEVAS COLUMNAS (MATERIALES NUEVOS) --
                    existing_headers = list(excel_cols.values())
                    nuevas_columnas = [col for col in df_to_save.columns if col in MATERIALES_COLUMNAS and col not in existing_headers]
                    
                    # Para que Excel NO "repare" la tabla (y la convierta en rango) hay que
                    # registrar cada columna nueva también en tbl.tableColumns (con id y
                    # nombre únicos, idénticos al encabezado) y ampliar el autoFilter.
                    from openpyxl.worksheet.table import TableColumn
                    ids_usados = [tc.id for tc in tbl.tableColumns if tc.id is not None]
                    siguiente_id = (max(ids_usados) if ids_usados else 0) + 1
                    nombres_tabla = {str(tc.name).strip().upper() for tc in tbl.tableColumns}

                    for ncol in nuevas_columnas:
                        if str(ncol).strip().upper() in nombres_tabla:
                            continue
                        c_fin += 1
                        celda_hdr = ws.cell(row=fi, column=c_fin, value=ncol)
                        # Copiar el formato del encabezado anterior
                        try:
                            from copy import copy
                            ref_hdr = ws.cell(row=fi, column=c_fin - 1)
                            celda_hdr.font = copy(ref_hdr.font)
                            celda_hdr.fill = copy(ref_hdr.fill)
                            celda_hdr.border = copy(ref_hdr.border)
                            celda_hdr.alignment = copy(ref_hdr.alignment)
                            celda_hdr.number_format = ref_hdr.number_format
                            ws.column_dimensions[get_column_letter(c_fin)].width = \
                                ws.column_dimensions[get_column_letter(c_fin - 1)].width or 14
                        except Exception:
                            pass
                        tbl.tableColumns.append(TableColumn(id=siguiente_id, name=str(ncol)))
                        nombres_tabla.add(str(ncol).strip().upper())
                        siguiente_id += 1
                        excel_cols[c_fin] = ncol
                    # -----------------------------------------------------------

                    for ri, (_, row_data) in enumerate(df_to_save.iterrows(), start=fd):
                        for c_idx, col_name in excel_cols.items():
                            if col_name in df_to_save.columns:
                                val = row_data[col_name]
                                val_s = str(val).lower() if pd.notna(val) else "none"
                                
                                if col_name == "_hijo_vii":
                                    continue
                                
                                if pd.isna(val) or val_s in ["none", "nan", ""]: 
                                    ws.cell(row=ri, column=c_idx, value="")
                                elif col_name in ["Po", "Pf", "NO POSTE"] + COLUMNAS_LINEA + [m for m in MATERIALES_COLUMNAS if m not in MATERIALES_TEXTO]:
                                    try:
                                        num_val = float(val)
                                        ws.cell(row=ri, column=c_idx, value=int(num_val) if num_val.is_integer() else num_val)
                                    except:
                                        ws.cell(row=ri, column=c_idx, value=val)
                                elif isinstance(val, (int, float, np.integer, np.floating)):
                                    ws.cell(row=ri, column=c_idx, value=float(val) if val!=0 else 0)
                                else: 
                                    ws.cell(row=ri, column=c_idx, value=str(val))
                    
                    mf = fd + len(df_to_save) - 1
                    if mf < fd: mf = fd 
                    
                    if max_row_old > mf:
                        ws.delete_rows(mf + 1, max_row_old - mf)
                    
                    # Ref ahora contempla las nuevas columnas generadas dinámicamente
                    tbl.ref = f"{get_column_letter(ci)}{fi}:{get_column_letter(c_fin)}{mf}"
                    # El autoFilter debe cubrir exactamente el mismo rango que la tabla
                    if tbl.autoFilter is not None:
                        tbl.autoFilter.ref = tbl.ref
                    break
            
            if not found: wb.close(); QMessageBox.critical(self, "Error", "No P_P"); return
            wb.save(self.archivo_excel); wb.close()
            
            time.sleep(1) 
            
            if xw: self.aplicar_filtros_con_excel_real(self.archivo_excel)
            else: QMessageBox.warning(self, "Aviso", "xlwings no instalado")
            
            self.hay_cambios_pendientes = False; self.registros_sin_guardar = 0; self.cambios_catalogo = False
            self.btn_guardar.setText("Guardar al Excel")
            if not mantener_modo_nuevo: self.establecer_modo_editar()
            QMessageBox.information(self, "Éxito", "Guardado")
        except Exception as e: QMessageBox.critical(self, "Error", str(e))

    def alternar_visor_cad(self):
        self._cad_visible = not getattr(self, "_cad_visible", False)
        self.panel_cad.setVisible(self._cad_visible)
        if self._cad_visible:
            self.resize(self.ANCHO_CON_CAD, self.height())
            QTimer.singleShot(100, self._refrescar_cad)
        else:
            self.resize(self.ANCHO_BASE, self.height())

    def _refrescar_cad(self):
        if hasattr(self, 'actualizar_dibujo'):
            self.actualizar_dibujo()
        if hasattr(self, 'vista_grafica'):
            self.vista_grafica.ajustar_a_contenido()

    def actualizar_dibujo(self):
        try:
            if not getattr(self, "_cad_visible", False):
                return
            import re
            no_poste = str(self.txt_poste.text()).strip()
            if not no_poste:
                self.escena.clear()
                return

            tiene_cimentacion = self.rb_cim_si.isChecked()
            tipo_linea = self.cmb_tipo_linea.currentText()
            
            fases = "3F" if self.rb_3f.isChecked() else ("2F" if self.rb_2f.isChecked() else "1F")
            textos_info = [fases]
            
            texto_poste_largo = self.cmb_tipo_poste.currentText().strip()
            nombre_bloque_poste = ""
            altura_poste = ""
            dan_poste = ""

            if texto_poste_largo:
                if "concreto" in texto_poste_largo.lower(): nombre_bloque_poste = "P.C"
                elif "madera" in texto_poste_largo.lower(): nombre_bloque_poste = "P.M."
                elif "metal" in texto_poste_largo.lower(): nombre_bloque_poste = "P.METAL"
                elif "fibra" in texto_poste_largo.lower(): nombre_bloque_poste = "P.F"
                
                for h in ["30 pies", "35 pies", "40 pies", "44 pies", "45 pies"]:
                    if h in texto_poste_largo:
                        altura_poste = h
                        break
                        
                if "madera" in texto_poste_largo.lower():
                    match = re.search(r'Clase\s+(\d+)', texto_poste_largo, re.IGNORECASE)
                    if match: dan_poste = match.group(1)
                else:
                    match = re.search(r'(\d+)\s*daN', texto_poste_largo, re.IGNORECASE)
                    if match: dan_poste = match.group(1)

            lista_anclas = []
            elementos_rotantes = []
            estructura_tipo = ""

            for i in range(self.lst_materiales.count()):
                texto = self.lst_materiales.item(i).text()
                if " – " in texto:
                    mat, val = texto.split(" – ")
                    mat, val = mat.strip(), val.strip()
                    if "ANCLA" in mat:
                        for _ in range(int(val) if val.isdigit() else 0): 
                            lista_anclas.append(mat)
                    elif mat == "ESTRUCTURA":
                        estructura_tipo = val
                    elif mat in ["ESTRUCTURA REMATE DOBLE", "ESTRUCTURA REMATE SIMPLE"]:
                        elementos_rotantes.append({"id": mat, "texto": val})
                    elif mat in ["TRANSFORMADORES", "CORTACIRCUITO", "PARARRAYOS", "AMORTIGUADORES", "TIERRAS", "ANILLO DE TIERRA"]:
                        elementos_rotantes.append({"id": mat, "texto": val})

            if hasattr(self, 'df') and self.df is not None:
                if "No. Poste" in self.df.columns:
                    registros_historicos = self.df[self.df["No. Poste"].astype(str).str.strip() == no_poste]
                    idx_actual = getattr(self, 'registro_actual_idx', -1)
                    
                    for idx, row in registros_historicos.iterrows():
                        if idx == idx_actual: continue 
                        if str(row.get("CIMENTACION", "")).lower() == "si": tiene_cimentacion = True
                        
                        for col_a in ["ANCLAS DOBLES", "ANCLA SIMPLE", "ANCLAS DOBLES VERTICALES", "ANCLA SIMPLE VERTICAL"]:
                            val_a = row.get(col_a, 0)
                            if pd.notna(val_a) and str(val_a).strip() not in ["", "0", "0.0", "nan"]:
                                for _ in range(int(float(val_a))): lista_anclas.append(col_a)
                        
                        for mat_h in ["TRANSFORMADORES", "CORTACIRCUITO", "PARARRAYOS", "AMORTIGUADORES", "TIERRAS", "ANILLO DE TIERRA"]:
                            val_h = row.get(mat_h, "")
                            if pd.notna(val_h) and str(val_h).strip() not in ["", "0", "0.0", "nan"]:
                                elementos_rotantes.append({"id": mat_h, "texto": str(val_h).strip()})

            import graphics
            bloques_faltantes = graphics.ensamblar_y_dibujar(
                escena=self.escena, 
                vista_grafica=self.vista_grafica, 
                nombre_poste=nombre_bloque_poste,
                altura_poste=altura_poste,
                dan_poste=dan_poste,
                tiene_cimentacion=tiene_cimentacion, 
                textos_info=list(set(textos_info)),
                lista_anclas=lista_anclas, 
                elementos_rotantes=elementos_rotantes,
                tipo_linea_texto=tipo_linea, 
                num_poste=no_poste,
                fases=fases,
                estructura_tipo=estructura_tipo,
                modo_dia=not self.modo_oscuro
            )

            if bloques_faltantes:
                self.lbl_bloques_faltantes.setText(
                    "Bloques genéricos (sin bloque útil en LIBRERIA.dxf): " + ", ".join(bloques_faltantes)
                )
            else:
                self.lbl_bloques_faltantes.setText("")
        except Exception:
            import traceback
            print(traceback.format_exc())

            
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    ruta_splash = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__))), "Carga.png")
    
    if os.path.exists(ruta_splash):
        pixmap = QPixmap(ruta_splash)
        
        screen_geometry = app.primaryScreen().geometry()
        
        target_width = int(screen_geometry.width() * 0.45)
        target_height = int(screen_geometry.height() * 0.45)
        
        pixmap_escalado = pixmap.scaled(
            target_width, 
            target_height, 
            Qt.KeepAspectRatio, 
            Qt.SmoothTransformation
        )
        
        splash = QSplashScreen(pixmap_escalado, Qt.WindowStaysOnTopHint)
        splash.show()
        app.processEvents()
        time.sleep(3)  
    else:
        splash = None

    w = Formulario()
    w.show()
    
    if splash:
        splash.finish(w)
        
    sys.exit(app.exec())