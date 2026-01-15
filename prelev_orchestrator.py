# -*- coding: utf-8 -*-

import os
import shutil
import traceback
from qgis.PyQt import QtWidgets, QtCore, QtGui
from qgis.core import (
    QgsApplication, QgsProject, QgsVectorLayer, QgsFeatureRequest,
    QgsWkbTypes, QgsFeature
)
from qgis import processing
from qgis.utils import iface


# =========================
# CONFIG
# =========================

PLUGIN_DIR = os.path.abspath(os.path.dirname(__file__))
BASE_FOLDER = os.path.join(PLUGIN_DIR, 'Couches')
NETWORK_SCRIPTS_FOLDER = os.path.join(PLUGIN_DIR, 'scripts')
QML_COUCHES_FOLDER = os.path.join(BASE_FOLDER, 'QML_Couches')

VOCAL_ZONE_PROPERTY = "VOCAL_ACTIVE_ZONE_LAYER_ID"   # ### NOUVEAU


GPKG_MAP = {
    'Délégation': 'limite_Deleg.gpkg',
    'Départements': 'departements.gpkg',
    'Bassins versants': 'BV.gpkg',
    'Communes': 'communes.gpkg',
    'Nappes': 'nappes.gpkg',
    'UG PGRE 34': 'UG_PGRE_34.gpkg'
}

DEPT_FIELD = 'nom_dept'
BV_FIELD = 'lib_ssbv'


ALGO_INFOS = {
    'Evolution des volumes prélevés par ouvrage': {
        'alg_id': 'script:compute_slopes_ouvrage_only',
        'script_name': 'compute_slopes_qgis_ouvrages.py'
    },
    'Evolution des volumes prélevés agrégés par zone': {
        'alg_id': 'script:compute_slopes_zones',
        'script_name': 'compute_slopes_qgis_zonages.py'
    },
    'Ratio VolPrelev/VolAutorise par ouvrage': {
        'alg_id': 'script:compare_prelevements_autorises',
        'script_name': 'compute_ratio_VPVA_ouvrages.py'
    },
    'Ratio VolPrelev/VolAutorise par zonage': {
        'alg_id': 'script:zones_compare_prelev_autorise',
        'script_name': 'compute_ratio_VPVA_zonages.py'
    },
    "État connaissance - ouvrages Agence": {
        'alg_id': 'script:compute_connaissance_ouvrages_agence',
        'script_name': 'compute_connaissance_ouvrages_agence.py'
    }
}


# =========================
# HELPERS ZONE VOCAL
# =========================

def get_active_vocal_zone():
    """Retourne la couche zone d’étude VOCAL si définie et valide"""
    lid = QgsProject.instance().customProperty(VOCAL_ZONE_PROPERTY, None)
    if not lid:
        return None
    layer = QgsProject.instance().mapLayer(lid)
    if not layer or not isinstance(layer, QgsVectorLayer):
        return None
    if QgsWkbTypes.geometryType(layer.wkbType()) != QgsWkbTypes.PolygonGeometry:
        return None
    return layer


def set_active_vocal_zone(layer):
    QgsProject.instance().setCustomProperty(VOCAL_ZONE_PROPERTY, layer.id())


def zoom_to_layer(layer):
    if not layer:
        return
    canvas = iface.mapCanvas()
    ext = layer.extent()
    if not ext or ext.isEmpty():
        return
    canvas.setExtent(ext)
    canvas.refresh()


# =========================
# SCRIPTS PROCESSING
# =========================

def ensure_scripts_in_user_folder(feedback=None):
    try:
        user_proc_scripts = os.path.join(
            QgsApplication.qgisSettingsDirPath(), 'processing', 'scripts'
        )
        os.makedirs(user_proc_scripts, exist_ok=True)

        for info in ALGO_INFOS.values():
            src = os.path.join(NETWORK_SCRIPTS_FOLDER, info['script_name'])
            dst = os.path.join(user_proc_scripts, info['script_name'])
            if os.path.exists(src) and not os.path.exists(dst):
                shutil.copy2(src, dst)
                if feedback:
                    feedback(f"[VOCAL] Script copié : {dst}")
    except Exception as e:
        if feedback:
            feedback(str(e))


# =========================
# UI
# =========================

class PrelevOrchestratorDialog(QtWidgets.QDialog):

    def __init__(self, parent=None):
        super().__init__(parent or iface.mainWindow())
        self.setWindowTitle("VOCAL — Orchestrateur")
        self.resize(820, 520)

        self.selected_program = None

        self.stack = QtWidgets.QStackedWidget()
        self.page1 = QtWidgets.QWidget()
        self.page2 = QtWidgets.QWidget()

        self._build_page1()
        self._build_page2()

        self.stack.addWidget(self.page1)
        self.stack.addWidget(self.page2)

        self.prev_btn = QtWidgets.QPushButton("Précédent")
        self.next_btn = QtWidgets.QPushButton("Suivant")
        self.open_btn = QtWidgets.QPushButton("Ouvrir l’outil Processing")

        self.prev_btn.clicked.connect(self.on_prev)
        self.next_btn.clicked.connect(self.on_next)
        self.open_btn.clicked.connect(self.on_open_algo)

        self.open_btn.setEnabled(False)
        self.prev_btn.setEnabled(False)

        btns = QtWidgets.QHBoxLayout()
        btns.addStretch()
        btns.addWidget(self.prev_btn)
        btns.addWidget(self.next_btn)
        btns.addWidget(self.open_btn)

        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.stack)
        layout.addLayout(btns)
        self.setLayout(layout)

        self.refresh_vocal_zone_ui()

    # =====================
    # PAGE 1
    # =====================

    def _build_page1(self):
        layout = QtWidgets.QVBoxLayout()

        # ---- Zone VOCAL ----
        grp_zone = QtWidgets.QGroupBox("Zone d’étude VOCAL")
        v = QtWidgets.QVBoxLayout()

        self.zone_status_label = QtWidgets.QLabel()
        self.define_zone_btn = QtWidgets.QPushButton("Définir / modifier la zone d’étude")
        self.define_zone_btn.clicked.connect(self.on_define_zone)

        v.addWidget(self.zone_status_label)
        v.addWidget(self.define_zone_btn)
        grp_zone.setLayout(v)

        # ---- Programme ----
        grp_prog = QtWidgets.QGroupBox("Choix du programme")
        p = QtWidgets.QVBoxLayout()
        self.prog_combo = QtWidgets.QComboBox()
        for k in ALGO_INFOS:
            self.prog_combo.addItem(k)
        p.addWidget(self.prog_combo)
        grp_prog.setLayout(p)

        layout.addWidget(grp_zone)
        layout.addWidget(grp_prog)
        self.page1.setLayout(layout)

    def refresh_vocal_zone_ui(self):
        layer = get_active_vocal_zone()
        if layer:
            self.zone_status_label.setText(
                f"✔ Zone active : <b>{layer.name()}</b>"
            )
            self.next_btn.setEnabled(True)
        else:
            self.zone_status_label.setText(
                "⚠ Aucune zone d’étude définie"
            )
            self.next_btn.setEnabled(False)

    def on_define_zone(self):
        layers = [
            l for l in QgsProject.instance().mapLayers().values()
            if isinstance(l, QgsVectorLayer)
            and QgsWkbTypes.geometryType(l.wkbType()) == QgsWkbTypes.PolygonGeometry
        ]

        if not layers:
            QtWidgets.QMessageBox.warning(
                self, "Erreur",
                "Aucune couche polygonale disponible dans le projet."
            )
            return

        names = {l.name(): l for l in layers}
        name, ok = QtWidgets.QInputDialog.getItem(
            self, "Zone d’étude VOCAL",
            "Choisir une couche polygonale :",
            list(names.keys()), 0, False
        )

        if ok:
            layer = names[name]
            set_active_vocal_zone(layer)
            zoom_to_layer(layer)
            self.refresh_vocal_zone_ui()

    # =====================
    # PAGE 2
    # =====================

    def _build_page2(self):
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(QtWidgets.QLabel(
            "Les scripts Processing sont prêts.\n"
            "Clique sur « Ouvrir l’outil Processing »."
        ))
        self.page2.setLayout(layout)

    # =====================
    # NAVIGATION
    # =====================

    def on_next(self):
        if not get_active_vocal_zone():
            QtWidgets.QMessageBox.warning(
                self, "Zone manquante",
                "Veuillez définir une zone d’étude VOCAL."
            )
            return

        self.selected_program = self.prog_combo.currentText()
        ensure_scripts_in_user_folder(print)

        self.stack.setCurrentIndex(1)
        self.prev_btn.setEnabled(True)
        self.next_btn.setEnabled(False)
        self.open_btn.setEnabled(True)

    def on_prev(self):
        self.stack.setCurrentIndex(0)
        self.prev_btn.setEnabled(False)
        self.next_btn.setEnabled(True)
        self.open_btn.setEnabled(False)

    def on_open_algo(self):
        info = ALGO_INFOS.get(self.selected_program)
        if not info:
            return

        alg_id = info['alg_id']
        try:
            self.accept()
            QtCore.QTimer.singleShot(
                150, lambda: processing.execAlgorithmDialog(alg_id)
            )
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Erreur",
                f"Impossible d’ouvrir l’outil Processing\n{e}"
            )


# =========================
# PLUGIN
# =========================

class PrelevOrchestratorPlugin:

    def __init__(self, iface):
        self.iface = iface
        self.action = None

    def initGui(self):
        self.action = QtWidgets.QAction("VOCAL", self.iface.mainWindow())
        self.action.triggered.connect(self.run)
        self.iface.addToolBarIcon(self.action)

    def unload(self):
        self.iface.removeToolBarIcon(self.action)

    def run(self):
        dlg = PrelevOrchestratorDialog()
        dlg.exec_()
