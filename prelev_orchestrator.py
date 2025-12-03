# -*- coding: utf-8 -*-
"""
Prelev Orchestrator plugin - updated version with additional program integration
Remplace intégralement le fichier prelev_orchestrator.py précédent.

Cette version ajoute l'algorithme :
  - "Etat de la connaissance Agence des ouvrages de prélèvements d'eau"
    id recommandé (tel que fourni) : 'script:compute_connaissance_ouvrages_agence'
    script filename : compute_connaissance_ouvrages_agence.py

Le comportement du plugin reste identique : il prépare la zone d'étude,
copie les scripts réseau dans le dossier Processing/scripts utilisateur si nécessaire,
et ouvre la boîte de dialogue native Processing pour l'algorithme choisi.
"""

import os
import shutil
import traceback
from qgis.PyQt import QtWidgets, QtCore, QtGui
from qgis.core import (
    QgsApplication, QgsProject, QgsVectorLayer, QgsFeatureRequest,
    QgsWkbTypes, QgsFeature, QgsFields, QgsGeometry
)
from qgis import processing
from qgis.utils import iface

# ---------------- USER CONFIG ----------------
PLUGIN_DIR = os.path.abspath(os.path.dirname(__file__))

BASE_FOLDER = os.path.join(PLUGIN_DIR, 'Couches')
NETWORK_SCRIPTS_FOLDER = os.path.join(PLUGIN_DIR, 'scripts')
QML_COUCHES_FOLDER = os.path.join(BASE_FOLDER, 'QML_Couches')


# DEBUG (optionnel) : affiche dans la console où l'on lira les couches/scripts
print(f"[Orch] PLUGIN_DIR = {PLUGIN_DIR}")
print(f"[Orch] BASE_FOLDER = {BASE_FOLDER}")
print(f"[Orch] NETWORK_SCRIPTS_FOLDER = {NETWORK_SCRIPTS_FOLDER}")
print(f"[Orch] QML_COUCHES_FOLDER = {QML_COUCHES_FOLDER}")

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

# ---------------- Added/Updated Algorithms ----------------
ALGO_INFOS = {
    'Pentes par ouvrage': {
        'alg_id': 'script:compute_slopes_ouvrage_only',
        'script_name': 'compute_slopes_qgis_ouvrages.py'
    },
    'Pentes par zonage': {
        'alg_id': 'script:compute_slopes_zones',
        'script_name': 'compute_slopes_qgis_zonages.py'
    },
    'Ratio VP/ VA par ouvrage': {
        'alg_id': 'script:compare_prelevements_autorises',
        'script_name': 'compute_ratio_VPVA_ouvrages.py'
    },
    'Ratio VP/ VA par zonage': {
        'alg_id': 'script:zones_compare_prelev_autorise',
        'script_name': 'compute_ratio_VPVA_zonages.py'
    },
    
    "État connaissance - ouvrages Agence": {
        # use the exact id you provided (include 'script:' prefix if that's how it appears in the Toolbox)
        'alg_id': 'script:compute_connaissance_ouvrages_agence',
        'script_name': 'compute_connaissance_ouvrages_agence.py'
    }
    #D'autres programmes peuvent etre ajouter ici. Attention aux noms des programmes en mettant script: devant.
}

# ---------------- Helpers ----------------

def gpkg_path_for(scale_label):
    fname = GPKG_MAP.get(scale_label)
    if not fname:
        return None
    return os.path.join(BASE_FOLDER, fname)

def try_load_gpkg_layer(gpkg_path):
    """Essaie de charger une couche depuis un GeoPackage. Retourne QgsVectorLayer (non ajouté) ou None."""
    if not gpkg_path or not os.path.exists(gpkg_path):
        return None
    base = os.path.splitext(os.path.basename(gpkg_path))[0]
    candidates = [base, base.lower(), base.upper(), 'departements', 'communes', 'BV', 'bv', 'nappes', 'limite_Deleg']
    for name in candidates:
        uri = f"{gpkg_path}|layername={name}"
        layer = QgsVectorLayer(uri, f"{name}", "ogr")
        if layer.isValid():
            return layer
    layer = QgsVectorLayer(gpkg_path, base, "ogr")
    if layer.isValid():
        return layer
    return None

def list_zone_values(layer, fieldname):
    vals = set()
    if layer is None or fieldname is None:
        return []
    if layer.fields().indexFromName(fieldname) < 0:
        return []
    for f in layer.getFeatures():
        v = f[fieldname]
        if v is None:
            continue
        vals.add(str(v))
    return sorted(vals)

def load_layer_to_project(layer, add_if_not=True):
    if layer is None:
        return None
    existing = QgsProject.instance().mapLayersByName(layer.name())
    if existing:
        return existing[0]
    if add_if_not:
        QgsProject.instance().addMapLayer(layer)
    return layer

def zoom_to_layer(layer):
    if not layer:
        return
    canvas = iface.mapCanvas()
    try:
        extent = layer.extent()
    except Exception:
        return
    if extent is None or extent.isEmpty():
        return
    canvas.setExtent(extent)
    canvas.refresh()

def _geom_type_string_from_wkb(wkb):
    """Retourne 'Polygon'/'LineString'/'Point' ou 'Unknown' à partir d'un wkbType."""
    try:
        gt = QgsWkbTypes.geometryType(wkb)
        if gt == QgsWkbTypes.PolygonGeometry:
            return 'Polygon'
        if gt == QgsWkbTypes.LineGeometry:
            return 'LineString'
        if gt == QgsWkbTypes.PointGeometry:
            return 'Point'
    except Exception:
        pass
    try:
        s = QgsWkbTypes.displayString(wkb)
        s = s.lower() if isinstance(s, str) else ''
        if 'polygon' in s:
            return 'Polygon'
        if 'line' in s or 'linestring' in s:
            return 'LineString'
        if 'point' in s:
            return 'Point'
    except Exception:
        pass
    return 'Unknown'

def create_memory_layer_from_features(source_layer, features, name_suffix="_mem"):
    """Crée une couche mémoire à partir d'une liste de features (copie champs/crs/geom)."""
    if source_layer is None or not features:
        return None

    geom_type = _geom_type_string_from_wkb(source_layer.wkbType())
    if geom_type == 'Unknown':
        try:
            test_geom = features[0].geometry()
            if test_geom is not None:
                t = test_geom.type()
                if t == 2:  # polygon
                    geom_type = 'Polygon'
                elif t == 1:
                    geom_type = 'LineString'
                else:
                    geom_type = 'Point'
        except Exception:
            geom_type = 'Polygon'

    crs_auth = source_layer.crs().authid() if source_layer.crs() else ''
    layer_name = f"{source_layer.name()}{name_suffix}"
    uri = f"{geom_type}?crs={crs_auth}"
    mem = QgsVectorLayer(uri, layer_name, "memory")
    dp = mem.dataProvider()

    try:
        dp.addAttributes(list(source_layer.fields()))
        mem.updateFields()
    except Exception:
        pass

    feats_to_add = []
    for f in features:
        nf = QgsFeature()
        nf.setFields(mem.fields())
        try:
            nf.setGeometry(f.geometry())
        except Exception:
            pass
        try:
            nf.setAttributes(f.attributes())
        except Exception:
            attrs = []
            for fld in mem.fields():
                try:
                    attrs.append(f.attribute(fld.name()))
                except Exception:
                    attrs.append(None)
            try:
                nf.setAttributes(attrs)
            except Exception:
                pass
        feats_to_add.append(nf)

    try:
        dp.addFeatures(feats_to_add)
        mem.updateExtents()
        QgsProject.instance().addMapLayer(mem)
        return mem
    except Exception:
        return None

def ensure_scripts_in_user_folder(feedback=None):
    """Copy network scripts into the user's processing scripts folder (if missing)."""
    out = []
    try:
        user_proc_scripts = os.path.join(QgsApplication.qgisSettingsDirPath(), 'processing', 'scripts')
        os.makedirs(user_proc_scripts, exist_ok=True)
        for info in ALGO_INFOS.values():
            sn = info.get('script_name')
            if not sn:
                continue
            src = os.path.join(NETWORK_SCRIPTS_FOLDER, sn)
            dst = os.path.join(user_proc_scripts, sn)
            if os.path.exists(src):
                try:
                    if not os.path.exists(dst) or os.path.getsize(dst) != os.path.getsize(src):
                        shutil.copy2(src, dst)
                        if feedback:
                            feedback(f"[Orch] Copié script -> {dst}")
                    else:
                        if feedback:
                            feedback(f"[Orch] Script déjà présent -> {dst}")
                    out.append(dst)
                except Exception as e:
                    if feedback:
                        feedback(f"[Orch] Erreur copie script {src} : {e}")
            else:
                if feedback:
                    feedback(f"[Orch] Script source introuvable (réseau) : {src}")
    except Exception as e:
        if feedback:
            feedback(f"[Orch] Erreur lors de la mise en place des scripts utilisateurs : {e}")
    return out

# ---------------- Main UI ----------------
class PrelevOrchestratorDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent or iface.mainWindow())
        self.setWindowTitle('Le VOCAL')
        self.resize(920, 560)

        # stacked pages
        self.stack = QtWidgets.QStackedWidget()
        self.page1 = QtWidgets.QWidget()
        self.page2 = QtWidgets.QWidget()
        self._build_page1()
        self._build_page2()
        self.stack.addWidget(self.page1)
        self.stack.addWidget(self.page2)

        # buttons
        btn_box = QtWidgets.QHBoxLayout()
        self.prev_btn = QtWidgets.QPushButton('Précédent')
        self.prev_btn.clicked.connect(self.on_prev)
        self.next_btn = QtWidgets.QPushButton('Suivant')
        self.next_btn.clicked.connect(self.on_next)
        self.open_algo_btn = QtWidgets.QPushButton("Ouvrir l'outil Processing")
        self.open_algo_btn.clicked.connect(self.on_open_algo)
        self.open_algo_btn.setEnabled(False)
        btn_box.addStretch()
        btn_box.addWidget(self.prev_btn)
        btn_box.addWidget(self.next_btn)
        btn_box.addWidget(self.open_algo_btn)

        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.stack)
        layout.addLayout(btn_box)
        self.setLayout(layout)

        self.prev_btn.setEnabled(False)

        # state
        self.selected_program = None
        self.zone_layer = None           # couche source gpkg (non ajoutée)
        self.zone_mem_layer = None       # couche mémoire limitée
        self.zone_value = None
        self.optional_zonage_layer = None

    def _build_page1(self):
        layout = QtWidgets.QVBoxLayout()

        grp_prog = QtWidgets.QGroupBox('1) Choix du programme')
        v = QtWidgets.QVBoxLayout()
        self.prog_combo = QtWidgets.QComboBox()
        for k in ALGO_INFOS.keys():
            self.prog_combo.addItem(k)
        v.addWidget(self.prog_combo)
        grp_prog.setLayout(v)

        grp_zone = QtWidgets.QGroupBox('2) Zone d\'étude')
        gz = QtWidgets.QGridLayout()
        gz.addWidget(QtWidgets.QLabel('Echelle'), 0, 0)
        self.scale_combo = QtWidgets.QComboBox()
        self.scale_combo.addItems(list(GPKG_MAP.keys()))
        self.scale_combo.currentTextChanged.connect(self.on_scale_changed)
        gz.addWidget(self.scale_combo, 0, 1)
        gz.addWidget(QtWidgets.QLabel('Valeur'), 1, 0)
        self.zone_value_combo = QtWidgets.QComboBox()
        gz.addWidget(self.zone_value_combo, 1, 1)
        self.load_zone_btn = QtWidgets.QPushButton('Charger zone et zoom')
        self.load_zone_btn.clicked.connect(self.on_load_zone)
        gz.addWidget(self.load_zone_btn, 2, 0, 1, 2)
        grp_zone.setLayout(gz)

        # checkbox memory layer
        self.create_memory_checkbox = QtWidgets.QCheckBox("Créer couche mémoire limitée à la zone d'étude (recommandé)")
        self.create_memory_checkbox.setChecked(True)

        # checkbox to reveal zonage options (NEW)
        self.show_zonage_checkbox = QtWidgets.QCheckBox("Voulez-vous charger un sous-zonage ?")
        self.show_zonage_checkbox.setChecked(False)

        # optional zonage: combo + browse, combo contains server gpkg entries + project vector layers
        grp_optional = QtWidgets.QGroupBox('Optionnel : zonage (server / projet / fichier)')
        gh = QtWidgets.QGridLayout()
        gh.addWidget(QtWidgets.QLabel("Choisir zonage (serveur / projet)"), 0, 0)
        self.zonage_combo = QtWidgets.QComboBox()
        gh.addWidget(self.zonage_combo, 0, 1)
        gh.addWidget(QtWidgets.QLabel("Ou parcourir un fichier"), 1, 0)
        self.zonage_path_edit = QtWidgets.QLineEdit()
        self.zonage_browse = QtWidgets.QPushButton('Parcourir')
        self.zonage_browse.clicked.connect(self.on_browse_zonage)
        h2 = QtWidgets.QHBoxLayout()
        h2.addWidget(self.zonage_path_edit)
        h2.addWidget(self.zonage_browse)
        gh.addLayout(h2, 1, 1)
        grp_optional.setLayout(gh)

        # fill zonage_combo with server gpkg bases + project vector layers
        self._populate_zonage_combo()

        # hide optional zonage by default; will be shown only when user checks the checkbox
        grp_optional.setVisible(False)
        self.show_zonage_checkbox.toggled.connect(lambda checked: grp_optional.setVisible(checked))

        # QML options
        qml_box = QtWidgets.QGroupBox('QML (appliquer aux couches chargées)')
        qh = QtWidgets.QFormLayout()
        self.qml_zone_checkbox = QtWidgets.QCheckBox('Appliquer QML à la couche zone d\'étude (ou mémoire)')
        self.qml_zonage_checkbox = QtWidgets.QCheckBox('Appliquer QML à la couche zonage (si fournie)')
        self.qml_zone_checkbox.setChecked(True)
        self.qml_zonage_checkbox.setChecked(True)
        qh.addRow(self.qml_zone_checkbox)
        qh.addRow(self.qml_zonage_checkbox)
        qml_box.setLayout(qh)

        info = QtWidgets.QLabel('Remarque : les scripts seront copiés dans ton dossier Processing/scripts utilisateur si nécessaire.\n'
                                 'Clique "Suivant" pour copier les scripts et préparer l\'ouverture de l\'outil Processing.\n'
                                 'Le plugin se fermera automatiquement lorsqu\'il ouvrira l\'outil Processing.')

        layout.addWidget(grp_prog)
        layout.addWidget(grp_zone)
        layout.addWidget(self.create_memory_checkbox)
        layout.addWidget(self.show_zonage_checkbox)
        layout.addWidget(grp_optional)
        layout.addWidget(qml_box)
        layout.addWidget(info)
        self.page1.setLayout(layout)

        self.on_scale_changed(self.scale_combo.currentText())

    def _populate_zonage_combo(self):
        self.zonage_combo.clear()
        # server gpkg candidates (use file basename as label, store full path in data)
        try:
            for fname in os.listdir(BASE_FOLDER):
                if fname.lower().endswith('.gpkg'):
                    full = os.path.join(BASE_FOLDER, fname)
                    self.zonage_combo.addItem(f"[srv] {fname}", full)
        except Exception:
            pass
        # separator
        self.zonage_combo.addItem("---- Couches du projet ----", None)
        # project vector layers
        for lyr in QgsProject.instance().mapLayers().values():
            if isinstance(lyr, QgsVectorLayer):
                self.zonage_combo.addItem(f"[proj] {lyr.name()}", lyr.id())

    def on_scale_changed(self, text):
        gpkg = gpkg_path_for(text)
        layer = try_load_gpkg_layer(gpkg)
        self.zone_value_combo.clear()
        if layer is None:
            self.zone_value_combo.addItem('-- couche introuvable --')
            return
        if text == 'Départements':
            field = DEPT_FIELD
        elif text == 'Bassins versants':
            field = BV_FIELD
        else:
            if layer.fields().indexFromName('name') >= 0:
                field = 'name'
            else:
                field = None
                for f in layer.fields():
                    if f.typeName().lower().startswith('string'):
                        field = f.name()
                        break
                if field is None:
                    field = layer.fields()[0].name()
        vals = list_zone_values(layer, field)
        if vals:
            self.zone_value_combo.addItems(vals)
        else:
            self.zone_value_combo.addItem('-- Aucun attribut trouvé --')

    def on_browse_zonage(self):
        fp, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Choisir une couche zonage (gpkg/shp)', BASE_FOLDER, 'GeoPackage (*.gpkg);;Shapefile (*.shp);;All (*)')
        if not fp:
            return
        self.zonage_path_edit.setText(fp)

    def on_load_zone(self):
        scale = self.scale_combo.currentText()
        val = self.zone_value_combo.currentText()
        gpkg = gpkg_path_for(scale)
        layer = try_load_gpkg_layer(gpkg)
        if layer is None:
            QtWidgets.QMessageBox.warning(self, 'Erreur', f'Impossible de charger {gpkg}.')
            return

        if scale == 'Départements':
            field = DEPT_FIELD
        elif scale == 'Bassins versants':
            field = BV_FIELD
        else:
            if layer.fields().indexFromName('name') >= 0:
                field = 'name'
            else:
                field = layer.fields()[0].name()

        expr = f'"{field}" = \'{val}\''
        try:
            it = layer.getFeatures(QgsFeatureRequest().setFilterExpression(expr))
            feats = [f for f in it]
        except Exception:
            feats = []

        self.zone_layer = layer
        self.zone_value = val

        if self.create_memory_checkbox.isChecked():
            if feats:
                mem = create_memory_layer_from_features(layer, feats, name_suffix=f"_INTER_{self.zone_value or ''}")
                if mem is not None:
                    self.zone_mem_layer = mem
                    if self.qml_zone_checkbox.isChecked():
                        qmlname = f"QML_{os.path.splitext(os.path.basename(gpkg or ''))[0]}"
                        qmlpath = os.path.join(QML_COUCHES_FOLDER, qmlname + '.qml')
                        if os.path.exists(qmlpath):
                            try:
                                mem.loadNamedStyle(qmlpath)
                                mem.triggerRepaint()
                            except Exception:
                                pass
                    zoom_to_layer(mem)
                else:
                    self.zone_mem_layer = None
                    load_layer_to_project(layer, add_if_not=True)
                    try:
                        ids = [f.id() for f in feats]
                        layer.removeSelection()
                        layer.selectByIds(ids)
                        canvas = iface.mapCanvas()
                        canvas.setExtent(layer.boundingBoxOfSelected())
                        canvas.refresh()
                    except Exception:
                        zoom_to_layer(layer)
            else:
                QtWidgets.QMessageBox.information(self, 'Info', f"Aucune entité correspondant à {val} trouvée dans la couche.")
                self.zone_mem_layer = None
                load_layer_to_project(layer, add_if_not=True)
                zoom_to_layer(layer)
        else:
            load_layer_to_project(layer, add_if_not=True)
            if feats:
                try:
                    ids = [f.id() for f in feats]
                    layer.removeSelection()
                    layer.selectByIds(ids)
                    canvas = iface.mapCanvas()
                    canvas.setExtent(layer.boundingBoxOfSelected())
                    canvas.refresh()
                except Exception:
                    zoom_to_layer(layer)
            else:
                zoom_to_layer(layer)

        QtWidgets.QMessageBox.information(self, 'Zone chargée', f"Zone '{val}' chargée et affichée.")

    def _build_page2(self):
        layout = QtWidgets.QVBoxLayout()
        lbl = QtWidgets.QLabel('Page suivante : copie des scripts et ouverture de l\'outil Processing choisi.')
        layout.addWidget(lbl)
        self.page2.setLayout(layout)

    def on_prev(self):
        self.stack.setCurrentIndex(0)
        self.prev_btn.setEnabled(False)
        self.next_btn.setEnabled(True)
        self.open_algo_btn.setEnabled(False)

    def on_next(self):
        prog = self.prog_combo.currentText()
        if not prog:
            QtWidgets.QMessageBox.warning(self, 'Erreur', 'Choisis un programme.')
            return
        if not (self.zone_layer or self.zone_mem_layer):
            QtWidgets.QMessageBox.warning(self, 'Erreur', 'Charge d\'abord la zone d\'étude (page précédente).')
            return

        # prepare optional zonage chosen in combo or via browse
        self.optional_zonage_layer = None
        chosen_data = self.zonage_combo.currentData()
        chosen_text = self.zonage_combo.currentText()
        browse_fp = self.zonage_path_edit.text().strip()

        ref_layer = self.zone_mem_layer or self.zone_layer

        # If user didn't check the show zonage box, ignore zonage inputs as truly optional
        if not self.show_zonage_checkbox.isChecked():
            self.optional_zonage_layer = None
        else:
            # Case 1: user provided a browse file -> use it
            if browse_fp:
                zl_src = QgsVectorLayer(browse_fp, os.path.basename(browse_fp), 'ogr')
                if not zl_src or not zl_src.isValid():
                    QtWidgets.QMessageBox.warning(self, 'Erreur', f"Impossible de charger la couche zonage : {browse_fp}")
                    self.optional_zonage_layer = None
                else:
                    # intersect with ref_layer if present
                    if ref_layer is not None:
                        ref_geoms = [f.geometry() for f in ref_layer.getFeatures()]
                        intersects = []
                        for f in zl_src.getFeatures():
                            try:
                                fg = f.geometry()
                            except Exception:
                                continue
                            ok = False
                            for rg in ref_geoms:
                                try:
                                    if rg is not None and not rg.isEmpty() and fg is not None and not fg.isEmpty():
                                        if rg.intersects(fg):
                                            ok = True
                                            break
                                except Exception:
                                    continue
                            if ok:
                                intersects.append(f)
                        if intersects:
                            mem_zon = create_memory_layer_from_features(zl_src, intersects, name_suffix=f"_INTER_{self.zone_value or ''}")
                            if mem_zon:
                                self.optional_zonage_layer = mem_zon
                            else:
                                QgsProject.instance().addMapLayer(zl_src)
                                self.optional_zonage_layer = zl_src
                        else:
                            QtWidgets.QMessageBox.information(self, 'Info', 'Aucune entité du zonage choisi n\'intersecte la zone d\'étude.')
                            self.optional_zonage_layer = None
                    else:
                        QgsProject.instance().addMapLayer(zl_src)
                        self.optional_zonage_layer = zl_src

            # Case 2: choose from combo (server gpkg path or project layer id)
            elif chosen_data:
                # if chosen_data is a path -> server gpkg
                if isinstance(chosen_data, str) and os.path.exists(chosen_data):
                    # try load gpkg layer (first valid layer)
                    zl_src = try_load_gpkg_layer(chosen_data)
                    if zl_src is None:
                        QtWidgets.QMessageBox.information(self, 'Info', f"Aucune couche utilisable trouvée dans {chosen_data}")
                        self.optional_zonage_layer = None
                    else:
                        if ref_layer is not None:
                            ref_geoms = [f.geometry() for f in ref_layer.getFeatures()]
                            intersects = []
                            for f in zl_src.getFeatures():
                                try:
                                    fg = f.geometry()
                                except Exception:
                                    continue
                                ok = False
                                for rg in ref_geoms:
                                    try:
                                        if rg is not None and not rg.isEmpty() and fg is not None and not fg.isEmpty():
                                            if rg.intersects(fg):
                                                ok = True
                                                break
                                    except Exception:
                                        continue
                                if ok:
                                    intersects.append(f)
                            if intersects:
                                mem_zon = create_memory_layer_from_features(zl_src, intersects, name_suffix=f"_INTER_{self.zone_value or ''}")
                                if mem_zon:
                                    self.optional_zonage_layer = mem_zon
                                else:
                                    load_layer_to_project(zl_src, add_if_not=True)
                                    self.optional_zonage_layer = zl_src
                            else:
                                QtWidgets.QMessageBox.information(self, 'Info', 'Aucune entité du zonage serveur n\'intersecte la zone d\'étude.')
                                self.optional_zonage_layer = None
                        else:
                            load_layer_to_project(zl_src, add_if_not=True)
                            self.optional_zonage_layer = zl_src
                else:
                    # chosen_data is likely a project layer id
                    lyr = QgsProject.instance().mapLayer(chosen_data)
                    if lyr and isinstance(lyr, QgsVectorLayer):
                        # intersect with ref_layer
                        if ref_layer is not None:
                            ref_geoms = [f.geometry() for f in ref_layer.getFeatures()]
                            intersects = []
                            for f in lyr.getFeatures():
                                try:
                                    fg = f.geometry()
                                except Exception:
                                    continue
                                ok = False
                                for rg in ref_geoms:
                                    try:
                                        if rg is not None and not rg.isEmpty() and fg is not None and not fg.isEmpty():
                                            if rg.intersects(fg):
                                                ok = True
                                                break
                                    except Exception:
                                        continue
                                if ok:
                                    intersects.append(f)
                            if intersects:
                                mem_zon = create_memory_layer_from_features(lyr, intersects, name_suffix=f"_INTER_{self.zone_value or ''}")
                                if mem_zon:
                                    self.optional_zonage_layer = mem_zon
                                else:
                                    self.optional_zonage_layer = lyr
                            else:
                                QtWidgets.QMessageBox.information(self, 'Info', 'Aucune entité du zonage choisi n\'intersecte la zone d\'étude.')
                                self.optional_zonage_layer = None
                        else:
                            self.optional_zonage_layer = lyr
                    else:
                        self.optional_zonage_layer = None

        # apply qml if requested and available (for optional zonage)
        if self.optional_zonage_layer is not None and self.qml_zonage_checkbox.isChecked():
            try:
                name = self.optional_zonage_layer.name()
                base = name.split('_INTER_')[0]
                qmlpath = os.path.join(QML_COUCHES_FOLDER, f"QML_{base}.qml")
                if os.path.exists(qmlpath):
                    try:
                        self.optional_zonage_layer.loadNamedStyle(qmlpath)
                        self.optional_zonage_layer.triggerRepaint()
                    except Exception:
                        pass
            except Exception:
                pass

        # ensure scripts copied
        def fb(m):
            print(m)
        ensure_scripts_in_user_folder(feedback=fb)

        # move to page2 and enable open button
        self.stack.setCurrentIndex(1)
        self.prev_btn.setEnabled(True)
        self.next_btn.setEnabled(False)
        self.open_algo_btn.setEnabled(True)
        self.selected_program = prog

    def on_open_algo(self):
        """Ferme le dialogue plugin puis ouvre la fenêtre Processing pour l'algo choisi."""
        if not self.selected_program:
            QtWidgets.QMessageBox.warning(self, 'Erreur', 'Programme non sélectionné.')
            return
        info = ALGO_INFOS.get(self.selected_program)
        if not info:
            QtWidgets.QMessageBox.warning(self, 'Erreur', 'Algorithme non configuré.')
            return
        alg_id = info.get('alg_id')
        alg = QgsApplication.processingRegistry().algorithmById(alg_id)
        if alg is None:
            QtWidgets.QMessageBox.information(self, 'Algorithme manquant',
                f"L'algorithme {alg_id} n'est pas trouvé dans le Toolbox.\n"
                "Nous avons copié les scripts réseau vers ton dossier Processing/scripts utilisateur (si disponible).\n"
                "Si l'algorithme n'apparaît pas, redémarre QGIS.\n\n"
                f"Script source (réseau) : {os.path.join(NETWORK_SCRIPTS_FOLDER, info.get('script_name','-'))}")
            return

        # close/accept dialog, ensure event loop processed, then schedule opening of Processing dialog
        try:
            try:
                self.accept()
            except Exception:
                try:
                    self.close()
                except Exception:
                    pass
            # process pending events
            app = QtWidgets.QApplication.instance()
            if app:
                app.processEvents()

            def _open_proc():
                try:
                    processing.execAlgorithmDialog(alg_id)
                except Exception:
                    try:
                        processing.execAlgorithmDialog(alg_id)
                    except Exception:
                        QtWidgets.QMessageBox.information(None, 'Ouverture manuelle requise',
                            'Impossible d\'ouvrir automatiquement la fenêtre d\'outil Processing pour cet algorithme.\n'
                            'Ouvre manuellement Processing Toolbox et cherche : ' + alg_id)

            QtCore.QTimer.singleShot(150, _open_proc)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, 'Erreur', f"Erreur lors de l'ouverture de l'outil : {e}\n{traceback.format_exc()}")

# ---------------- Plugin entry-point convenience ----------------
class PrelevOrchestratorPlugin:
    def __init__(self, iface):
        self.iface = iface
        self.action = None

    def initGui(self):
        self.action = QtWidgets.QAction('Le VOCAL', self.iface.mainWindow())
        self.action.triggered.connect(self.run)
        self.iface.addPluginToMenu('&Prelev Orchestrator', self.action)
        self.iface.addToolBarIcon(self.action)

    def unload(self):
        if self.action:
            self.iface.removePluginMenu('&Prelev Orchestrator', self.action)
            self.iface.removeToolBarIcon(self.action)

    def run(self):
        dlg = PrelevOrchestratorDialog(iface.mainWindow())
        dlg.exec_()

if __name__ == '__main__':
    try:
        dlg = PrelevOrchestratorDialog()
        dlg.show()
    except Exception:
        print('Run inside QGIS only')
