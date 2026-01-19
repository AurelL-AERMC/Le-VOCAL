# -*- coding: utf-8 -*-
"""
## VOCAL - Version 2.0 (Refactorisée)
Améliorations :
- Interface mono-page fluide
- Mémorisation de la zone d'étude entre sessions
- Code nettoyé (extraction fonctions, logs structurés)
- Validation robuste des inputs
"""

import os
import shutil
import traceback
from qgis.PyQt import QtWidgets, QtCore, QtGui
from qgis.core import (
    QgsApplication, QgsProject, QgsVectorLayer, QgsFeatureRequest,
    QgsWkbTypes, QgsFeature, QgsFields, QgsGeometry, QgsSettings,
    QgsMessageLog, Qgis, QgsSpatialIndex
)
from qgis import processing
from qgis.utils import iface

# ============================================================================
# CONFIGURATION
# ============================================================================

PLUGIN_DIR = os.path.abspath(os.path.dirname(__file__))
BASE_FOLDER = os.path.join(PLUGIN_DIR, 'Couches')
NETWORK_SCRIPTS_FOLDER = os.path.join(PLUGIN_DIR, 'scripts')
QML_COUCHES_FOLDER = os.path.join(PLUGIN_DIR, 'QML')  # Modifié : utilise maintenant ~/QML au lieu de ~/Couches/QML_Couches

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
UNASSIGNED_LABEL = 'Non assigné'

ALGO_INFOS = {
    'Evolution des volumes prélevés par ouvrage': {
        'alg_id': 'script:compute_slopes_ouvrage_only',
        'script_name': 'compute_slopes_qgis_ouvrages.py',
        'qml_name': 'ouvrages_slopes_QML.qml'  # Nouveau : nom du QML associé
    },
    'Evolution des volumes prélevés agrégés par zone': {
        'alg_id': 'script:compute_slopes_zones',
        'script_name': 'compute_slopes_qgis_zonages.py',
        'qml_name': 'zonage_slopes_QML.qml'  # À adapter selon ton QML
    },
    'Ratio VolPrelev/VolAutorise par ouvrage': {
        'alg_id': 'script:compare_prelevements_autorises',
        'script_name': 'compute_ratio_VPVA_ouvrages.py',
        'qml_name': 'ratio_VPVA_ouvrages_QML.qml'  # À adapter
    },
    'Ratio VolPrelev/VolAutorise par zonage': {
        'alg_id': 'script:zones_compare_prelev_autorise',
        'script_name': 'compute_ratio_VPVA_zonages.py',
        'qml_name': 'QML_ratio_VPVA_zonage.qml'  # À adapter
    },
    "État connaissance - ouvrages Agence": {
        'alg_id': 'script:compute_connaissance_ouvrages_agence',
        'script_name': 'compute_connaissance_ouvrages_agence.py',
        'qml_name': 'connaissance_ouvrages_QML.qml'  # À adapter
    }
}

# ============================================================================
# LOGS STRUCTURÉS
# ============================================================================

def log_info(msg):
    QgsMessageLog.logMessage(msg, 'VOCAL', Qgis.Info)

def log_warning(msg):
    QgsMessageLog.logMessage(msg, 'VOCAL', Qgis.Warning)

def log_error(msg):
    QgsMessageLog.logMessage(msg, 'VOCAL', Qgis.Critical)

# ============================================================================
# HELPERS - GESTION GPKG/COUCHES
# ============================================================================

def gpkg_path_for(scale_label):
    fname = GPKG_MAP.get(scale_label)
    return os.path.join(BASE_FOLDER, fname) if fname else None

def try_load_gpkg_layer(gpkg_path):
    """Essaie de charger une couche depuis un GeoPackage."""
    if not gpkg_path or not os.path.exists(gpkg_path):
        return None
    base = os.path.splitext(os.path.basename(gpkg_path))[0]
    candidates = [base, base.lower(), base.upper(), 'departements', 'communes', 'BV', 'bv', 'nappes', 'limite_Deleg']
    for name in candidates:
        uri = f"{gpkg_path}|layername={name}"
        layer = QgsVectorLayer(uri, name, "ogr")
        if layer.isValid():
            return layer
    layer = QgsVectorLayer(gpkg_path, base, "ogr")
    return layer if layer.isValid() else None

def list_zone_values(layer, fieldname):
    if not layer or not fieldname or layer.fields().indexFromName(fieldname) < 0:
        return []
    vals = set()
    for f in layer.getFeatures():
        v = f[fieldname]
        if v is not None:
            vals.add(str(v))
    return sorted(vals)

def load_layer_to_project(layer, add_if_not=True):
    if not layer:
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
    try:
        extent = layer.extent()
        if extent and not extent.isEmpty():
            iface.mapCanvas().setExtent(extent)
            iface.mapCanvas().refresh()
    except Exception:
        pass

def _geom_type_string_from_wkb(wkb):
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
        s = QgsWkbTypes.displayString(wkb).lower()
        if 'polygon' in s:
            return 'Polygon'
        if 'line' in s:
            return 'LineString'
        if 'point' in s:
            return 'Point'
    except Exception:
        pass
    return 'Unknown'

def create_memory_layer_from_features(source_layer, features, name_suffix="_mem", add_to_project=True):
    """Crée une couche mémoire à partir d'une liste de features."""
    if not source_layer or not features:
        return None

    geom_type = _geom_type_string_from_wkb(source_layer.wkbType())
    if geom_type == 'Unknown':
        try:
            test_geom = features[0].geometry()
            if test_geom:
                t = test_geom.type()
                geom_type = {2: 'Polygon', 1: 'LineString', 0: 'Point'}.get(t, 'Polygon')
        except Exception:
            geom_type = 'Polygon'

    crs_auth = source_layer.crs().authid() if source_layer.crs() else ''
    layer_name = f"{source_layer.name()}{name_suffix}"
    
    # IMPORTANT : Vérifier si une couche avec ce nom existe déjà dans le projet
    existing = QgsProject.instance().mapLayersByName(layer_name)
    if existing:
        log_info(f"Couche mémoire existante réutilisée : {layer_name}")
        return existing[0]
    
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
            nf.setAttributes(attrs)
        feats_to_add.append(nf)

    try:
        dp.addFeatures(feats_to_add)
        mem.updateExtents()
        if add_to_project:
            QgsProject.instance().addMapLayer(mem)
        return mem
    except Exception:
        return None

# ============================================================================
# HELPER - INTERSECTION SPATIALE (FONCTION UNIQUE RÉUTILISABLE)
# ============================================================================

def intersect_layer_with_reference(source_layer, ref_layer, zone_value=None):
    """
    Intersecte source_layer avec ref_layer et retourne une couche mémoire.
    
    Args:
        source_layer: couche à filtrer
        ref_layer: couche de référence (zone d'étude)
        zone_value: valeur pour le suffixe du nom (optionnel)
    
    Returns:
        QgsVectorLayer (mémoire) ou None
    """
    if not source_layer or not ref_layer:
        return None
    
    try:
        # Construction index spatial pour performance
        ref_index = QgsSpatialIndex(ref_layer.getFeatures())
        ref_geoms = {}
        for f in ref_layer.getFeatures():
            ref_geoms[f.id()] = f.geometry()
        
        intersects = []
        for f in source_layer.getFeatures():
            try:
                fg = f.geometry()
                if not fg or fg.isEmpty():
                    continue
                
                # Recherche candidates par bbox
                candidates = ref_index.intersects(fg.boundingBox())
                for cid in candidates:
                    rg = ref_geoms.get(cid)
                    if rg and not rg.isEmpty() and rg.intersects(fg):
                        intersects.append(f)
                        break
            except Exception:
                continue
        
        if intersects:
            suffix = f"_INTER_{zone_value}" if zone_value else "_INTER"
            return create_memory_layer_from_features(source_layer, intersects, name_suffix=suffix)
        else:
            log_warning("Aucune entité n'intersecte la zone de référence")
            return None
            
    except Exception as e:
        log_error(f"Erreur intersection spatiale: {e}")
        return None

# ============================================================================
# HELPER - GESTION DES SCRIPTS
# ============================================================================

def ensure_scripts_in_user_folder():
    """Copie les scripts réseau vers le dossier utilisateur ET injecte les chemins QML dynamiques."""
    copied = []
    try:
        user_proc_scripts = os.path.join(QgsApplication.qgisSettingsDirPath(), 'processing', 'scripts')
        os.makedirs(user_proc_scripts, exist_ok=True)
        
        for prog_name, info in ALGO_INFOS.items():
            sn = info.get('script_name')
            qml_name = info.get('qml_name')
            if not sn:
                continue
            
            src = os.path.join(NETWORK_SCRIPTS_FOLDER, sn)
            dst = os.path.join(user_proc_scripts, sn)
            
            if os.path.exists(src):
                try:
                    # Lire le contenu du script source
                    with open(src, 'r', encoding='utf-8') as f:
                        script_content = f.read()
                    
                    # Si un QML est associé, remplacer le chemin par défaut par le chemin dynamique
                    if qml_name:
                        qml_path_dynamic = os.path.join(QML_COUCHES_FOLDER, qml_name)
                        qml_path_normalized = os.path.normpath(qml_path_dynamic).replace('\\', '\\\\')
                        
                        # Pattern pour détecter : default_qml = r"N:\..." ou defaultValue="N:\..."
                        import re
                        # Remplacer les chemins hardcodés par le chemin dynamique
                        script_content = re.sub(
                            r'(default_qml\s*=\s*r?["\']).*?(["\'])',
                            r'\1' + qml_path_normalized + r'\2',
                            script_content
                        )
                        script_content = re.sub(
                            r'(defaultValue\s*=\s*r?["\']).*?(\.qml["\'])',
                            r'\1' + qml_path_normalized + r'\2',
                            script_content
                        )
                    
                    # Vérifier si le fichier destination existe et est différent
                    need_write = True
                    if os.path.exists(dst):
                        try:
                            with open(dst, 'r', encoding='utf-8') as f:
                                existing_content = f.read()
                            if existing_content == script_content:
                                need_write = False
                        except Exception:
                            pass
                    
                    if need_write:
                        with open(dst, 'w', encoding='utf-8') as f:
                            f.write(script_content)
                        log_info(f"Script copié avec chemin QML injecté: {sn}")
                    else:
                        log_info(f"Script déjà à jour: {sn}")
                    
                    copied.append(dst)
                except Exception as e:
                    log_error(f"Erreur copie/injection {sn}: {e}")
            else:
                log_warning(f"Script introuvable: {src}")
    except Exception as e:
        log_error(f"Erreur copie scripts: {e}")
    return copied

# ============================================================================
# HELPER - GESTION QML
# ============================================================================

def apply_qml_to_layer(layer, gpkg_basename):
    """Applique un QML à une couche si le fichier existe."""
    if not layer:
        return False
    try:
        qmlname = f"QML_{os.path.splitext(gpkg_basename)[0]}"
        qmlpath = os.path.join(QML_COUCHES_FOLDER, qmlname + '.qml')
        if os.path.exists(qmlpath):
            layer.loadNamedStyle(qmlpath)
            layer.triggerRepaint()
            return True
    except Exception as e:
        log_warning(f"Erreur application QML: {e}")
    return False

# ============================================================================
# DIALOGUE PRINCIPAL (VERSION MONO-PAGE)
# ============================================================================

class PrelevOrchestratorDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent or iface.mainWindow())
        self.setWindowTitle('VOCAL — Orchestrateur Prélèvements')
        self.resize(950, 700)
        
        # État interne
        self.zone_layer = None
        self.zone_mem_layer = None
        self.zone_value = None
        self.optional_zonage_layer = None
        
        self._build_ui()
        self._load_saved_zone()
        
    def _build_ui(self):
        main_layout = QtWidgets.QVBoxLayout()
        
        # ========== SECTION 1 : PROGRAMME ==========
        grp_prog = QtWidgets.QGroupBox('1. Choix du programme')
        v_prog = QtWidgets.QVBoxLayout()
        self.prog_combo = QtWidgets.QComboBox()
        for k in ALGO_INFOS.keys():
            self.prog_combo.addItem(k)
        v_prog.addWidget(self.prog_combo)
        grp_prog.setLayout(v_prog)
        
        # ========== SECTION 2 : ZONE D'ÉTUDE (COLLAPSIBLE) ==========
        self.grp_zone = QtWidgets.QGroupBox("2. Zone d'étude")
        self.grp_zone.setCheckable(True)
        self.grp_zone.setChecked(False)  # Replié par défaut si zone en mémoire
        
        zone_layout = QtWidgets.QGridLayout()
        
        # Indicateur zone active
        self.zone_status_label = QtWidgets.QLabel()
        self.zone_status_label.setStyleSheet("color: green; font-weight: bold;")
        zone_layout.addWidget(self.zone_status_label, 0, 0, 1, 2)
        
        zone_layout.addWidget(QtWidgets.QLabel('Échelle'), 1, 0)
        self.scale_combo = QtWidgets.QComboBox()
        self.scale_combo.addItems(list(GPKG_MAP.keys()))
        self.scale_combo.currentTextChanged.connect(self._on_scale_changed)
        zone_layout.addWidget(self.scale_combo, 1, 1)
        
        zone_layout.addWidget(QtWidgets.QLabel('Valeur'), 2, 0)
        self.zone_value_combo = QtWidgets.QComboBox()
        zone_layout.addWidget(self.zone_value_combo, 2, 1)
        
        self.load_zone_btn = QtWidgets.QPushButton('📍 Charger la zone et zoomer')
        self.load_zone_btn.clicked.connect(self._on_load_zone)
        zone_layout.addWidget(self.load_zone_btn, 3, 0, 1, 2)
        
        self.create_memory_checkbox = QtWidgets.QCheckBox("Créer couche mémoire limitée (recommandé)")
        self.create_memory_checkbox.setChecked(True)
        zone_layout.addWidget(self.create_memory_checkbox, 4, 0, 1, 2)
        
        self.grp_zone.setLayout(zone_layout)
        
        # ========== SECTION 3 : ZONAGE OPTIONNEL ==========
        self.show_zonage_checkbox = QtWidgets.QCheckBox("3. Charger un sous-zonage (optionnel)")
        self.show_zonage_checkbox.setChecked(False)
        
        self.grp_zonage = QtWidgets.QGroupBox('Configuration du sous-zonage')
        zonage_layout = QtWidgets.QGridLayout()
        
        zonage_layout.addWidget(QtWidgets.QLabel("Source (serveur/projet)"), 0, 0)
        self.zonage_combo = QtWidgets.QComboBox()
        self._populate_zonage_combo()
        zonage_layout.addWidget(self.zonage_combo, 0, 1)
        
        zonage_layout.addWidget(QtWidgets.QLabel("Ou fichier externe"), 1, 0)
        self.zonage_path_edit = QtWidgets.QLineEdit()
        self.zonage_browse = QtWidgets.QPushButton('📂 Parcourir')
        self.zonage_browse.clicked.connect(self._on_browse_zonage)
        h_browse = QtWidgets.QHBoxLayout()
        h_browse.addWidget(self.zonage_path_edit)
        h_browse.addWidget(self.zonage_browse)
        zonage_layout.addLayout(h_browse, 1, 1)
        
        self.grp_zonage.setLayout(zonage_layout)
        self.grp_zonage.setVisible(False)
        self.show_zonage_checkbox.toggled.connect(lambda c: self.grp_zonage.setVisible(c))
        
        # ========== SECTION 4 : OPTIONS QML ==========
        grp_qml = QtWidgets.QGroupBox('4. Styles visuels (QML)')
        qml_layout = QtWidgets.QVBoxLayout()
        self.qml_zone_checkbox = QtWidgets.QCheckBox("Appliquer QML à la zone d'étude (recommmandé)")
        self.qml_zonage_checkbox = QtWidgets.QCheckBox("Appliquer QML au sous-zonage (recommmandé)")
        self.qml_zone_checkbox.setChecked(True)
        self.qml_zonage_checkbox.setChecked(True)
        qml_layout.addWidget(self.qml_zone_checkbox)
        qml_layout.addWidget(self.qml_zonage_checkbox)
        grp_qml.setLayout(qml_layout)
        
        # ========== SECTION 5 : INFO ==========
        info_label = QtWidgets.QLabel(
            '💡 <b>Astuce</b> : La zone d\'étude est mémorisée entre les sessions.<br>'
            'Si une zone est active, tu peux directement lancer le programme sans la recharger.'
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("background-color: #e8f4f8; padding: 10px; border-radius: 5px;")
        
        # ========== BOUTONS D'ACTION ==========
        btn_layout = QtWidgets.QHBoxLayout()
        btn_layout.addStretch()
        
        self.clear_zone_btn = QtWidgets.QPushButton('🗑️ Effacer zone')
        self.clear_zone_btn.clicked.connect(self._clear_zone)
        
        self.launch_btn = QtWidgets.QPushButton('🚀 Valider les paramètres')
        self.launch_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        self.launch_btn.clicked.connect(self._on_launch)
        
        btn_layout.addWidget(self.clear_zone_btn)
        btn_layout.addWidget(self.launch_btn)
        
        # ========== ASSEMBLAGE ==========
        main_layout.addWidget(grp_prog)
        main_layout.addWidget(self.grp_zone)
        main_layout.addWidget(self.show_zonage_checkbox)
        main_layout.addWidget(self.grp_zonage)
        main_layout.addWidget(grp_qml)
        main_layout.addWidget(info_label)
        main_layout.addStretch()
        main_layout.addLayout(btn_layout)
        
        self.setLayout(main_layout)
        self._on_scale_changed(self.scale_combo.currentText())
        
    def _populate_zonage_combo(self):
        self.zonage_combo.clear()
        try:
            for fname in os.listdir(BASE_FOLDER):
                if fname.lower().endswith('.gpkg'):
                    full = os.path.join(BASE_FOLDER, fname)
                    self.zonage_combo.addItem(f"[srv] {fname}", full)
        except Exception:
            pass
        self.zonage_combo.addItem("──── Couches projet ────", None)
        for lyr in QgsProject.instance().mapLayers().values():
            if isinstance(lyr, QgsVectorLayer):
                self.zonage_combo.addItem(f"[proj] {lyr.name()}", lyr.id())
    
    def _on_scale_changed(self, text):
        gpkg = gpkg_path_for(text)
        layer = try_load_gpkg_layer(gpkg)
        self.zone_value_combo.clear()
        
        if not layer:
            self.zone_value_combo.addItem('-- couche introuvable --')
            return
        
        if text == 'Départements':
            field = DEPT_FIELD
        elif text == 'Bassins versants':
            field = BV_FIELD
        else:
            field = 'name' if layer.fields().indexFromName('name') >= 0 else layer.fields()[0].name()
        
        vals = list_zone_values(layer, field)
        if vals:
            self.zone_value_combo.addItems(vals)
        else:
            self.zone_value_combo.addItem('-- Aucun attribut --')
    
    def _on_browse_zonage(self):
        fp, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, 'Choisir zonage', BASE_FOLDER, 
            'GeoPackage (*.gpkg);;Shapefile (*.shp);;All (*)'
        )
        if fp:
            self.zonage_path_edit.setText(fp)
    
    def _on_load_zone(self):
        scale = self.scale_combo.currentText()
        val = self.zone_value_combo.currentText()
        gpkg = gpkg_path_for(scale)
        layer = try_load_gpkg_layer(gpkg)
        
        if not layer:
            QtWidgets.QMessageBox.warning(self, 'Erreur', f'Impossible de charger {gpkg}.')
            return
        
        if scale == 'Départements':
            field = DEPT_FIELD
        elif scale == 'Bassins versants':
            field = BV_FIELD
        else:
            field = 'name' if layer.fields().indexFromName('name') >= 0 else layer.fields()[0].name()
        
        expr = f'"{field}" = \'{val}\''
        try:
            it = layer.getFeatures(QgsFeatureRequest().setFilterExpression(expr))
            feats = [f for f in it]
        except Exception:
            feats = []
        
        self.zone_layer = layer
        self.zone_value = val
        
        if self.create_memory_checkbox.isChecked() and feats:
            # create_memory_layer_from_features vérifie maintenant en interne si la couche existe déjà
            mem = create_memory_layer_from_features(layer, feats, name_suffix=f"_INTER_{val}")
            if mem:
                self.zone_mem_layer = mem
                if self.qml_zone_checkbox.isChecked():
                    apply_qml_to_layer(mem, os.path.basename(gpkg or ''))
                zoom_to_layer(mem)
            else:
                self._load_full_zone(layer, feats)
        else:
            self._load_full_zone(layer, feats)
        
        self._save_zone_to_settings()
        self._update_zone_status()
        QtWidgets.QMessageBox.information(self, 'Zone chargée', f"Zone '{val}' chargée avec succès !")
    
    def _load_full_zone(self, layer, feats):
        load_layer_to_project(layer, add_if_not=True)
        if feats:
            try:
                ids = [f.id() for f in feats]
                layer.selectByIds(ids)
                iface.mapCanvas().setExtent(layer.boundingBoxOfSelected())
                iface.mapCanvas().refresh()
            except Exception:
                zoom_to_layer(layer)
        else:
            zoom_to_layer(layer)
    
    def _clear_zone(self):
        self.zone_layer = None
        self.zone_mem_layer = None
        self.zone_value = None
        settings = QgsSettings()
        settings.remove('vocal/last_zone_scale')
        settings.remove('vocal/last_zone_value')
        self._update_zone_status()
        QtWidgets.QMessageBox.information(self, 'Zone effacée', 'La zone mémorisée a été supprimée.')
    
    def _save_zone_to_settings(self):
        if self.zone_value:
            settings = QgsSettings()
            settings.setValue('vocal/last_zone_scale', self.scale_combo.currentText())
            settings.setValue('vocal/last_zone_value', self.zone_value)
            log_info(f"Zone sauvegardée: {self.zone_value}")
    
    def _load_saved_zone(self):
        settings = QgsSettings()
        scale = settings.value('vocal/last_zone_scale')
        val = settings.value('vocal/last_zone_value')
        
        if scale and val:
            # Restaurer la sélection dans les combos
            idx = self.scale_combo.findText(scale)
            if idx >= 0:
                self.scale_combo.setCurrentIndex(idx)
            idx_val = self.zone_value_combo.findText(val)
            if idx_val >= 0:
                self.zone_value_combo.setCurrentIndex(idx_val)
            
            # Chercher d'abord si une couche avec ce nom existe déjà dans le projet
            gpkg = gpkg_path_for(scale)
            if gpkg:
                base_layer = try_load_gpkg_layer(gpkg)
                if base_layer:
                    expected_mem_name = f"{base_layer.name()}_INTER_{val}"
                    existing_layers = QgsProject.instance().mapLayersByName(expected_mem_name)
                    
                    if existing_layers:
                        # Réutiliser la couche existante (pas de rechargement)
                        self.zone_mem_layer = existing_layers[0]
                        self.zone_layer = base_layer
                        self.zone_value = val
                        log_info(f"Zone existante réutilisée : {val}")
                    else:
                        # La couche n'existe pas dans le projet -> on ne fait rien
                        # L'utilisateur devra recharger manuellement si besoin
                        log_info(f"Zone sauvegardée trouvée ({val}) mais couche absente du projet - chargement manuel requis")
                        self.zone_value = val
        
        self._update_zone_status()
    
    def _update_zone_status(self):
        if self.zone_mem_layer or self.zone_layer:
            txt = f"✅ Zone active : {self.zone_value or 'inconnue'}"
            self.zone_status_label.setText(txt)
            self.grp_zone.setChecked(False)  # Replier si zone déjà chargée
        else:
            self.zone_status_label.setText("⚠️ Aucune zone active (charge une zone ci-dessous)")
            self.zone_status_label.setStyleSheet("color: orange; font-weight: bold;")
            self.grp_zone.setChecked(True)  # Déplier pour forcer le chargement
    
    def _prepare_zonage(self):
        """Prépare le zonage optionnel (intersecté avec zone d'étude)."""
        if not self.show_zonage_checkbox.isChecked():
            self.optional_zonage_layer = None
            return
        
        ref_layer = self.zone_mem_layer or self.zone_layer
        browse_fp = self.zonage_path_edit.text().strip()
        chosen_data = self.zonage_combo.currentData()
        
        # Cas 1 : fichier externe
        if browse_fp:
            zl_src = QgsVectorLayer(browse_fp, os.path.basename(browse_fp), 'ogr')
            if not zl_src.isValid():
                QtWidgets.QMessageBox.warning(self, 'Erreur', f"Impossible de charger: {browse_fp}")
                self.optional_zonage_layer = None
                return
            self.optional_zonage_layer = intersect_layer_with_reference(zl_src, ref_layer, self.zone_value)
        
        # Cas 2 : combo (serveur ou projet)
        elif chosen_data:
            if isinstance(chosen_data, str) and os.path.exists(chosen_data):
                # GPKG serveur
                zl_src = try_load_gpkg_layer(chosen_data)
                if zl_src:
                    self.optional_zonage_layer = intersect_layer_with_reference(zl_src, ref_layer, self.zone_value)
            else:
                # Couche projet
                lyr = QgsProject.instance().mapLayer(chosen_data)
                if lyr and isinstance(lyr, QgsVectorLayer):
                    self.optional_zonage_layer = intersect_layer_with_reference(lyr, ref_layer, self.zone_value)
        
        # Appliquer QML si demandé
        if self.optional_zonage_layer and self.qml_zonage_checkbox.isChecked():
            try:
                name = self.optional_zonage_layer.name()
                base = name.split('_INTER_')[0]
                apply_qml_to_layer(self.optional_zonage_layer, base)
            except Exception:
                pass
    
    def _on_launch(self):
        prog = self.prog_combo.currentText()
        if not prog:
            QtWidgets.QMessageBox.warning(self, 'Erreur', 'Choisis un programme.')
            return
        
        if not (self.zone_layer or self.zone_mem_layer):
            QtWidgets.QMessageBox.warning(
                self, 'Erreur', 
                'Charge d\'abord une zone d\'étude (section 2).'
            )
            return
        
        # Préparer zonage optionnel
        self._prepare_zonage()
        
        # Copier scripts
        ensure_scripts_in_user_folder()
        
        # Récupérer l'algo
        info = ALGO_INFOS.get(prog)
        if not info:
            QtWidgets.QMessageBox.warning(self, 'Erreur', 'Algorithme non configuré.')
            return
        
        alg_id = info.get('alg_id')
        alg = QgsApplication.processingRegistry().algorithmById(alg_id)
        
        if not alg:
            QtWidgets.QMessageBox.information(
                self, 'Algorithme manquant',
                f"L'algorithme {alg_id} n'est pas trouvé.\n"
                "Scripts copiés vers Processing/scripts. Redémarre QGIS ou rafraîchis le Toolbox."
            )
            return
        
        # Fermer et lancer Processing
        try:
            self.accept()
            QtWidgets.QApplication.instance().processEvents()
            
            def _open():
                try:
                    processing.execAlgorithmDialog(alg_id)
                except Exception:
                    QtWidgets.QMessageBox.information(
                        None, 'Ouverture manuelle',
                        f'Ouvre manuellement le Toolbox et cherche: {alg_id}'
                    )
            
            QtCore.QTimer.singleShot(150, _open)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, 'Erreur', f"Erreur ouverture: {e}\n{traceback.format_exc()}")

# ============================================================================
# CLASSE PLUGIN (ENTRY POINT)
# ============================================================================

class PrelevOrchestratorPlugin:
    def __init__(self, iface):
        self.iface = iface
        self.action = None
    
    def initGui(self):
        icon_path = os.path.join(PLUGIN_DIR, 'icon.png')
        qicon = None
        if os.path.exists(icon_path):
            try:
                qicon = QtGui.QIcon(icon_path)
            except Exception:
                pass
        
        if qicon and not qicon.isNull():
            self.action = QtWidgets.QAction(qicon, '', self.iface.mainWindow())
            self.action.setToolTip('Le VOCAL')
        else:
            self.action = QtWidgets.QAction('VOCAL', self.iface.mainWindow())
        
        self.action.triggered.connect(self.run)
        self.iface.addPluginToMenu('&VOCAL', self.action)
        self.iface.addToolBarIcon(self.action)
    
    def unload(self):
        if self.action:
            try:
                self.iface.removePluginMenu('&VOCAL', self.action)
                self.iface.removeToolBarIcon(self.action)
            except Exception:
                pass
    
    def run(self):
        dlg = PrelevOrchestratorDialog(iface.mainWindow())
        dlg.exec_()

if __name__ == '__main__':
    try:
        dlg = PrelevOrchestratorDialog()
        dlg.show()
    except Exception:
        print('Run inside QGIS only')