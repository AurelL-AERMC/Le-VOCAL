# -*- coding: utf-8 -*-
"""
Processing script QGIS : agrégation zonale — comparer volumes prélevés vs volumes autorisés
(modification : protection pour ouvrages non assignés -> agrégés sous "Non assigné" sans géométrie)
- Entrées : couche zonage (polygones) + champ libellé, prélèvements (points/table), volumes autorisés (table)
- Pour une année donnée : agrège prélèvements par ouvrage, joint avec autorisés (MAX si multiples),
  garde uniquement ouvrages appariés, affecte aux zones (multi-affectation possible),
  somme prélevé et autorisé par zone et calcule ratio / pourcentages.
- Option d'appliquer un QML sur la couche de sortie.
"""
from qgis.PyQt.QtCore import QVariant
from qgis.core import (
    QgsProcessing,
    QgsProcessingAlgorithm,
    QgsProcessingParameterVectorLayer,
    QgsProcessingParameterField,
    QgsProcessingParameterNumber,
    QgsProcessingParameterFeatureSink,
    QgsProcessingParameterBoolean,
    QgsProcessingParameterString,
    QgsFeature,
    QgsField,
    QgsFields,
    QgsProject,
    QgsProcessingUtils,
    QgsSpatialIndex,
    QgsFeatureSink   # <-- import ajouté pour éviter NameError
)
import re
import os
import math
from collections import defaultdict

# label utilisé pour agréger les ouvrages non assignés à une zone
UNASSIGNED_LABEL = 'Non assigné'

# ---------- utilitaires ----------
def parse_number(x):
    """Parse un nombre au format français/anglais -> float ou NaN"""
    if x is None:
        return float('nan')
    if isinstance(x, (int, float)):
        try:
            return float(x)
        except:
            return float('nan')
    s = str(x).strip()
    if s == '':
        return float('nan')
    s = s.replace('\xa0', ' ')
    s_nosp = s.replace(' ', '')
    if '.' in s_nosp and ',' in s_nosp:
        if s_nosp.find('.') < s_nosp.find(','):
            s_clean = s_nosp.replace('.', '').replace(',', '.')
        else:
            s_clean = s_nosp.replace(',', '')
    elif ',' in s_nosp:
        s_clean = s_nosp.replace(',', '.')
    else:
        s_clean = s_nosp
    s_clean = re.sub(r'[^0-9\.\-]', '', s_clean)
    if s_clean in ['', '.', '-', '-.']:
        return float('nan')
    try:
        return float(s_clean)
    except:
        return float('nan')

def parse_year_to_int(y_raw):
    """Convertit une valeur d'année en int si possible (gère string contenant '2023', etc.)"""
    if y_raw is None:
        return None
    if isinstance(y_raw, int):
        return y_raw
    if isinstance(y_raw, float):
        try:
            return int(y_raw)
        except:
            pass
    s = str(y_raw).strip()
    if s == '':
        return None
    try:
        return int(float(s))
    except:
        pass
    m = re.search(r'(\d{4})', s)
    if m:
        try:
            return int(m.group(1))
        except:
            return None
    return None

# ---------- Algorithm ----------
class ZonesComparePrelevAutorise(QgsProcessingAlgorithm):
    """
    Agrège par zones : prélèvements vs volumes autorisés (année unique).
    """

    # paramètres
    ZONES = 'ZONES'
    ZONE_LABEL = 'ZONE_LABEL'
    PRELEV = 'PRELEV'
    PRELEV_YEAR = 'PRELEV_YEAR'
    PRELEV_OUV = 'PRELEV_OUV'
    PRELEV_ASSIETTE = 'PRELEV_ASSIETTE'
    AUTOR = 'AUTOR'
    AUTOR_OUV = 'AUTOR_OUV'
    AUTOR_VOL = 'AUTOR_VOL'
    AUTOR_DDTM = 'AUTOR_DDTM'
    YEAR = 'YEAR'
    APPLY_QML = 'APPLY_QML'
    QML_PATH = 'QML_PATH'
    OUTPUT = 'OUTPUT'

    def tr(self, s):
        return s

    def createInstance(self):
        return ZonesComparePrelevAutorise()

    def name(self):
        return 'zones_compare_prelev_autorise'

    def displayName(self):
        return self.tr('Zonage : comparer prélevés vs autorisés (année unique)')

    def group(self):
        return self.tr('Analyses temporelles')

    def groupId(self):
        return 'temporal_analysis'

    def shortHelpString(self):
        return self.tr(
            "Pour une année donnée, agrège les prélèvements par ouvrage, joint avec les volumes autorisés (MAX si multiples), "
            "garde uniquement les ouvrages appariés, affecte aux zones (multi-affectation possible), "
            "somme prélevé et autorisé par zone et calcule ratio / pourcentages. Les ouvrages non-intersectés sont "
            "agrégés sous '{}' sans géométrie.".format(UNASSIGNED_LABEL)
        )

    def initAlgorithm(self, config=None):
        # zones
        self.addParameter(
            QgsProcessingParameterVectorLayer(self.ZONES, self.tr("Couche de zonage (polygones)"), [QgsProcessing.TypeVectorAnyGeometry])
        )
        self.addParameter(
            QgsProcessingParameterField(self.ZONE_LABEL, self.tr("Champ libellé du zonage (sera conservé)"), parentLayerParameterName=self.ZONES)
        )
        # prélèvements
        self.addParameter(
            QgsProcessingParameterVectorLayer(self.PRELEV, self.tr("Couche prélèvements (points/table)"), [QgsProcessing.TypeVectorAnyGeometry])
        )
        self.addParameter(
            QgsProcessingParameterField(self.PRELEV_YEAR, self.tr("Champ année (prélèvements)"), parentLayerParameterName=self.PRELEV, type=QgsProcessingParameterField.Any)
        )
        self.addParameter(
            QgsProcessingParameterField(self.PRELEV_OUV, self.tr("Champ ID Ouvrage (prélèvements)"), parentLayerParameterName=self.PRELEV)
        )
        self.addParameter(
            QgsProcessingParameterField(self.PRELEV_ASSIETTE, self.tr("Champ Assiette (volume prélevé)"), parentLayerParameterName=self.PRELEV)
        )
        # volumes autorisés
        self.addParameter(
            QgsProcessingParameterVectorLayer(self.AUTOR, self.tr("Table / couche volumes autorisés"), [QgsProcessing.TypeVectorAnyGeometry])
        )
        self.addParameter(
            QgsProcessingParameterField(self.AUTOR_OUV, self.tr("Champ ID Ouvrage (autorises) - pour la jointure"), parentLayerParameterName=self.AUTOR)
        )
        self.addParameter(
            QgsProcessingParameterField(self.AUTOR_VOL, self.tr("Champ Volume autorisé (autorises)"), parentLayerParameterName=self.AUTOR)
        )
        self.addParameter(
            QgsProcessingParameterField(self.AUTOR_DDTM, self.tr("Champ Identifiant DDTM (autorises) - optionnel"), parentLayerParameterName=self.AUTOR, optional=True)
        )
        # autres
        self.addParameter(
            QgsProcessingParameterNumber(self.YEAR, self.tr("Année (ex : 2023)"), type=QgsProcessingParameterNumber.Integer, defaultValue=2023)
        )
        self.addParameter(
            QgsProcessingParameterBoolean(self.APPLY_QML, self.tr("Appliquer un style QML sur la couche de sortie ?"), defaultValue=True)
        )
        default_qml = r"N:\_MTP\Public\01-ORGANISATION\G-Services\RAGAF\REDEVANCES\Recherche redevables\Etude données prélèvements\06_Valorisation_Visualisation\Outils\QML\QML_ratio_VPVA_zonage.qml"
        self.addParameter(
            QgsProcessingParameterString(self.QML_PATH, self.tr("Chemin du fichier QML (si appliqué)"), defaultValue=default_qml)
        )
        self.addParameter(
            QgsProcessingParameterFeatureSink(self.OUTPUT, self.tr("Couche de sortie (zones enrichies)"))
        )

    def processAlgorithm(self, parameters, context, feedback):
        # lire paramètres
        zones_lyr = self.parameterAsVectorLayer(parameters, self.ZONES, context)
        zone_label_field = self.parameterAsString(parameters, self.ZONE_LABEL, context)

        prelev_lyr = self.parameterAsVectorLayer(parameters, self.PRELEV, context)
        prelev_year_field = self.parameterAsString(parameters, self.PRELEV_YEAR, context)
        prelev_ouv_field = self.parameterAsString(parameters, self.PRELEV_OUV, context)
        prelev_assiette_field = self.parameterAsString(parameters, self.PRELEV_ASSIETTE, context)

        autor_lyr = self.parameterAsVectorLayer(parameters, self.AUTOR, context)
        autor_ouv_field = self.parameterAsString(parameters, self.AUTOR_OUV, context)
        autor_vol_field = self.parameterAsString(parameters, self.AUTOR_VOL, context)
        autor_ddtm_field = self.parameterAsString(parameters, self.AUTOR_DDTM, context) if self.AUTOR_DDTM in parameters else None

        year_param = int(self.parameterAsInt(parameters, self.YEAR, context))
        apply_qml = bool(self.parameterAsBool(parameters, self.APPLY_QML, context))
        qml_path_param = self.parameterAsString(parameters, self.QML_PATH, context)

        feedback.pushInfo(self.tr(f"Paramètres : année={year_param}"))

        # ---------- 1) Index des volumes autorisés (par ouvrage) ----------
        # Prendre MAX(volume autorisé) si plusieurs enregistrements, concaténer DDTM distincts
        autor_index = {}  # key -> {'vol_max': float or NaN, 'ddtm': set()}
        n_autor = 0
        for f in autor_lyr.getFeatures():
            n_autor += 1
            k_raw = f[autor_ouv_field]
            if k_raw is None:
                continue
            k = str(k_raw).strip()
            vol_raw = f[autor_vol_field]
            vol = parse_number(vol_raw)
            ddtm_val = None
            if autor_ddtm_field:
                try:
                    ddtm_val = f[autor_ddtm_field]
                    if ddtm_val is not None:
                        ddtm_val = str(ddtm_val).strip()
                except Exception:
                    ddtm_val = None
            ent = autor_index.get(k)
            if ent is None:
                s = set()
                if ddtm_val:
                    s.add(ddtm_val)
                autor_index[k] = {'vol_max': vol if not math.isnan(vol) else float('nan'), 'ddtm': s}
            else:
                try:
                    if not math.isnan(vol):
                        if math.isnan(ent['vol_max']) or vol > ent['vol_max']:
                            ent['vol_max'] = vol
                except Exception:
                    pass
                if ddtm_val:
                    ent['ddtm'].add(ddtm_val)
        feedback.pushInfo(self.tr(f"Index volumes autorisés : {len(autor_index)} clés construites (parcours {n_autor} enregistrements)."))

        # ---------- 2) Parcourir prélèvements pour l'année, agréger par ouvrage ----------
        assiette_by_ouv = defaultdict(float)
        geom_by_ouv = {}
        prelev_count = 0
        skipped_year = 0
        for f in prelev_lyr.getFeatures():
            prelev_count += 1
            if feedback.isCanceled():
                break
            y_raw = f[prelev_year_field]
            y_int = parse_year_to_int(y_raw)
            if y_int is None:
                skipped_year += 1
                continue
            if y_int != year_param:
                continue
            id_raw = f[prelev_ouv_field]
            if id_raw is None:
                continue
            key = str(id_raw).strip()
            ass_raw = f[prelev_assiette_field]
            ass = parse_number(ass_raw)
            ass_val = 0.0 if math.isnan(ass) else ass
            assiette_by_ouv[key] += ass_val
            # conserver premier point rencontré comme géométrie (pour affectation spatiale)
            if prelev_lyr.geometryType() != -1 and key not in geom_by_ouv:
                geom = f.geometry()
                if geom and not geom.isEmpty():
                    geom_by_ouv[key] = geom
            if prelev_count % 500 == 0:
                feedback.setProgress(int(100 * prelev_count / max(1, prelev_lyr.featureCount())))
        feedback.pushInfo(self.tr(f"Prélèvements parcourus: {prelev_count}, ignorés (année non parsable): {skipped_year}, ouvrages agrégés: {len(assiette_by_ouv)}"))

        # ---------- 3) Conserver uniquement ouvrages qui ont une entrée autorisée (jointure possible) ----------
        matched_ouvrages = {}
        for k, ass_sum in assiette_by_ouv.items():
            autor_ent = autor_index.get(k)
            if autor_ent is None:
                # on exclut les non appariés (consigne)
                continue
            vol_auth = autor_ent.get('vol_max')
            if vol_auth is None or (isinstance(vol_auth, float) and math.isnan(vol_auth)):
                vol_auth = None
            ddtm_concat = ';'.join(sorted(autor_ent['ddtm'])) if autor_ent['ddtm'] else None
            matched_ouvrages[k] = {'assiette': ass_sum, 'vol_autorise': vol_auth, 'ddtm': ddtm_concat, 'geom': geom_by_ouv.get(k)}

        if not matched_ouvrages:
            raise Exception(self.tr("Aucun ouvrage apparié aux volumes autorisés pour l'année et les données fournies."))

        feedback.pushInfo(self.tr(f"Ouvrages appariés retenus : {len(matched_ouvrages)} (les non-appariés ont été exclus)."))

        # ---------- 4) Affectation spatiale : ouvrages -> zones (multi-affectation : toutes les zones intersectées)
        feedback.pushInfo(self.tr("Création d'un index spatial des zones..."))
        zone_index = QgsSpatialIndex(zones_lyr.getFeatures())
        # map fid -> label (et geometry feature)
        fid2label = {}
        fid2feat = {}
        for zf in zones_lyr.getFeatures():
            fid2label[zf.id()] = zf[zone_label_field]
            fid2feat[zf.id()] = zf

        zone_prelev_sum = defaultdict(float)   # key zone_label -> sum prelev
        zone_autor_sum = defaultdict(float)    # key zone_label -> sum autor
        zone_count_ouvrages = defaultdict(int)
        # ensure unassigned key exists
        zone_prelev_sum[UNASSIGNED_LABEL] = 0.0
        zone_autor_sum[UNASSIGNED_LABEL] = 0.0
        zone_count_ouvrages[UNASSIGNED_LABEL] = 0

        # for diagnostics
        iter_n = 0
        for k, info in matched_ouvrages.items():
            iter_n += 1
            if feedback.isCanceled():
                break
            geom = info.get('geom')
            assigned_any = False
            if geom is None or geom.isEmpty():
                # no geometry -> directly count as non-assigned
                zone_prelev_sum[UNASSIGNED_LABEL] += info['assiette'] if info['assiette'] is not None else 0.0
                if info['vol_autorise'] is not None and not (isinstance(info['vol_autorise'], float) and math.isnan(info['vol_autorise'])):
                    zone_autor_sum[UNASSIGNED_LABEL] += info['vol_autorise']
                zone_count_ouvrages[UNASSIGNED_LABEL] += 1
                assigned_any = False
            else:
                # find candidate zone fids by bbox
                candidates = zone_index.intersects(geom.boundingBox())
                for fid in candidates:
                    zfeat = fid2feat.get(fid)
                    if zfeat is None:
                        continue
                    try:
                        if zfeat.geometry().intersects(geom):
                            label = fid2label.get(fid)
                            # accumulate sums per label (string)
                            zone_prelev_sum[label] += info['assiette'] if info['assiette'] is not None else 0.0
                            if info['vol_autorise'] is not None and not (isinstance(info['vol_autorise'], float) and math.isnan(info['vol_autorise'])):
                                zone_autor_sum[label] += info['vol_autorise']
                            zone_count_ouvrages[label] += 1
                            assigned_any = True
                    except Exception:
                        continue
                if not assigned_any:
                    # intersects no zone -> aggregate under UNASSIGNED_LABEL
                    zone_prelev_sum[UNASSIGNED_LABEL] += info['assiette'] if info['assiette'] is not None else 0.0
                    if info['vol_autorise'] is not None and not (isinstance(info['vol_autorise'], float) and math.isnan(info['vol_autorise'])):
                        zone_autor_sum[UNASSIGNED_LABEL] += info['vol_autorise']
                    zone_count_ouvrages[UNASSIGNED_LABEL] += 1
            if iter_n % 200 == 0:
                feedback.setProgress(int(100 * iter_n / max(1, len(matched_ouvrages))))
        feedback.pushInfo(self.tr("Affectation spatiale terminée. Les ouvrages sans intersection ont été agrégés sous '{}'.".format(UNASSIGNED_LABEL)))

        # ---------- 5) Calculs par zone : ratio, pourcentage, etc. ----------
        zone_list = sorted(set(list(zone_prelev_sum.keys()) + list(zone_autor_sum.keys())))
        if not zone_list:
            raise Exception(self.tr("Aucune zone n'a reçu d'agrégats — vérifie intersections et géométries."))

        zone_ratio = {}
        zone_ratio_possible = {}
        zone_percent_prelev_auth = {}
        zone_percent_overrun = {}
        for z in zone_list:
            prelev = zone_prelev_sum.get(z, 0.0)
            autor = zone_autor_sum.get(z)
            if autor is None or (isinstance(autor, float) and math.isnan(autor)) or autor == 0:
                zone_ratio_possible[z] = 0
                zone_ratio[z] = None
                zone_percent_prelev_auth[z] = None
                zone_percent_overrun[z] = None
            else:
                try:
                    r = prelev / autor
                    zone_ratio[z] = r
                    zone_ratio_possible[z] = 1
                    zone_percent_prelev_auth[z] = r * 100.0
                    zone_percent_overrun[z] = ((prelev - autor) / autor) * 100.0
                except Exception:
                    zone_ratio_possible[z] = 0
                    zone_ratio[z] = None
                    zone_percent_prelev_auth[z] = None
                    zone_percent_overrun[z] = None

        # ---------- 6) Préparer sink (couche de sortie = géométrie des polygones d'entrée + feature Non assigné sans géométrie) ----------
        out_fields = QgsFields()
        out_fields.append(QgsField(zone_label_field, QVariant.String))
        out_fields.append(QgsField('prelev_sum', QVariant.Double))
        out_fields.append(QgsField('autor_sum', QVariant.Double))
        out_fields.append(QgsField('ratio', QVariant.Double))
        out_fields.append(QgsField('ratio_possible', QVariant.Int))
        out_fields.append(QgsField('percent_prelev_auth', QVariant.Double))
        out_fields.append(QgsField('percent_overrun', QVariant.Double))
        out_fields.append(QgsField('n_ouvrages', QVariant.Int))

        (sink, dest_id) = self.parameterAsSink(parameters, self.OUTPUT, context,
                                               out_fields,
                                               zones_lyr.wkbType(), zones_lyr.sourceCrs())

        # écrire : parcourir les features de zones et ajouter champs correspondants (pour conserver géométrie)
        total_z = zones_lyr.featureCount()
        cnt = 0
        for zf in zones_lyr.getFeatures():
            if feedback.isCanceled():
                break
            label = zf[zone_label_field]
            prelev = zone_prelev_sum.get(label, 0.0)
            autor = zone_autor_sum.get(label)
            r = zone_ratio.get(label)
            rpos = int(zone_ratio_possible.get(label, 0))
            ppre = zone_percent_prelev_auth.get(label)
            pover = zone_percent_overrun.get(label)
            n_ouv = zone_count_ouvrages.get(label, 0)
            feat = QgsFeature()
            feat.setFields(out_fields)
            try:
                feat.setGeometry(zf.geometry())
            except Exception:
                pass
            feat[zone_label_field] = str(label) if label is not None else None
            feat['prelev_sum'] = float(prelev) if prelev is not None else None
            feat['autor_sum'] = float(autor) if autor is not None else None
            feat['ratio'] = float(r) if r is not None else None
            feat['ratio_possible'] = int(rpos)
            feat['percent_prelev_auth'] = float(ppre) if ppre is not None else None
            feat['percent_overrun'] = float(pover) if pover is not None else None
            feat['n_ouvrages'] = int(n_ouv)
            try:
                sink.addFeature(feat, QgsFeatureSink.FastInsert)
            except TypeError:
                sink.addFeature(feat)
            cnt += 1
            feedback.setProgress(int(100 * cnt / max(1, total_z)))

        # écrire la feature "Non assigné" (sans géométrie) si elle contient quelque chose
        un_prelev = zone_prelev_sum.get(UNASSIGNED_LABEL, 0.0)
        un_n = zone_count_ouvrages.get(UNASSIGNED_LABEL, 0)
        un_autor = zone_autor_sum.get(UNASSIGNED_LABEL)
        if un_n > 0 or (un_prelev != 0.0) or (un_autor is not None):
            feat_un = QgsFeature()
            feat_un.setFields(out_fields)
            # pas de géométrie (on laisse la géométrie None)
            feat_un[zone_label_field] = UNASSIGNED_LABEL
            feat_un['prelev_sum'] = float(un_prelev) if un_prelev is not None else None
            feat_un['autor_sum'] = float(un_autor) if un_autor is not None else None
            r_un = zone_ratio.get(UNASSIGNED_LABEL)
            feat_un['ratio'] = float(r_un) if r_un is not None else None
            feat_un['ratio_possible'] = int(zone_ratio_possible.get(UNASSIGNED_LABEL, 0))
            feat_un['percent_prelev_auth'] = float(zone_percent_prelev_auth.get(UNASSIGNED_LABEL)) if zone_percent_prelev_auth.get(UNASSIGNED_LABEL) is not None else None
            feat_un['percent_overrun'] = float(zone_percent_overrun.get(UNASSIGNED_LABEL)) if zone_percent_overrun.get(UNASSIGNED_LABEL) is not None else None
            feat_un['n_ouvrages'] = int(un_n)
            try:
                # certains drivers acceptent la géométrie nulle ; on tente de l'ajouter sans géométrie
                sink.addFeature(feat_un, QgsFeatureSink.FastInsert)
            except TypeError:
                try:
                    sink.addFeature(feat_un)
                except Exception:
                    # si l'écriture sans géométrie échoue (rare), on lève une info mais pas d'erreur critique
                    feedback.pushInfo(self.tr("Impossible d'écrire la feature 'Non assigné' sans géométrie avec ce fournisseur de sortie."))

        feedback.pushInfo(self.tr(f"Ecriture terminée : {cnt} entités (zones) écrites + éventuelle entrée '{UNASSIGNED_LABEL}'."))

        # ---------- 7) Appliquer QML si demandé ----------
        try:
            if apply_qml:
                result_layer = QgsProcessingUtils.mapLayerFromString(dest_id, context)
                if result_layer is not None:
                    qml_path = os.path.normpath(qml_path_param) if qml_path_param else ''
                    if qml_path and os.path.exists(qml_path):
                        try:
                            res = result_layer.loadNamedStyle(qml_path)
                            if isinstance(res, tuple):
                                ok, message = res
                            else:
                                ok = bool(res)
                                message = ''
                        except TypeError:
                            ok = result_layer.loadNamedStyle(qml_path)
                            message = ''
                        except Exception as e:
                            ok = False
                            message = str(e)
                        result_layer.triggerRepaint()
                        if QgsProject.instance().mapLayer(result_layer.id()) is None:
                            QgsProject.instance().addMapLayer(result_layer)
                        if not ok:
                            feedback.pushInfo(self.tr(f"Style QML chargé, mais QGIS a renvoyé un message : {message}"))
                        else:
                            feedback.pushInfo(self.tr(f"Style QML appliqué depuis : {qml_path}"))
                    else:
                        feedback.pushInfo(self.tr(f"QML introuvable au chemin : {qml_path}"))
                else:
                    feedback.pushInfo(self.tr("Impossible de récupérer la couche de sortie pour appliquer le QML."))
        except Exception as e:
            feedback.pushInfo(self.tr(f"Erreur lors de l'application du QML : {e}"))

        return {self.OUTPUT: dest_id}

# Fin du script
