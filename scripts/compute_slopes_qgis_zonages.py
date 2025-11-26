# Script Processing QGIS : Coef directeur par zonage (points -> zones)
from qgis.PyQt.QtCore import QVariant
from qgis.core import (
    QgsProcessing,
    QgsProcessingAlgorithm,
    QgsProcessingParameterVectorLayer,
    QgsProcessingParameterField,
    QgsProcessingParameterEnum,
    QgsProcessingParameterNumber,
    QgsProcessingParameterFeatureSink,
    QgsProcessingParameterBoolean,
    QgsProcessingParameterString,
    QgsFeature,
    QgsField,
    QgsProject,
    QgsFields,
    QgsFeatureSink,
    QgsProcessingUtils,
    QgsSpatialIndex,
)
import math
from collections import defaultdict
import re
import os

# Optional libs
use_numpy = False
use_scipy = False
try:
    import numpy as np
    use_numpy = True
except Exception:
    pass
try:
    from scipy.stats import theilslopes
    use_scipy = True
except Exception:
    pass


def parse_number(x):
    """
    Parse un nombre donné au format français/anglais :
    - Accepte "12 000,56" ou "12000.56" etc.
    - Supprime unités et caractères non numériques
    - Retourne float ou NaN
    """
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


def median_of_pairwise_slopes(xs, ys):
    """Fallback Theil-Sen: médiane des pentes pairwise (O(n^2))."""
    n = len(xs)
    slopes = []
    for i in range(n - 1):
        for j in range(i + 1, n):
            dx = xs[j] - xs[i]
            if dx != 0:
                slopes.append((ys[j] - ys[i]) / dx)
    if not slopes:
        return None
    slopes.sort()
    m = len(slopes)
    if m % 2 == 1:
        return float(slopes[m // 2])
    else:
        return float((slopes[m // 2 - 1] + slopes[m // 2]) / 2.0)


def compute_slope_years(years, values, method='OLS'):
    """
    Retourne la pente (units = vol / an).
    method: 'OLS' ou 'Theil-Sen'
    """
    pairs = [(y, v) for y, v in zip(years, values) if v is not None and not (isinstance(v, float) and math.isnan(v))]
    if len(pairs) < 2:
        return None
    ys, vs = zip(*pairs)
    if method == 'Theil-Sen':
        try:
            if use_scipy and use_numpy:
                # theilslopes returns (slope, intercept, lower, upper)
                res = theilslopes(np.array(vs, dtype=float), np.array(ys, dtype=float))
                return float(res[0])
            else:
                return median_of_pairwise_slopes(list(ys), list(vs))
        except Exception:
            return median_of_pairwise_slopes(list(ys), list(vs))
    else:
        try:
            if use_numpy:
                m, b = np.polyfit(np.array(ys, dtype=float), np.array(vs, dtype=float), 1)
                return float(m)
            else:
                n = len(ys)
                x_mean = sum(ys) / n
                y_mean = sum(vs) / n
                num = sum((xi - x_mean) * (yi - y_mean) for xi, yi in zip(ys, vs))
                den = sum((xi - x_mean) ** 2 for xi in ys)
                if den == 0:
                    return None
                return float(num / den)
        except Exception:
            return None


class ZonesSlopesAlgorithm(QgsProcessingAlgorithm):
    """
    Algorithme Processing : agrège points (ouvrages) par zones et calcule pentes par zone.
    Multi-affectation : un ouvrage affecté à toutes les zones intersectées.
    La géométrie d'affectation d'un ouvrage est prise depuis l'enregistrement de l'année
    la plus récente disponible (dans la période sélectionnée).
    """

    # paramètres
    ZONES = 'ZONES'
    ZONE_ID = 'ZONE_ID'
    OUVRAGES = 'OUVRAGES'
    YEAR = 'YEAR'
    OUV_ID = 'OUV_ID'
    VOL = 'VOL'
    METHOD = 'METHOD'
    MIN_YEARS = 'MIN_YEARS'
    START_YEAR = 'START_YEAR'
    END_YEAR = 'END_YEAR'
    APPLY_QML = 'APPLY_QML'
    QML_PATH = 'QML_PATH'
    OUTPUT = 'OUTPUT'
    OUTPUT_ZONE_YEAR = 'OUTPUT_ZONE_YEAR'

    def tr(self, string):
        return string

    def createInstance(self):
        return ZonesSlopesAlgorithm()

    def name(self):
        return 'compute_slopes_zones'

    def displayName(self):
        return self.tr('Pentes par zonage (agg. points → zones)')

    def group(self):
        return self.tr('Analyses temporelles')

    def groupId(self):
        return 'temporal_analysis'

    def shortHelpString(self):
        return self.tr(
            "Agrège les volumes des ouvrages (points) par zone (multi-affectation si intersecte plusieurs zones), "
            "puis calcule la pente (OLS/Theil-Sen) par zone sur la période choisie. "
            "La géométrie utilisée pour assigner chaque ouvrage est celle de l'enregistrement "
            "contenant l'année la plus récente disponible (dans la période)."
        )

    def initAlgorithm(self, config=None):
        self.addParameter(
            QgsProcessingParameterVectorLayer(self.ZONES, self.tr("Couche de zonage (polygones)"), [QgsProcessing.TypeVectorAnyGeometry])
        )
        self.addParameter(
            QgsProcessingParameterField(self.ZONE_ID, self.tr("Champ identifiant de la zone"), parentLayerParameterName=self.ZONES)
        )
        self.addParameter(
            QgsProcessingParameterVectorLayer(self.OUVRAGES, self.tr("Couche ouvrages (points / table)"), [QgsProcessing.TypeVectorAnyGeometry])
        )
        self.addParameter(
            QgsProcessingParameterField(self.YEAR, self.tr("Champ année (ouvrages)"), parentLayerParameterName=self.OUVRAGES, type=QgsProcessingParameterField.Numeric)
        )
        self.addParameter(
            QgsProcessingParameterField(self.OUV_ID, self.tr("Champ identifiant ouvrage (N°Ouvrage)"), parentLayerParameterName=self.OUVRAGES)
        )
        self.addParameter(
            QgsProcessingParameterField(self.VOL, self.tr("Champ volume (Assiette)"), parentLayerParameterName=self.OUVRAGES)
        )
        self.addParameter(
            QgsProcessingParameterEnum(self.METHOD, self.tr("Méthode pour estimer la pente"), options=['OLS', 'Theil-Sen'])
        )
        self.addParameter(
            QgsProcessingParameterNumber(self.MIN_YEARS, self.tr("Années minimales pour calculer une pente"), type=QgsProcessingParameterNumber.Integer, defaultValue=4)
        )
        self.addParameter(
            QgsProcessingParameterNumber(self.START_YEAR, self.tr("Année de début"), type=QgsProcessingParameterNumber.Integer, defaultValue=2012)
        )
        self.addParameter(
            QgsProcessingParameterNumber(self.END_YEAR, self.tr("Année de fin"), type=QgsProcessingParameterNumber.Integer, defaultValue=2023)
        )
        self.addParameter(
            QgsProcessingParameterBoolean(self.APPLY_QML, self.tr("Appliquer un style QML sur la couche de sortie ?"), defaultValue=True)
        )
        # default QML same as earlier; user can change
        default_qml = r"N:\_MTP\Public\01-ORGANISATION\G-Services\RAGAF\REDEVANCES\Recherche redevables\Etude données prélèvements\06_Valorisation_Visualisation\Outils\QML\Zonages_Slopes_QML.qml"
        self.addParameter(
            QgsProcessingParameterString(self.QML_PATH, self.tr("Chemin du fichier QML (si appliqué)"), defaultValue=default_qml)
        )
        self.addParameter(
            QgsProcessingParameterFeatureSink(self.OUTPUT, self.tr("Couche de sortie (pentes par zone)"))
        )
        # optional: table zone x year for diagnostics
        self.addParameter(
            QgsProcessingParameterFeatureSink(self.OUTPUT_ZONE_YEAR, self.tr("Table (zone x année) - optionnel (laisser vide si pas besoin)"))
        )

    def processAlgorithm(self, parameters, context, feedback):
        zones_lyr = self.parameterAsVectorLayer(parameters, self.ZONES, context)
        zone_id_field = self.parameterAsString(parameters, self.ZONE_ID, context)
        ouvrages_lyr = self.parameterAsVectorLayer(parameters, self.OUVRAGES, context)
        year_field = self.parameterAsString(parameters, self.YEAR, context)
        ouv_id_field = self.parameterAsString(parameters, self.OUV_ID, context)
        vol_field = self.parameterAsString(parameters, self.VOL, context)
        method_idx = self.parameterAsInt(parameters, self.METHOD, context)
        method = ['OLS', 'Theil-Sen'][method_idx]
        min_years = int(self.parameterAsInt(parameters, self.MIN_YEARS, context))
        start_year = int(self.parameterAsInt(parameters, self.START_YEAR, context))
        end_year = int(self.parameterAsInt(parameters, self.END_YEAR, context))
        apply_qml = bool(self.parameterAsBool(parameters, self.APPLY_QML, context))
        qml_path_param = self.parameterAsString(parameters, self.QML_PATH, context)

        # 1) Lire les ouvrages : garder tous les enregistrements entre start_year et end_year
        #    On enregistre pour chaque ouvrage :
        #      - tous les tuples (year, volume)
        #      - la géométrie associée à l'année la plus récente disponible (pour l'affectation spatiale)
        rows = []  # tuples (ouv_id, year, vol)
        geom_by_ouv_latest = {}  # ouv_id -> (year, geometry)
        total = ouvrages_lyr.featureCount()
        processed = 0
        for f in ouvrages_lyr.getFeatures():
            processed += 1
            if feedback.isCanceled():
                break
            try:
                y_raw = f[year_field]
                o = f[ouv_id_field]
                v_raw = f[vol_field]
            except Exception:
                raise Exception(self.tr("Impossible de lire au moins un des champs fournis dans la couche ouvrages. Vérifie les paramètres."))
            # parse year
            try:
                yv = int(y_raw)
            except:
                # ignore non-numeric years
                continue
            if yv < start_year or yv > end_year:
                continue
            vv = parse_number(v_raw)
            rows.append((o, yv, vv))
            # geometry handling : keep geometry of most recent year per ouvrage
            if ouvrages_lyr.geometryType() != -1:
                geom = f.geometry()
                prev = geom_by_ouv_latest.get(o)
                if prev is None or (isinstance(prev[0], int) and yv > prev[0]):
                    geom_by_ouv_latest[o] = (yv, geom)
            feedback.setProgress(int(100 * processed / total) if total else 0)

        if not rows:
            raise Exception(self.tr("Aucune donnée ouvrages valide pour la période sélectionnée."))

        # 2) Construire dictionnaire ouvrage -> liste des (year, vol)
        ouv_map = defaultdict(list)
        for o, y, v in rows:
            val = 0.0 if (v is None or (isinstance(v, float) and math.isnan(v))) else v
            ouv_map[o].append((y, val))

        # 3) Construire mapping ouvrage_id -> zones (multi-affectation)
        #    On utilise la géométrie 'latest' pour l'ouvrage (si disponible)
        feedback.pushInfo("Construction index spatial des zones...")
        zone_index = QgsSpatialIndex(zones_lyr.getFeatures())
        ouv_to_zones = defaultdict(list)  # ouv_id -> list of zone_ids
        missing_geom_count = 0
        for idx, (ouv, pairs) in enumerate(ouv_map.items()):
            if feedback.isCanceled():
                break
            # get latest geometry
            latest = geom_by_ouv_latest.get(ouv)
            if latest is None:
                missing_geom_count += 1
                continue
            geom = latest[1]
            if geom is None or geom.isEmpty():
                missing_geom_count += 1
                continue
            # find candidate zone feature ids by bbox
            candidates = zone_index.intersects(geom.boundingBox())
            # for each candidate, check real intersection
            found = False
            for fid in candidates:
                z_feat = zones_lyr.getFeature(fid)
                if z_feat is None:
                    continue
                z_geom = z_feat.geometry()
                try:
                    if z_geom.intersects(geom):
                        zone_id = z_feat[zone_id_field]
                        ouv_to_zones[ouv].append(zone_id)
                        found = True
                except Exception:
                    # fallback: bounding box intersection already true, but geometry op failed; skip
                    continue
            # if no zone found, keep empty list (ouvrage non assigné)
            if idx % 200 == 0:
                feedback.setProgress(int(100.0 * idx / max(1, len(ouv_map))))
        if missing_geom_count > 0:
            feedback.pushInfo(f"{missing_geom_count} ouvrages sans géométrie 'latest' et non assignés à des zones.")

        # 4) Agréger volumes par zone x year (multi-affectation -> ouvrage affecté à toutes les zones correspondantes)
        zone_year_sum = defaultdict(float)   # (zone_id, year) -> sum volumes
        zone_year_count_valid = defaultdict(int)
        for o, pairs in ouv_map.items():
            zones_for_o = ouv_to_zones.get(o, [])
            if not zones_for_o:
                # on n'affecte pas l'ouvrage s'il n'est dans aucune zone
                continue
            for (y, v) in pairs:
                for zid in zones_for_o:
                    zone_year_sum[(zid, y)] += v
                    if not (v is None or (isinstance(v, float) and math.isnan(v))):
                        zone_year_count_valid[(zid, y)] += 1

        # 5) Construire structure zone -> list of (year, total)
        zone_years_map = defaultdict(list)
        for (z, y), tot in zone_year_sum.items():
            zone_years_map[z].append((y, tot))

        if not zone_years_map:
            raise Exception(self.tr("Aucun agrégat zone×année n'a été produit (vérifie intersections / géométries)."))

        # 6) Calculer pentes par zone (et metrics)
        zone_to_slope = {}
        zone_to_nyears = {}
        zone_mean_vol = {}
        zone_first3_mean = {}
        zone_last3_mean = {}
        zone_slope_pct_mean = {}
        zone_slope_pct_first = {}
        zone_cagr_pct = {}

        # first compute slopes and n_years and means
        for z, lst in zone_years_map.items():
            lst_sorted = sorted(lst, key=lambda x: x[0])
            yrs = [x[0] for x in lst_sorted]
            tots = [x[1] for x in lst_sorted]
            nyrs = len([val for val in tots if not (isinstance(val, float) and math.isnan(val))])
            zone_to_nyears[z] = nyrs
            if nyrs >= min_years:
                s = compute_slope_years(yrs, tots, method=method)
            else:
                s = None
            zone_to_slope[z] = s
            vals = [v for (_, v) in lst_sorted if not (isinstance(v, float) and math.isnan(v))]
            zone_mean_vol[z] = (sum(vals) / len(vals)) if vals else float('nan')
            non_nan_pairs = [(y, v) for (y, v) in lst_sorted if not (isinstance(v, float) and math.isnan(v))]
            if non_nan_pairs:
                first3 = [v for (_, v) in non_nan_pairs[:3]]
                last3 = [v for (_, v) in non_nan_pairs[-3:]]
                first3_mean = sum(first3) / len(first3) if first3 else float('nan')
                last3_mean = sum(last3) / len(last3) if last3 else float('nan')
            else:
                first3_mean = float('nan')
                last3_mean = float('nan')
            zone_first3_mean[z] = first3_mean
            zone_last3_mean[z] = last3_mean

        for z in zone_to_slope.keys():
            slope = zone_to_slope.get(z)
            meanv = zone_mean_vol.get(z)
            first3 = zone_first3_mean.get(z)
            last3 = zone_last3_mean.get(z)
            if slope is None or meanv is None or (isinstance(meanv, float) and math.isnan(meanv)) or meanv == 0:
                zone_slope_pct_mean[z] = None
            else:
                zone_slope_pct_mean[z] = 100.0 * (slope / meanv)
            if slope is None or first3 is None or (isinstance(first3, float) and math.isnan(first3)) or first3 == 0:
                zone_slope_pct_first[z] = None
            else:
                zone_slope_pct_first[z] = 100.0 * (slope / first3)
            # CAGR using mean first3 / mean last3
            lst_sorted = sorted(zone_years_map.get(z, []), key=lambda x: x[0])
            if len(lst_sorted) >= 2:
                year_first = lst_sorted[0][0]
                year_last = lst_sorted[-1][0]
                n_periods = year_last - year_first
                if n_periods > 0 and first3 is not None and last3 is not None and first3 > 0:
                    try:
                        cagr = (last3 / first3) ** (1.0 / n_periods) - 1.0
                        zone_cagr_pct[z] = 100.0 * cagr
                    except Exception:
                        zone_cagr_pct[z] = None
                else:
                    zone_cagr_pct[z] = None
            else:
                zone_cagr_pct[z] = None

        # 7) z-score on slope_pct_mean across zones
        all_pct = [v for v in list(zone_slope_pct_mean.values()) if v is not None]
        if len(all_pct) >= 2:
            mean_pct = sum(all_pct) / len(all_pct)
            sd_pct = (sum((x - mean_pct) ** 2 for x in all_pct) / (len(all_pct) - 1)) ** 0.5
        else:
            mean_pct = None
            sd_pct = None
        zone_slope_pct_z = {}
        for z, pct in zone_slope_pct_mean.items():
            if pct is None or mean_pct is None or sd_pct is None or sd_pct == 0:
                zone_slope_pct_z[z] = None
            else:
                zone_slope_pct_z[z] = (pct - mean_pct) / sd_pct

        # 8) Préparer sink de sortie (une ligne par zone)
        out_fields = QgsFields()
        out_fields.append(QgsField(zone_id_field, QVariant.String))
        out_fields.append(QgsField('slope_zone', QVariant.Double))
        out_fields.append(QgsField('n_years_zone', QVariant.Int))
        out_fields.append(QgsField('mean_vol_zone', QVariant.Double))
        out_fields.append(QgsField('slope_pct_mean', QVariant.Double))
        out_fields.append(QgsField('slope_pct_first', QVariant.Double))
        out_fields.append(QgsField('cagr_pct', QVariant.Double))
        out_fields.append(QgsField('slope_pct_z', QVariant.Double))

        (sink, dest_id) = self.parameterAsSink(parameters, self.OUTPUT, context,
                                               out_fields,
                                               zones_lyr.wkbType(), zones_lyr.sourceCrs())

        # remplir le sink : parcourir les features des zones et écrire les valeurs correspondantes (pour garder la géométrie originale)
        total_z = zones_lyr.featureCount()
        p = 0
        for zf in zones_lyr.getFeatures():
            p += 1
            if feedback.isCanceled():
                break
            zid = zf[zone_id_field]
            feat = QgsFeature()
            feat.setFields(out_fields)
            feat.setGeometry(zf.geometry())
            feat[zone_id_field] = str(zid)
            feat['slope_zone'] = float(zone_to_slope.get(zid)) if zone_to_slope.get(zid) is not None else None
            feat['n_years_zone'] = int(zone_to_nyears.get(zid, 0))
            feat['mean_vol_zone'] = float(zone_mean_vol.get(zid)) if zone_mean_vol.get(zid) is not None else None
            feat['slope_pct_mean'] = float(zone_slope_pct_mean.get(zid)) if zone_slope_pct_mean.get(zid) is not None else None
            feat['slope_pct_first'] = float(zone_slope_pct_first.get(zid)) if zone_slope_pct_first.get(zid) is not None else None
            feat['cagr_pct'] = float(zone_cagr_pct.get(zid)) if zone_cagr_pct.get(zid) is not None else None
            feat['slope_pct_z'] = float(zone_slope_pct_z.get(zid)) if zone_slope_pct_z.get(zid) is not None else None
            # add feature
            try:
                sink.addFeature(feat, QgsFeatureSink.FastInsert)
            except TypeError:
                sink.addFeature(feat)
            feedback.setProgress(int(100 * p / total_z) if total_z else 100)

        # 9) Optionnel : produire une table zone x year (utile pour diagnostics)
        (sink2, dest_id2) = self.parameterAsSink(parameters, self.OUTPUT_ZONE_YEAR, context,
                                                 QgsFields(),  # fields will be created below if sink2 exists
                                                 zones_lyr.wkbType(), zones_lyr.sourceCrs()) if self.OUTPUT_ZONE_YEAR in self.parameterDefinitions() else (None, None)
        # note: parameterAsSink always returns a sink even if user left the parameter empty for optional sinks in some QGIS versions. We handle robustly.
        try:
            # create fields for zone-year table
            zy_fields = QgsFields()
            zy_fields.append(QgsField(zone_id_field, QVariant.String))
            zy_fields.append(QgsField('year', QVariant.Int))
            zy_fields.append(QgsField('sum_vol', QVariant.Double))
            if dest_id2:
                # recreate sink with correct fields (parameterAsSink already returned something; to be safe, we will attempt to write directly if possible)
                pass
        except Exception:
            pass

        # If OUTPUT_ZONE_YEAR was provided, try to write rows via processing mapLayerFromString to get the sink layer and write manually.
        try:
            # detect whether OUTPUT_ZONE_YEAR was configured by user: parameterAsSink returns dest id - check it
            ctx_sink = None
            try:
                # try retrieve dest_id2 as a layer
                if dest_id2:
                    ctx_sink = QgsProcessingUtils.mapLayerFromString(dest_id2, context)
            except Exception:
                ctx_sink = None
            if ctx_sink is not None:
                # write zone-year rows
                zy_fields = QgsFields()
                zy_fields.append(QgsField(zone_id_field, QVariant.String))
                zy_fields.append(QgsField('year', QVariant.Int))
                zy_fields.append(QgsField('sum_vol', QVariant.Double))
                # add features from zone_year_sum
                for (z, y), tot in sorted(zone_year_sum.items()):
                    fzy = QgsFeature()
                    fzy.setFields(zy_fields)
                    fzy[zone_id_field] = str(z)
                    fzy['year'] = int(y)
                    fzy['sum_vol'] = float(tot) if tot is not None else None
                    try:
                        ctx_sink.addFeature(fzy, QgsFeatureSink.FastInsert)
                    except TypeError:
                        ctx_sink.addFeature(fzy)
        except Exception:
            # ignore optional table write errors (not critical)
            pass

        # 10) Appliquer le QML si demandé (sur la couche de sortie zones)
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
                            feedback.pushInfo("Style QML chargé, mais QGIS a renvoyé un message : {}".format(message))
                        else:
                            feedback.pushInfo("Style QML appliqué depuis : {}".format(qml_path))
                    else:
                        feedback.pushInfo("QML introuvable au chemin : {}".format(qml_path))
                else:
                    feedback.pushInfo("Impossible de récupérer la couche de sortie pour appliquer le QML.")
        except Exception as e:
            feedback.pushInfo("Erreur lors de l'application du style QML : {}".format(e))

        return {self.OUTPUT: dest_id, self.OUTPUT_ZONE_YEAR: dest_id2 if 'dest_id2' in locals() else None}

# Fin du script - sauvegarder dans Processing > Scripts > Tools
