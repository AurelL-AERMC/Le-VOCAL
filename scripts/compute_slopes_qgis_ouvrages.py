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
)
import math
from collections import defaultdict
import re
import os

# Optional libs
use_pandas = False
use_numpy = False
use_scipy = False
try:
    import pandas as pd
    use_pandas = True
except Exception:
    pass
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
    Parse un nombre donné au format français ou anglais :
    - Accepte 12000,56 ou 12 000,56 ou 12.000,56 ou 12000.56
    - Supprime unités (ex: ' m3') et caractères non numériques
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
    """Retourne la pente (units = vol / an). method: 'OLS' ou 'Theil-Sen'"""
    pairs = [(y, v) for y, v in zip(years, values) if v is not None and not (isinstance(v, float) and math.isnan(v))]
    if len(pairs) < 2:
        return None
    ys, vs = zip(*pairs)
    if method == 'Theil-Sen':
        try:
            if use_scipy and use_numpy:
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


class ComputeSlopesByOuvrage(QgsProcessingAlgorithm):
    """Algorithme Processing : calcule pentes et indicateurs POUR CHAQUE OUVRAGE (sans BV)."""

    INPUT = 'INPUT'
    YEAR = 'YEAR'
    OUVRAGE = 'OUVRAGE'
    VOL = 'VOL'
    METHOD = 'METHOD'
    MIN_YEARS = 'MIN_YEARS'
    START_YEAR = 'START_YEAR'
    END_YEAR = 'END_YEAR'
    APPLY_QML = 'APPLY_QML'
    QML_PATH = 'QML_PATH'
    OUTPUT = 'OUTPUT'

    def tr(self, string):
        return string

    def createInstance(self):
        return ComputeSlopesByOuvrage()

    def name(self):
        return 'compute_slopes_ouvrage_only'

    def displayName(self):
        return self.tr('Pentes par ouvrage (nettoyé, sans BV)')

    def group(self):
        return self.tr('Analyses temporelles')

    def groupId(self):
        return 'temporal_analysis'

    def shortHelpString(self):
        return self.tr(
            "Calcule la pente (coef directeur) pour chaque ouvrage (somme par ouvrage×année). "
            "Méthodes: OLS ou Theil-Sen. Produit aussi pentes en %/an et CAGR (moyenne 3 premières / 3 dernières années)."
        )

    def initAlgorithm(self, config=None):
        self.addParameter(
            QgsProcessingParameterVectorLayer(self.INPUT, self.tr("Couche d'entrée (points/table)"), [QgsProcessing.TypeVectorAnyGeometry])
        )
        self.addParameter(
            QgsProcessingParameterField(self.YEAR, self.tr("Champ année"), parentLayerParameterName=self.INPUT, type=QgsProcessingParameterField.Numeric)
        )
        self.addParameter(
            QgsProcessingParameterField(self.OUVRAGE, self.tr("Champ identifiant ouvrage"), parentLayerParameterName=self.INPUT)
        )
        self.addParameter(
            QgsProcessingParameterField(self.VOL, self.tr("Champ volume (Assiette)"), parentLayerParameterName=self.INPUT)
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
        # demander chemin QML si l'utilisateur veut l'appliquer
        default_qml = r"N:\_MTP\Public\01-ORGANISATION\G-Services\RAGAF\REDEVANCES\Recherche redevables\Etude données prélèvements\06_Valorisation_Visualisation\Outils\QML\Ouvrages_Slopes_QML.qml"
        self.addParameter(
            QgsProcessingParameterString(self.QML_PATH, self.tr("Chemin du fichier QML (si appliqué)"), defaultValue=default_qml)
        )
        self.addParameter(
            QgsProcessingParameterFeatureSink(self.OUTPUT, self.tr("Couche de sortie (pentes par ouvrage)"))
        )

    def processAlgorithm(self, parameters, context, feedback):
        layer = self.parameterAsVectorLayer(parameters, self.INPUT, context)
        year_field = self.parameterAsString(parameters, self.YEAR, context)
        ouvrage_field = self.parameterAsString(parameters, self.OUVRAGE, context)
        vol_field = self.parameterAsString(parameters, self.VOL, context)
        method_idx = self.parameterAsInt(parameters, self.METHOD, context)
        method = ['OLS', 'Theil-Sen'][method_idx]
        min_years = int(self.parameterAsInt(parameters, self.MIN_YEARS, context))
        start_year = int(self.parameterAsInt(parameters, self.START_YEAR, context))
        end_year = int(self.parameterAsInt(parameters, self.END_YEAR, context))
        apply_qml = bool(self.parameterAsBool(parameters, self.APPLY_QML, context))
        qml_path_param = self.parameterAsString(parameters, self.QML_PATH, context)

        # lecture et filtrage initial
        rows = []
        has_geometry = (layer.geometryType() != -1)
        geom_by_ouvrage = {}
        total = layer.featureCount()
        processed = 0
        for f in layer.getFeatures():
            processed += 1
            if feedback.isCanceled():
                break
            try:
                y = f[year_field]
                o = f[ouvrage_field]
                v_raw = f[vol_field]
            except Exception:
                raise Exception(self.tr("Impossible de lire au moins un des champs fournis. Vérifie les paramètres."))
            try:
                yv = int(y)
            except:
                continue
            if yv < start_year or yv > end_year:
                continue
            vv = parse_number(v_raw)
            rows.append((o, yv, vv))
            if has_geometry and o not in geom_by_ouvrage:
                geom_by_ouvrage[o] = f.geometry()
            feedback.setProgress(int(100 * processed / total))

        if not rows:
            raise Exception(self.tr("Aucune donnée lue pour la période sélectionnée."))

        # --- AGREGATION DES VOLUMES PAR (ouvrage, year) ---
        ouvrage_year_sum = defaultdict(float)   # key (ouvrage, year) -> sum
        ouvrage_year_count_valid = defaultdict(int)

        for o, y, v in rows:
            if v is None or (isinstance(v, float) and math.isnan(v)):
                val = 0.0
            else:
                val = v
            ouvrage_year_sum[(o, y)] += val
            if not (v is None or (isinstance(v, float) and math.isnan(v))):
                ouvrage_year_count_valid[(o, y)] += 1

        # construire structure par ouvrage à partir des sommes
        ouvrage_map = defaultdict(list)   # ouvrage -> list of (year, total)
        for (o, y), tot in ouvrage_year_sum.items():
            ouvrage_map[o].append((y, tot))

        # calcul des pentes pour les ouvrages (sur les séries agrégées ouvrage x année)
        ouvrage_to_slope = {}
        ouvrage_to_nyears = {}
        for o, lst in ouvrage_map.items():
            lst_sorted = sorted(lst, key=lambda x: x[0])
            yrs = [x[0] for x in lst_sorted]
            vols = [x[1] for x in lst_sorted]
            nyrs = len([val for val in vols if not (isinstance(val, float) and math.isnan(val))])
            ouvrage_to_nyears[o] = nyrs
            if nyrs >= min_years:
                s = compute_slope_years(yrs, vols, method=method)
            else:
                s = None
            ouvrage_to_slope[o] = s

        # ---------- Normalisation en % / an et CAGR (moyennes 3 premières / 3 dernières années) ----------
        ouvrage_mean_vol = {}
        ouvrage_first3_mean = {}
        ouvrage_last3_mean = {}
        ouvrage_slope_pct_mean = {}
        ouvrage_slope_pct_first = {}
        ouvrage_cagr_pct = {}
        ouvrage_first_last_years = {}

        for o, lst in ouvrage_map.items():
            lst_sorted = sorted(lst, key=lambda x: x[0])
            vals = [v for (_, v) in lst_sorted if not (isinstance(v, float) and math.isnan(v))]
            if vals:
                ouvrage_mean_vol[o] = sum(vals) / len(vals)
            else:
                ouvrage_mean_vol[o] = float('nan')
            non_nan_pairs = [(y, v) for (y, v) in lst_sorted if not (isinstance(v, float) and math.isnan(v))]
            if non_nan_pairs:
                first3 = [v for (_, v) in non_nan_pairs[:3]]
                last3 = [v for (_, v) in non_nan_pairs[-3:]]
                first3_mean = sum(first3) / len(first3) if first3 else float('nan')
                last3_mean = sum(last3) / len(last3) if last3 else float('nan')
                year_first = non_nan_pairs[0][0]
                year_last = non_nan_pairs[-1][0]
            else:
                first3_mean = float('nan')
                last3_mean = float('nan')
                year_first = None
                year_last = None
            ouvrage_first3_mean[o] = first3_mean
            ouvrage_last3_mean[o] = last3_mean
            ouvrage_first_last_years[o] = (year_first, year_last)

        for o in ouvrage_to_slope.keys():
            slope = ouvrage_to_slope.get(o)
            meanv = ouvrage_mean_vol.get(o)
            first3 = ouvrage_first3_mean.get(o)
            last3 = ouvrage_last3_mean.get(o)
            # pct relatif par rapport à la moyenne
            if slope is None or meanv is None or (isinstance(meanv, float) and math.isnan(meanv)) or meanv == 0:
                ouvrage_slope_pct_mean[o] = None
            else:
                ouvrage_slope_pct_mean[o] = 100.0 * (slope / meanv)
            # pct relatif par rapport à first3_mean
            if slope is None or first3 is None or (isinstance(first3, float) and math.isnan(first3)) or first3 == 0:
                ouvrage_slope_pct_first[o] = None
            else:
                ouvrage_slope_pct_first[o] = 100.0 * (slope / first3)
            # CAGR using mean first3 / mean last3
            year_first, year_last = ouvrage_first_last_years.get(o, (None, None))
            if year_first is not None and year_last is not None and year_last > year_first and first3 is not None and last3 is not None and first3 > 0:
                n_periods = year_last - year_first
                try:
                    cagr = (last3 / first3) ** (1.0 / n_periods) - 1.0
                    ouvrage_cagr_pct[o] = 100.0 * cagr
                except Exception:
                    ouvrage_cagr_pct[o] = None
            else:
                ouvrage_cagr_pct[o] = None

        # z-score on slope_pct_mean across ouvrages
        all_pct = [v for v in list(ouvrage_slope_pct_mean.values()) if v is not None]
        if len(all_pct) >= 2:
            mean_pct = sum(all_pct) / len(all_pct)
            sd_pct = (sum((x - mean_pct) ** 2 for x in all_pct) / (len(all_pct) - 1)) ** 0.5
        else:
            mean_pct = None
            sd_pct = None

        ouvrage_slope_pct_z = {}
        for o, pct in ouvrage_slope_pct_mean.items():
            if pct is None or mean_pct is None or sd_pct is None or sd_pct == 0:
                ouvrage_slope_pct_z[o] = None
            else:
                ouvrage_slope_pct_z[o] = (pct - mean_pct) / sd_pct

        # --- PREPARER LE SINK DE SORTIE (QgsFields) ---
        out_fields = QgsFields()
        out_fields.append(QgsField('ouvrage_id', QVariant.String))
        out_fields.append(QgsField('slope_ouvrage', QVariant.Double))
        out_fields.append(QgsField('n_years_ouvrage', QVariant.Int))
        # normalization fields
        out_fields.append(QgsField('mean_vol_ouv', QVariant.Double))
        out_fields.append(QgsField('slope_pct_mean', QVariant.Double))
        out_fields.append(QgsField('slope_pct_first', QVariant.Double))
        out_fields.append(QgsField('cagr_pct', QVariant.Double))
        out_fields.append(QgsField('slope_pct_z', QVariant.Double))

        (sink, dest_id) = self.parameterAsSink(parameters, self.OUTPUT, context,
                                               out_fields,
                                               layer.wkbType(), layer.sourceCrs())

        # remplir le sink (une ligne par ouvrage)
        total2 = len(ouvrage_to_slope)
        cnt = 0
        for o in sorted(ouvrage_to_slope.keys()):
            if feedback.isCanceled():
                break
            feat = QgsFeature()
            feat.setFields(out_fields)
            feat['ouvrage_id'] = str(o)
            val_ouv = ouvrage_to_slope.get(o)
            feat['slope_ouvrage'] = float(val_ouv) if val_ouv is not None else None
            feat['n_years_ouvrage'] = int(ouvrage_to_nyears.get(o, 0))
            # mean volumes
            feat['mean_vol_ouv'] = float(ouvrage_mean_vol.get(o)) if ouvrage_mean_vol.get(o) is not None else None
            # normalized metrics
            feat['slope_pct_mean'] = float(ouvrage_slope_pct_mean.get(o)) if ouvrage_slope_pct_mean.get(o) is not None else None
            feat['slope_pct_first'] = float(ouvrage_slope_pct_first.get(o)) if ouvrage_slope_pct_first.get(o) is not None else None
            feat['cagr_pct'] = float(ouvrage_cagr_pct.get(o)) if ouvrage_cagr_pct.get(o) is not None else None
            feat['slope_pct_z'] = float(ouvrage_slope_pct_z.get(o)) if ouvrage_slope_pct_z.get(o) is not None else None
            # geometry
            if has_geometry and o in geom_by_ouvrage:
                feat.setGeometry(geom_by_ouvrage[o])
            # insertion dans le sink
            try:
                sink.addFeature(feat, QgsFeatureSink.FastInsert)
            except TypeError:
                sink.addFeature(feat)
            cnt += 1
            feedback.setProgress(int(100 * cnt / total2) if total2 > 0 else 100)

        # ---------------------------
        # --- APPLIQUER LE QML (optionnel) ---
        # ---------------------------
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

        return {self.OUTPUT: dest_id}

# Sauvegarde le script dans Scripts > Tools et lance-le depuis la Toolbox.
