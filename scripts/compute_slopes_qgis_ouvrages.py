# -*- coding: utf-8 -*-
"""
## Objectifs
Calculer, pour chaque ouvrage identifié, l'évolution temporelle des volumes prélevés par année. Produit des indicateurs normalisés : pentes en % par rapport à la moyenne, CAGR (growth rate) et z-score.
## Sortie
Couche (points ou mémoire) contenant par ouvrage : `ouvrage_id`, `slope_ouvrage`, `n_years_ouvrage`, `name_ouv`, `name_petitionaire`, `mean_vol_ouv`, `slope_pct_mean`, `slope_pct_first`, `cagr_pct`, `slope_pct_z`.

## Notes & recommandations
- Theil-Sen est recommandé s'il existe des valeurs aberrantes.
- Calculer les pentes sur des séries avec un nombre minimal d'années (paramétrable).
- Conserver la géométrie du premier point trouvé par ouvrage pour cartographie.

Note d'analyse des indicateurs : 
La _pente_ _(slope)_ mesure l’évolution moyenne absolue du volume prélevé par ouvrage en m³/an (estimée par régression _OLS_ ou _Theil-Sen_) et renseigne l’ampleur physique du changement. Le _slope_pct_mean_ exprime cette pente en pourcentage de la moyenne des volumes de l’ouvrage (100 × slope / mean), ce qui permet de comparer la dynamique relative entre ouvrages de tailles différentes. Le _slope_pct_first_ normalise la pente par rapport au niveau initial, la moyenne des 3 premières années, pour évaluer la variation par rapport au point de départ. Enfin, le _CAGR_ (_taux de croissance annuel composé_) synthétise la croissance équivalente entre une période de départ et une période finale ( moyenne 3 premières vs 3 dernières années) ; il est utile pour résumer une trajectoire début→fin mais masque les fluctuations intermédiaires.

"""

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
    QgsProcessingException
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

    # Ajout du paramètre ZONE
    ZONE = 'ZONE'
    INPUT = 'INPUT'
    YEAR = 'YEAR'
    OUVRAGE = 'OUVRAGE'
    OUV_NAME = 'OUV_NAME'        # nouveau param : champ nom de l'ouvrage (optionnel)
    INTERLOC = 'INTERLOC'       # nouveau param : champ interlocuteur (optionnel)
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
        # couche zone d'étude en première question
        self.addParameter(
            QgsProcessingParameterVectorLayer(
                self.ZONE,
                self.tr("Couche zone d'étude (polygones) - toutes les entités seront prises en compte"),
                [QgsProcessing.TypeVectorPolygon]
            )
        )

        self.addParameter(
            QgsProcessingParameterVectorLayer(self.INPUT, self.tr("Couche d'entrée (points/table)"), [QgsProcessing.TypeVectorAnyGeometry])
        )
        self.addParameter(
            QgsProcessingParameterField(self.YEAR, self.tr("Champ année"), parentLayerParameterName=self.INPUT, type=QgsProcessingParameterField.Numeric)
        )
        self.addParameter(
            QgsProcessingParameterField(self.OUVRAGE, self.tr("Champ identifiant ouvrage"), parentLayerParameterName=self.INPUT)
        )

        # nouveaux paramètres pour le nom de l'ouvrage et l'interlocuteur (tous deux optionnels)
        self.addParameter(
            QgsProcessingParameterField(self.OUV_NAME,
                                       self.tr("Champ nom de l'ouvrage (sera conservé dans la sortie, optionnel)"),
                                       parentLayerParameterName=self.INPUT,
                                       optional=True)
        )
        self.addParameter(
            QgsProcessingParameterField(self.INTERLOC,
                                       self.tr("Champ nom de l'interlocuteur (optionnel)"),
                                       parentLayerParameterName=self.INPUT,
                                       optional=True)
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
        zone_layer = self.parameterAsVectorLayer(parameters, self.ZONE, context)
        if zone_layer is None:
            raise QgsProcessingException(self.tr("La couche zone d'étude n'a pas pu être chargée."))

        layer = self.parameterAsVectorLayer(parameters, self.INPUT, context)
        if layer is None:
            raise QgsProcessingException(self.tr("La couche d'entrée n'a pas pu être chargée."))

        year_field = self.parameterAsString(parameters, self.YEAR, context)
        ouvrage_field = self.parameterAsString(parameters, self.OUVRAGE, context)

        # récupère les noms de champs optionnels ; si vide -> None
        ouvrage_name_field = self.parameterAsString(parameters, self.OUV_NAME, context) if self.OUV_NAME in parameters else None
        if isinstance(ouvrage_name_field, str) and ouvrage_name_field.strip() == '':
            ouvrage_name_field = None
        interloc_field = self.parameterAsString(parameters, self.INTERLOC, context) if self.INTERLOC in parameters else None
        if isinstance(interloc_field, str) and interloc_field.strip() == '':
            interloc_field = None

        vol_field = self.parameterAsString(parameters, self.VOL, context)
        method_idx = self.parameterAsInt(parameters, self.METHOD, context)
        method = ['OLS', 'Theil-Sen'][method_idx]
        min_years = int(self.parameterAsInt(parameters, self.MIN_YEARS, context))
        start_year = int(self.parameterAsInt(parameters, self.START_YEAR, context))
        end_year = int(self.parameterAsInt(parameters, self.END_YEAR, context))
        apply_qml = bool(self.parameterAsBool(parameters, self.APPLY_QML, context))
        qml_path_param = self.parameterAsString(parameters, self.QML_PATH, context)

        # --- Préparer l'index spatial pour la couche zone ---
        zone_feats = []
        try:
            for zf in zone_layer.getFeatures():
                zone_feats.append(zf)
        except Exception:
            zone_feats = []

        if not zone_feats:
            feedback.pushInfo(self.tr("Attention : la couche zone est vide -> aucun filtrage effectué (aucune entité)."))
            zone_index = None
            zone_feat_by_id = {}
        else:
            # crée un index spatial pour accélérer les recherches d'intersection
            try:
                zone_index = QgsSpatialIndex()
                zone_feat_by_id = {}
                for zf in zone_feats:
                    zone_index.insertFeature(zf)
                    zone_feat_by_id[zf.id()] = zf
                feedback.pushInfo(self.tr(f"Index spatial construit pour la couche zone ({len(zone_feats)} entités)."))
            except Exception as e:
                feedback.pushInfo(self.tr(f"Impossible de construire l'index spatial (fallback) : {e}"))
                zone_index = None
                zone_feat_by_id = {zf.id(): zf for zf in zone_feats}

        # lecture et filtrage initial : on ne garde que les prélèvements qui intersectent la zone
        rows = []
        has_geometry = (layer.geometryType() != -1)
        geom_by_ouvrage = {}
        # mappings pour nom & interlocuteur (on garde la valeur associée à la DERNIERE année connue)
        name_by_ouvrage = {}
        name_year_by_ouvrage = {}        # stocke l'année associée à name_by_ouvrage
        interloc_by_ouvrage = {}
        interloc_year_by_ouvrage = {}

        total = layer.featureCount()
        processed = 0
        kept_by_zone = 0
        for f in layer.getFeatures():
            processed += 1
            if feedback.isCanceled():
                break

            # si la couche d'entrée a une géométrie, tester l'intersection avec la zone
            if has_geometry:
                try:
                    fg = f.geometry()
                except Exception:
                    fg = None
                if fg is None or fg.isEmpty():
                    # pas de géométrie -> exclu
                    feedback.setProgress(int(100 * processed / max(1, total)))
                    continue

                intersects_zone = False
                if zone_index is not None:
                    try:
                        # recherche de candidats par bbox
                        cand_ids = zone_index.intersects(fg.boundingBox())
                        for cid in cand_ids:
                            zf = zone_feat_by_id.get(cid)
                            if zf is None:
                                continue
                            try:
                                zg = zf.geometry()
                                if zg is not None and not zg.isEmpty() and zg.intersects(fg):
                                    intersects_zone = True
                                    break
                            except Exception:
                                continue
                    except Exception:
                        # fallback : test direct sur toutes les géométries
                        for zf in zone_feats:
                            try:
                                zg = zf.geometry()
                                if zg is not None and not zg.isEmpty() and zg.intersects(fg):
                                    intersects_zone = True
                                    break
                            except Exception:
                                continue
                else:
                    # pas d'index : brute force
                    for zf in zone_feats:
                        try:
                            zg = zf.geometry()
                            if zg is not None and not zg.isEmpty() and zg.intersects(fg):
                                intersects_zone = True
                                break
                        except Exception:
                            continue

                if not intersects_zone:
                    # non dans la zone -> ignorer
                    feedback.setProgress(int(100 * processed / max(1, total)))
                    continue
                else:
                    kept_by_zone += 1

            # récupérer champs
            try:
                y = f[year_field]
                o = f[ouvrage_field]
                v_raw = f[vol_field]
            except Exception:
                raise Exception(self.tr("Impossible de lire au moins un des champs fournis. Vérifie les paramètres."))

            try:
                yv = int(y)
            except:
                # année non convertible -> ignorer
                feedback.setProgress(int(100 * processed / max(1, total)))
                continue
            if yv < start_year or yv > end_year:
                feedback.setProgress(int(100 * processed / max(1, total)))
                continue

            # récupérer nom & interlocuteur (si champs fournis) -> on garde la valeur de la DERNIERE année
            try:
                if ouvrage_name_field:
                    raw_name = f[ouvrage_name_field]
                    if raw_name is not None and str(raw_name).strip() != '':
                        prev_year = name_year_by_ouvrage.get(o, -9999)
                        # on prend la valeur si l'année courante >= année stockée (garde la plus récente)
                        if yv >= prev_year:
                            name_by_ouvrage[o] = str(raw_name).strip()
                            name_year_by_ouvrage[o] = yv
            except Exception:
                pass
            try:
                if interloc_field:
                    raw_int = f[interloc_field]
                    if raw_int is not None and str(raw_int).strip() != '':
                        prev_year_i = interloc_year_by_ouvrage.get(o, -9999)
                        if yv >= prev_year_i:
                            interloc_by_ouvrage[o] = str(raw_int).strip()
                            interloc_year_by_ouvrage[o] = yv
            except Exception:
                pass

            vv = parse_number(v_raw)
            rows.append((o, yv, vv))
            if has_geometry and o not in geom_by_ouvrage:
                try:
                    geom_by_ouvrage[o] = f.geometry()
                except Exception:
                    pass
            feedback.setProgress(int(100 * processed / max(1, total)))

        feedback.pushInfo(self.tr(f"Prélèvements parcourus: {processed}, conservés après filtrage spatial: {kept_by_zone}, enregistrements retenus pour la période: {len(rows)}."))

        if not rows:
            raise Exception(self.tr("Aucune donnée lue après application du filtre zone / période."))

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
        out_fields.append(QgsField('ouvrage_name', QVariant.String))     # nouveau champ
        out_fields.append(QgsField('interlocuteur', QVariant.String))    # nouveau champ (peut être vide)
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
            # new fields: name & interlocuteur (use stored mappings if present; these are values from the latest year seen)
            try:
                feat['ouvrage_name'] = name_by_ouvrage.get(o, None)
            except Exception:
                feat['ouvrage_name'] = None
            try:
                feat['interlocuteur'] = interloc_by_ouvrage.get(o, None)
            except Exception:
                feat['interlocuteur'] = None

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
                try:
                    feat.setGeometry(geom_by_ouvrage[o])
                except Exception:
                    pass
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
