# -*- coding: utf-8 -*-
"""
## Objectifs
Pour une année donnée, comparer le volume prélevé (VP, "assiettes" retenues à l'Agence) à un volume autorisé (VA, arrétés de déclaration ou d'autorisation DDTM). Déterminer les dépassements et fournir des indicateurs de ce ratio.
## Traitement
- Filtrage spatial (zone) si la couche zone a des géométries.
- Agréger les volumes par ID ouvrage pour l'année choisie.
- Joindre avec la table autorisée : prendre `MAX(VA)` si plusieurs enregistrements, concaténer champs DDTM distincts.
- Calculer `ratio = VP / VA` (si VA non nul) et `% overrun`.

## Sortie
Couche par ouvrage pour l'année choisie : `annee`, `ouvrage_id`, `ouvrage_name`, `interlocuteur`, `assiette`, `vol_autorise`, `ddtm_id`, `ratio`, `ratio_possible`, `percent_overrun`, `note`, `type_milieu`.

## Note sur les indicateurs
- Le ratio représente réellement la division du VP/VA
- Le %overrun présente le pourcentage que représente le VP/VA. 
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
    QgsFeatureSink,
    QgsSpatialIndex
)
import re
import os
import math
from collections import defaultdict

# -------- utilitaires --------
def parse_number(x):
    """
    Parse un nombre donné au format français/anglais :
    - Accepte "12 000,56" ou "12000.56" ou "12000,56"
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
    # remove spaces thousands separators, keep punctuation
    s_nosp = s.replace(' ', '')
    if '.' in s_nosp and ',' in s_nosp:
        # if dot before comma -> dot = thousands, comma = decimal
        if s_nosp.find('.') < s_nosp.find(','):
            s_clean = s_nosp.replace('.', '').replace(',', '.')
        else:
            # uncommon: comma thousands, dot decimal
            s_clean = s_nosp.replace(',', '')
    elif ',' in s_nosp:
        s_clean = s_nosp.replace(',', '.')
    else:
        s_clean = s_nosp
    # keep digits, dot and minus
    s_clean = re.sub(r'[^0-9\.\-]', '', s_clean)
    if s_clean in ['', '.', '-', '-.']:
        return float('nan')
    try:
        return float(s_clean)
    except:
        return float('nan')

def parse_year_to_int(y_raw):
    """
    Try to convert the year value to an int.
    Accepts ints, numeric strings, strings containing a 4-digit year.
    Returns int or None.
    """
    if y_raw is None:
        return None
    # if already int
    if isinstance(y_raw, int):
        return y_raw
    # if float but integer value
    if isinstance(y_raw, float):
        try:
            return int(y_raw)
        except:
            pass
    s = str(y_raw).strip()
    if s == '':
        return None
    # try direct int
    try:
        return int(float(s))
    except:
        pass
    # search for 4-digit year
    m = re.search(r'(\d{4})', s)
    if m:
        try:
            return int(m.group(1))
        except:
            return None
    return None

# -------- Algorithm --------
class ComparePrelevementsAutorises(QgsProcessingAlgorithm):
    """
    Algorithme Processing : compare volumes prélevés vs volumes autorisés pour une année donnée.
    Conserve le champ 'type de milieu' provenant de la couche prélèvements.
    """

    # paramètres
    ZONE = 'ZONE'  # nouvelle couche zone d'étude (polygones)
    PRELEV = 'PRELEV'
    PRELEV_YEAR_FIELD = 'PRELEV_YEAR_FIELD'
    PRELEV_OUV_FIELD = 'PRELEV_OUV_FIELD'
    PRELEV_ASSIETTE_FIELD = 'PRELEV_ASSIETTE_FIELD'
    PRELEV_MILIEU_FIELD = 'PRELEV_MILIEU_FIELD'  # nouveau param : champ type de milieu

    # nouveaux params optionnels pour nom ouvrage et interlocuteur
    PRELEV_OUV_NAME = 'PRELEV_OUV_NAME'
    PRELEV_INTERLOC = 'PRELEV_INTERLOC'

    AUTOR = 'AUTOR'
    AUTOR_OUV_FIELD = 'AUTOR_OUV_FIELD'
    AUTOR_VOL_FIELD = 'AUTOR_VOL_FIELD'
    AUTOR_DDTM_FIELD = 'AUTOR_DDTM_FIELD'  # optional

    YEAR = 'YEAR'
    INCLUDE_UNMATCHED = 'INCLUDE_UNMATCHED'
    APPLY_QML = 'APPLY_QML'
    QML_PATH = 'QML_PATH'
    OUTPUT = 'OUTPUT'

    def tr(self, s):
        return s

    def createInstance(self):
        return ComparePrelevementsAutorises()

    def name(self):
        return 'compare_prelevements_autorises'

    def displayName(self):
        return self.tr('Comparer prélèvements vs volumes autorisés (année unique)')

    def group(self):
        return self.tr('Analyses temporelles')

    def groupId(self):
        return 'temporal_analysis'

    def shortHelpString(self):
        return self.tr(
            "Agrège les volumes prélevés pour une année donnée par ID ouvrage, joint avec la table des volumes autorisés, "
            "calcule ratio et % dépassement. Conserve le champ 'type de milieu' (concaténation si plusieurs valeurs). "
            "Demande une couche de zone d'étude et ne conserve que les prélèvements situés dans cette zone. "
            "Si l'année renseignée est 0 (valeur par défaut), le script utilisera la dernière année disponible parmi les prélèvements retenus."
        )

    def initAlgorithm(self, config=None):
        # nouvelle : couche zone d'étude (polygones)
        self.addParameter(
            QgsProcessingParameterVectorLayer(
                self.ZONE,
                self.tr("Couche zone d'étude (polygones)"),
                [QgsProcessing.TypeVectorPolygon]
            )
        )

        # couche prélèvements (points/table)
        self.addParameter(
            QgsProcessingParameterVectorLayer(
                self.PRELEV,
                self.tr("Couche prélèvements (points ou table)"),
                [QgsProcessing.TypeVectorAnyGeometry]
            )
        )
        self.addParameter(
            QgsProcessingParameterField(
                self.PRELEV_YEAR_FIELD,
                self.tr("Champ année (prélèvements)"),
                parentLayerParameterName=self.PRELEV,
                type=QgsProcessingParameterField.Any
            )
        )
        self.addParameter(
            QgsProcessingParameterField(
                self.PRELEV_OUV_FIELD,
                self.tr("Champ ID Ouvrage (prélèvements)"),
                parentLayerParameterName=self.PRELEV
            )
        )
        self.addParameter(
            QgsProcessingParameterField(
                self.PRELEV_ASSIETTE_FIELD,
                self.tr("Champ Assiette (volume prélevé)"),
                parentLayerParameterName=self.PRELEV
            )
        )
        # nouveau : champ type de milieu (optionnel)
        self.addParameter(
            QgsProcessingParameterField(
                self.PRELEV_MILIEU_FIELD,
                self.tr("Champ 'type de milieu' (prélèvements) - optionnel"),
                parentLayerParameterName=self.PRELEV,
                optional=True
            )
        )

        # nouveaux champs optionnels : nom ouvrage & interlocuteur
        self.addParameter(
            QgsProcessingParameterField(
                self.PRELEV_OUV_NAME,
                self.tr("Champ nom de l'ouvrage (optionnel, conservé dans la sortie)"),
                parentLayerParameterName=self.PRELEV,
                optional=True
            )
        )
        self.addParameter(
            QgsProcessingParameterField(
                self.PRELEV_INTERLOC,
                self.tr("Champ interlocuteur (optionnel, conservé dans la sortie)"),
                parentLayerParameterName=self.PRELEV,
                optional=True
            )
        )

        # couche volumes autorisés (table ou couche)
        self.addParameter(
            QgsProcessingParameterVectorLayer(
                self.AUTOR,
                self.tr("Table / couche volumes autorisés"),
                [QgsProcessing.TypeVectorAnyGeometry]
            )
        )
        self.addParameter(
            QgsProcessingParameterField(
                self.AUTOR_OUV_FIELD,
                self.tr("Champ ID Ouvrage (autorises) - pour la jointure"),
                parentLayerParameterName=self.AUTOR
            )
        )
        self.addParameter(
            QgsProcessingParameterField(
                self.AUTOR_VOL_FIELD,
                self.tr("Champ Volume autorisé (autorises)"),
                parentLayerParameterName=self.AUTOR
            )
        )
        self.addParameter(
            QgsProcessingParameterField(
                self.AUTOR_DDTM_FIELD,
                self.tr("Champ Identifiant DDTM (autorises) - optionnel"),
                parentLayerParameterName=self.AUTOR,
                optional=True
            )
        )

        # autres paramètres
        # YEAR: default 0 => use last available year after spatial filtering
        self.addParameter(
            QgsProcessingParameterNumber(
                self.YEAR,
                self.tr("Année (ex: 2023). Mettre 0 pour utiliser la dernière année disponible"),
                type=QgsProcessingParameterNumber.Integer,
                defaultValue=0
            )
        )
        self.addParameter(
            QgsProcessingParameterBoolean(
                self.INCLUDE_UNMATCHED,
                self.tr("Inclure les ouvrages prélevés sans enregistrement autorisé ?"),
                defaultValue=True
            )
        )
        self.addParameter(
            QgsProcessingParameterBoolean(
                self.APPLY_QML,
                self.tr("Appliquer un style QML sur la couche de sortie ?"),
                defaultValue=True
            )
        )
        default_qml = r"N:\_MTP\Public\01-ORGANISATION\G-Services\RAGAF\REDEVANCES\Recherche redevables\Etude données prélèvements\06_Valorisation_Visualisation\Outils\QML\QML_ratio_VPVA_ouvrages.qml"
        self.addParameter(
            QgsProcessingParameterString(
                self.QML_PATH,
                self.tr("Chemin du fichier QML (si appliqué)"),
                defaultValue=default_qml
            )
        )

        self.addParameter(
            QgsProcessingParameterFeatureSink(
                self.OUTPUT,
                self.tr("Couche de sortie (comparaison prélèvements vs autorisés)")
            )
        )

    def processAlgorithm(self, parameters, context, feedback):
        # read parameters
        zone_lyr = self.parameterAsVectorLayer(parameters, self.ZONE, context)
        prelev_lyr = self.parameterAsVectorLayer(parameters, self.PRELEV, context)
        prelev_year_field = self.parameterAsString(parameters, self.PRELEV_YEAR_FIELD, context)
        prelev_ouv_field = self.parameterAsString(parameters, self.PRELEV_OUV_FIELD, context)
        prelev_assiette_field = self.parameterAsString(parameters, self.PRELEV_ASSIETTE_FIELD, context)
        prelev_milieu_field = self.parameterAsString(parameters, self.PRELEV_MILIEU_FIELD, context) if self.PRELEV_MILIEU_FIELD in parameters else None

        # optional name & interloc fields
        prelev_ouv_name_field = None
        try:
            prelev_ouv_name_field = self.parameterAsString(parameters, self.PRELEV_OUV_NAME, context)
            if prelev_ouv_name_field == '':
                prelev_ouv_name_field = None
        except Exception:
            prelev_ouv_name_field = None
        prelev_interloc_field = None
        try:
            prelev_interloc_field = self.parameterAsString(parameters, self.PRELEV_INTERLOC, context)
            if prelev_interloc_field == '':
                prelev_interloc_field = None
        except Exception:
            prelev_interloc_field = None

        autor_lyr = self.parameterAsVectorLayer(parameters, self.AUTOR, context)
        autor_ouv_field = self.parameterAsString(parameters, self.AUTOR_OUV_FIELD, context)
        autor_vol_field = self.parameterAsString(parameters, self.AUTOR_VOL_FIELD, context)
        autor_ddtm_field = self.parameterAsString(parameters, self.AUTOR_DDTM_FIELD, context) if self.AUTOR_DDTM_FIELD in parameters else None

        year_param_input = int(self.parameterAsInt(parameters, self.YEAR, context))
        include_unmatched = bool(self.parameterAsBool(parameters, self.INCLUDE_UNMATCHED, context))
        apply_qml = bool(self.parameterAsBool(parameters, self.APPLY_QML, context))
        qml_path_param = self.parameterAsString(parameters, self.QML_PATH, context)

        feedback.pushInfo(self.tr(f"Paramètres : année={year_param_input} (0 => dernière dispo), inclure_unmatched={include_unmatched}, apply_qml={apply_qml}"))

        # Validate zone layer
        if zone_lyr is None:
            raise Exception(self.tr("Paramètre 'Zone d'étude' manquant ou invalide."))
        if prelev_lyr is None:
            raise Exception(self.tr("Paramètre 'Couche prélèvements' manquant ou invalide."))

        if zone_lyr.featureCount() == 0:
            feedback.pushInfo(self.tr("La couche zone d'étude est vide (0 entité). Aucun prélèvement ne sera retenu."))

        # Build spatial index for zone layer (if polygon geometry available)
        zone_index = None
        zone_geoms = {}
        try:
            # create index only if zone has geometries
            if zone_lyr.geometryType() != -1 and zone_lyr.featureCount() > 0:
                zone_index = QgsSpatialIndex()
                for zf in zone_lyr.getFeatures():
                    try:
                        zg = zf.geometry()
                        if zg is None or zg.isEmpty():
                            continue
                        zone_geoms[zf.id()] = zg
                        zone_index.addFeature(zf)
                    except Exception:
                        continue
                feedback.pushInfo(self.tr(f"Index spatial zone construit ({len(zone_geoms)} géométries)."))
            else:
                feedback.pushInfo(self.tr("La couche zone d'étude n'a pas de géométrie exploitable. Filtrage spatial désactivé."))
                zone_index = None
        except Exception as e:
            feedback.pushInfo(self.tr(f"Erreur création index spatial zone : {e}"))
            zone_index = None

        # 1) lire la table des volumes autorisés et construire un index par ID ouvrage
        #    -> prendre MAX(volume autorisé) si plusieurs enregistrements, concatener DDTM distincts
        autor_index = {}  # key (str id) -> dict { 'vol_max': float, 'ddtm': set(...) }
        autor_count = 0
        for f in autor_lyr.getFeatures():
            autor_count += 1
            key_raw = f[autor_ouv_field]
            if key_raw is None:
                continue
            key = str(key_raw).strip()
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
            entry = autor_index.get(key)
            if entry is None:
                dd = set()
                if ddtm_val:
                    dd.add(ddtm_val)
                autor_index[key] = {'vol_max': vol if not math.isnan(vol) else float('nan'), 'ddtm': dd}
            else:
                # update vol_max if numeric and larger
                try:
                    if not math.isnan(vol):
                        if math.isnan(entry['vol_max']) or vol > entry['vol_max']:
                            entry['vol_max'] = vol
                except Exception:
                    pass
                # add ddtm
                if ddtm_val:
                    entry['ddtm'].add(ddtm_val)
        feedback.pushInfo(self.tr(f"Chargé {autor_count} enregistrements volumes autorisés -> index de {len(autor_index)} clés."))

        # 2) parcourir les prélèvements : 1ère passe = filtrage spatial + collecte des années disponibles (si YEAR=0)
        prelev_count = 0
        kept_spatial = 0
        prelev_has_geom = (prelev_lyr.geometryType() != -1)
        records = []  # store tuples for second pass: (key, year_int, ass_raw, geom, milieu_raw, name_raw, interloc_raw)
        available_years = set()

        for f in prelev_lyr.getFeatures():
            prelev_count += 1
            if feedback.isCanceled():
                break

            # spatial filter if applicable
            if prelev_has_geom and zone_index is not None:
                try:
                    fg = f.geometry()
                    if fg is None or fg.isEmpty():
                        continue
                    cands = zone_index.intersects(fg.boundingBox())
                    if not cands:
                        continue
                    inside = False
                    for fid in cands:
                        zg = zone_geoms.get(fid)
                        if zg is None:
                            try:
                                zf_tmp = zone_lyr.getFeature(fid)
                                zg = zf_tmp.geometry() if zf_tmp is not None else None
                            except Exception:
                                zg = None
                        if zg is None:
                            continue
                        try:
                            if zg.contains(fg) or zg.intersects(fg):
                                inside = True
                                break
                        except Exception:
                            try:
                                if zg.intersects(fg):
                                    inside = True
                                    break
                            except Exception:
                                continue
                    if not inside:
                        continue
                except Exception:
                    # on erreur, exclure la géométrie
                    continue
            # passed spatial filter (or no spatial filtering applied)
            kept_spatial += 1

            # read year (may be string)
            try:
                y_raw = f[prelev_year_field]
            except Exception:
                continue
            y_int = parse_year_to_int(y_raw)
            if y_int is None:
                # keep record? no - it's unusable for year selection/aggregation
                continue

            # store available year
            available_years.add(y_int)

            # read key and other raw fields (we will filter by year later)
            try:
                key_raw = f[prelev_ouv_field]
                if key_raw is None:
                    continue
                key = str(key_raw).strip()
            except Exception:
                continue

            # assiette raw
            try:
                ass_raw = f[prelev_assiette_field]
            except Exception:
                ass_raw = None

            # milieu raw (optional)
            milieu_raw = None
            if prelev_milieu_field:
                try:
                    milieu_raw = f[prelev_milieu_field]
                except Exception:
                    milieu_raw = None

            # name & interloc raw (optional)
            name_raw = None
            if prelev_ouv_name_field:
                try:
                    name_raw = f[prelev_ouv_name_field]
                except Exception:
                    name_raw = None
            interloc_raw = None
            if prelev_interloc_field:
                try:
                    interloc_raw = f[prelev_interloc_field]
                except Exception:
                    interloc_raw = None

            # geometry
            geom = None
            if prelev_has_geom:
                try:
                    geom = f.geometry()
                except Exception:
                    geom = None

            records.append((key, y_int, ass_raw, geom, milieu_raw, name_raw, interloc_raw))
            feedback.setProgress(int(100 * prelev_count / max(1, prelev_lyr.featureCount())))

        feedback.pushInfo(self.tr(f"Prélèvements parcourus: {prelev_count}, conservés après filtrage spatial: {kept_spatial}, années disponibles: {sorted(available_years)}"))

        # determine year to use
        if year_param_input == 0:
            if not available_years:
                raise Exception(self.tr("Aucune année disponible parmi les prélèvements retenus — impossible de déterminer la dernière année."))
            year_param = max(available_years)
            feedback.pushInfo(self.tr(f"Aucune année fournie (0) -> usage de la dernière année disponible : {year_param}"))
        else:
            year_param = int(year_param_input)
            feedback.pushInfo(self.tr(f"Année fournie par l'utilisateur : {year_param}"))

        # 3) deuxième passe : agréger assiette par ouvrage pour l'année choisie, collecter géom, milieu, name, interloc
        assiette_by_ouv = defaultdict(float)
        geom_by_ouv = {}
        milieu_by_ouv = defaultdict(set)
        name_by_ouv = {}
        interloc_by_ouv = {}

        for rec in records:
            key, y_int, ass_raw, geom, milieu_raw, name_raw, interloc_raw = rec
            if y_int != year_param:
                continue
            # assiette
            ass = parse_number(ass_raw)
            if math.isnan(ass):
                ass_val = 0.0
            else:
                ass_val = ass
            assiette_by_ouv[key] += ass_val
            # geometry -> keep first geometry found
            if prelev_has_geom and key not in geom_by_ouv and geom is not None and not geom.isEmpty():
                geom_by_ouv[key] = geom
            # milieu
            if prelev_milieu_field and milieu_raw is not None:
                mm = str(milieu_raw).strip()
                if mm != '':
                    milieu_by_ouv[key].add(mm)
            # name (first non-empty)
            if prelev_ouv_name_field and name_raw is not None:
                try:
                    nm = str(name_raw).strip()
                    if nm != '' and key not in name_by_ouv:
                        name_by_ouv[key] = nm
                except Exception:
                    pass
            # interloc (first non-empty)
            if prelev_interloc_field and interloc_raw is not None:
                try:
                    it = str(interloc_raw).strip()
                    if it != '' and key not in interloc_by_ouv:
                        interloc_by_ouv[key] = it
                except Exception:
                    pass

        feedback.pushInfo(self.tr(f"Ouvrages agrégés pour l'année {year_param} : {len(assiette_by_ouv)}"))

        # 4) pour chaque ouvrage agrégé, joindre avec autor_index
        rows_out = []  # tuples of (key, annee, assiette_sum, vol_autorise, ddtm_concat, ratio, ratio_possible, percent_overrun, note, geom, milieu_concat, name, interloc)
        cnt_included = 0
        cnt_unmatched = 0
        cnt_vol_zero = 0
        for key, ass_sum in sorted(assiette_by_ouv.items()):
            autor_entry = autor_index.get(key)
            if autor_entry is None:
                # unmatched
                if not include_unmatched:
                    cnt_unmatched += 1
                    continue
                vol_auth = None
                ddtm_concat = None
                note = 'unmatched'
            else:
                vol_auth = autor_entry.get('vol_max')
                if vol_auth is None or (isinstance(vol_auth, float) and math.isnan(vol_auth)):
                    vol_auth = None
                ddset = autor_entry.get('ddtm', set())
                ddtm_concat = ';'.join(sorted(ddset)) if ddset else None
                note = 'matched'
            # ratio logic
            ratio_possible = False
            ratio = None
            percent_overrun = None
            if vol_auth is None:
                ratio_possible = False
            else:
                # vol_auth numeric
                try:
                    if vol_auth == 0:
                        ratio_possible = False
                        cnt_vol_zero += 1
                    else:
                        ratio = ass_sum / vol_auth
                        ratio_possible = True
                        percent_overrun = ((ass_sum - vol_auth) / vol_auth) * 100.0
                except Exception:
                    ratio_possible = False
            geom = geom_by_ouv.get(key) if key in geom_by_ouv else None
            # milieu concat
            milset = milieu_by_ouv.get(key, set())
            milieu_concat = ';'.join(sorted(milset)) if milset else None
            # name/interloc
            nm = name_by_ouv.get(key) if name_by_ouv.get(key) is not None else None
            itc = interloc_by_ouv.get(key) if interloc_by_ouv.get(key) is not None else None
            rows_out.append((key, year_param, ass_sum, vol_auth, ddtm_concat, ratio, 1 if ratio_possible else 0, percent_overrun, note, geom, milieu_concat, nm, itc))
            cnt_included += 1

        feedback.pushInfo(self.tr(f"Ouvrages inclus dans la sortie : {cnt_included} (non appariés exclus: {cnt_unmatched}) ; vols autorisés nuls: {cnt_vol_zero}"))

        # 5) préparer sink et écrire la couche de sortie (géométrie = de la couche prélèvements si disponible)
        out_fields = QgsFields()
        out_fields.append(QgsField('annee', QVariant.Int))
        out_fields.append(QgsField('ouvrage_id', QVariant.String))
        out_fields.append(QgsField('ouvrage_name', QVariant.String))     # nouveau champ
        out_fields.append(QgsField('interlocuteur', QVariant.String))    # nouveau champ
        out_fields.append(QgsField('assiette', QVariant.Double))
        out_fields.append(QgsField('vol_autorise', QVariant.Double))
        out_fields.append(QgsField('ddtm_id', QVariant.String))
        out_fields.append(QgsField('ratio', QVariant.Double))
        out_fields.append(QgsField('ratio_possible', QVariant.Int))  # 1/0
        out_fields.append(QgsField('percent_overrun', QVariant.Double))
        out_fields.append(QgsField('note', QVariant.String))
        out_fields.append(QgsField('type_milieu', QVariant.String))  # nouveau champ de sortie

        # geometry type from prelev layer (points or None -> use wkbType)
        wkbtype = prelev_lyr.wkbType()
        crs = prelev_lyr.sourceCrs()

        (sink, dest_id) = self.parameterAsSink(parameters, self.OUTPUT, context,
                                               out_fields, wkbtype, crs)

        total_rows = len(rows_out)
        written = 0
        for i, rec in enumerate(rows_out):
            if feedback.isCanceled():
                break
            key, annee, ass_sum, vol_auth, ddtm_concat, ratio, ratio_possible_int, percent_overrun, note, geom, milieu_concat, nm, itc = rec
            feat = QgsFeature()
            feat.setFields(out_fields)
            feat['annee'] = int(annee)
            feat['ouvrage_id'] = str(key)
            feat['ouvrage_name'] = str(nm) if nm is not None else None
            feat['interlocuteur'] = str(itc) if itc is not None else None
            feat['assiette'] = float(ass_sum) if ass_sum is not None else None
            feat['vol_autorise'] = float(vol_auth) if vol_auth is not None else None
            feat['ddtm_id'] = str(ddtm_concat) if ddtm_concat is not None else None
            feat['ratio'] = float(ratio) if ratio is not None else None
            feat['ratio_possible'] = int(ratio_possible_int)
            feat['percent_overrun'] = float(percent_overrun) if percent_overrun is not None else None
            feat['note'] = str(note)
            feat['type_milieu'] = str(milieu_concat) if milieu_concat is not None else None
            if geom is not None:
                try:
                    feat.setGeometry(geom)
                except Exception:
                    pass
            try:
                sink.addFeature(feat, QgsFeatureSink.FastInsert)
            except TypeError:
                sink.addFeature(feat)
            written += 1
            feedback.setProgress(int(100 * written / max(1, total_rows)))

        feedback.pushInfo(self.tr(f"Ecriture terminée : {written} entités écrites."))

        # 6) appliquer QML si demandé
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
            feedback.pushInfo(self.tr(f"Erreur lors de l'application du style QML : {e}"))

        return {self.OUTPUT: dest_id}

# End of script
