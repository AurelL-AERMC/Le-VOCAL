# -*- coding: utf-8 -*-
"""
Initialisation du plugin Prelev_Mapping_AERMC_V2
La fonction classFactory(iface) doit retourner une instance de la classe
principale du plugin (ici PrelevOrchestrator).
"""

def classFactory(iface):
    """
    Cette fonction est appelée par QGIS au chargement du plugin.
    Elle doit retourner une instance de la classe principale du plugin,
    qui prend 'iface' (QgisInterface) en argument.
    """
    # Import local (retarde l'import du module jusqu'à l'appel de QGIS)
    from .prelev_orchestrator import PrelevOrchestratorPlugin
    return PrelevOrchestratorPlugin(iface)
