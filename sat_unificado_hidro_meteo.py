# =============================================================
# MODIFICACIÓN EN FUNCIONES DE EVALUACIÓN Y CLASIFICACIÓN
# =============================================================

def clasificar_alerta_multicriterio(lluvia_24h, viento_kmh, nivel_freatico=None):
    """
    Clasificación semafórica graduada para evitar sobrealertas rojas en el portal.
    Nivel freático: profundidad en metros desde la superficie.
    """
    # 1. Caso de napa casi aflorante (< 0.80 m) + lluvia
    if nivel_freatico is not None and nivel_freatico <= 0.80:
        if lluvia_24h >= 30:
            return "#dc2626", "Rojo - Alerta Severa (Saturación Total)", "Emergencia por escorrentía inminente"
        elif lluvia_24h >= 10:
            return "#ea580c", "Naranja - Alerta Moderada", "Napa alta con lluvias en cuenca"
        else:
            return "#ca8a04", "Amarillo - Atención Freática", "Napa superficial sin lluvias inmediatas"

    # 2. Umbrales puramente meteorológicos graduados
    if lluvia_24h >= 65 or viento_kmh >= 80:
        return "#dc2626", "Rojo - Alerta Severa", "Condiciones severas previstas"
    elif lluvia_24h >= 35 or viento_kmh >= 60:
        return "#ea580c", "Naranja - Alerta Moderada", "Precipitaciones acumuladas importantes"
    elif lluvia_24h >= 15 or viento_kmh >= 40:
        return "#ca8a04", "Amarillo - Vigilancia", "Lluvias / ráfagas leves a moderadas"
    else:
        return "#16a34a", "Verde - Normal", "Condiciones ordinarias estables"
