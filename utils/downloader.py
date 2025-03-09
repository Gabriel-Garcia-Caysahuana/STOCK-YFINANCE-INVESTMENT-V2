"""
Módulo para descargar datos de precios ajustados de acciones desde Yahoo Finance.
"""

import yfinance as yf
import pandas as pd
from typing import List, Optional


def download_data(
    tickers: List[str],
    date_start: str,
    date_end: str
) -> Optional[pd.DataFrame]:
    """
    Descarga los precios ajustados ("Adj Close") de una lista de tickers
    entre dos fechas dadas, haciendo uso de la biblioteca yfinance.

    :param tickers: Lista de símbolos de acciones.
    :param date_start: Fecha de inicio en formato 'YYYY-MM-DD'.
    :param date_end: Fecha de fin en formato 'YYYY-MM-DD'.
    :return: DataFrame con los precios ajustados de cada ticker.
             Retorna None si no se puede descargar información.
    """
    if not tickers:
        print("La lista de tickers está vacía.")
        return None

    try:
        data = yf.download(
            tickers,
            start=date_start,
            end=date_end,
            timeout=60,
            progress=False,
            auto_adjust=False
        )["Adj Close"].dropna(how='all', axis=0)

        # Si el DataFrame está vacío, retornamos None
        if data.empty:
            print("No se obtuvieron datos para los tickers proporcionados.")
            return None

        # Si se solicita un solo ticker, la descarga regresa una Serie
        # Convertimos a DataFrame para mantener consistencia
        if isinstance(data, pd.Series):
            data = data.to_frame()

        return data

    except Exception as e:
        print(f"Ocurrió un error al descargar datos: {e}")
        return None
