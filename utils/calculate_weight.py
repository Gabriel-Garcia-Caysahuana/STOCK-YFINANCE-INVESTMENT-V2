"""
Módulo para calcular pesos óptimos de un portafolio utilizando PyPortfolioOpt.
"""

import pandas as pd
from pypfopt import EfficientFrontier, risk_models, expected_returns
from typing import List, Tuple, Dict

def calculate_mu_sigma(data: pd.DataFrame, tickers: List[str]) -> Tuple[pd.Series, pd.DataFrame]:
    """
    Calcula la media de retornos anualizados (mu) y la matriz de covarianza (sigma)
    para una lista de tickers dada, usando retornos logarítmicos.

    :param data: DataFrame con precios de cierre o ajustados, por columnas.
    :param tickers: Lista de símbolos de acciones.
    :return: Una tupla (mu, sigma) con:
             - mu: Serie de medias de retornos anualizados.
             - sigma: DataFrame de covarianza de los retornos.
    """
    if data.empty:
        print("El DataFrame proporcionado está vacío.")
        return pd.Series(dtype=float), pd.DataFrame()

    # Creamos un DataFrame 'precios' solo con las columnas de los tickers solicitados
    precios = pd.DataFrame()
    for ticker in tickers:
        if ticker in data.columns:
            precios[ticker] = data[ticker]
        else:
            print(f"El ticker '{ticker}' no se encuentra en el DataFrame.")

    if precios.empty:
        print("No se encontraron datos para ninguno de los tickers solicitados.")
        return pd.Series(dtype=float), pd.DataFrame()

    # Cálculo de retornos esperados y matriz de covarianza
    mu = expected_returns.mean_historical_return(
        precios,
        frequency=252,
        log_returns=True
    )
    sigma = risk_models.sample_cov(precios)

    return mu, sigma

def calculate_weight(mu: pd.Series, sigma: pd.DataFrame) -> Dict[str, float]:
    """
    Calcula la asignación óptima de pesos para maximizar el ratio de Sharpe,
    retornando un diccionario con los pesos de cada activo.

    :param mu: Serie de retornos anualizados.
    :param sigma: Matriz de covarianza de los retornos.
    :return: Diccionario con los pesos óptimos de cada activo (e.g., {'AAPL': 0.25, 'TSLA': 0.75}).
    """
    if mu.empty:
        print("El objeto mu está vacío. No se pueden calcular pesos.")
        return {}

    if sigma.empty:
        print("El objeto sigma está vacío. No se pueden calcular pesos.")
        return {}

    try:
        ef = EfficientFrontier(mu, sigma)
        ef.max_sharpe()  # Ajusta los pesos internamente
        cleaned_weights = ef.clean_weights()  # Retorna un dict
        ef.portfolio_performance(verbose=True)

        # Ahora retornamos directamente el diccionario
        return cleaned_weights

    except Exception as e:
        print(f"Ocurrió un error al calcular los pesos: {e}")
        return {}
