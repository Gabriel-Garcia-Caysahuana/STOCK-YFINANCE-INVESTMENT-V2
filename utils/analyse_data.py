"""
Módulo para analizar datos de precios y calcular métricas estadísticas y visualizaciones.
Incluye funciones para:
  - Calcular retornos logarítmicos.
  - Generar estadísticas descriptivas.
  - Graficar series temporales y boxplots de retornos.
  - Calcular y graficar matriz de correlación.
  - Calcular volatilidad móvil.
  - Generar histogramas de retornos.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from typing import List, Optional


# -----------------------------------------------------------------------------
# 1) Calcular retornos logarítmicos y estadísticos básicos
# -----------------------------------------------------------------------------
def add_performance(data: pd.DataFrame, tickers: List[str]) -> pd.DataFrame:
    """
    Calcula los retornos logarítmicos para cada ticker y los agrega como columnas
    en el DataFrame con el prefijo 'R'.

    :param data: DataFrame con columnas de precios (Close, o Ajustado, etc.)
    :param tickers: Lista de símbolos de acciones presentes en data.
    :return: DataFrame con nuevas columnas de retornos logarítmicos y sin valores NaN.
    """
    if data.empty:
        print("El DataFrame proporcionado está vacío.")
        return data

    for ticker in tickers:
        if ticker not in data.columns:
            print(f"El ticker '{ticker}' no se encuentra en el DataFrame.")
            continue
        data[f'R{ticker}'] = np.log(data[ticker] / data[ticker].shift(1))

    return data.dropna()


def descriptive(data: pd.DataFrame) -> pd.DataFrame:
    """
    Retorna un resumen estadístico (count, mean, std, min, 25%, 50%, 75%, max)
    de las columnas del DataFrame.

    :param data: DataFrame con datos numéricos.
    :return: DataFrame con el resumen estadístico de cada columna.
    """
    if data.empty:
        print("El DataFrame proporcionado está vacío.")
        return pd.DataFrame()

    return data.describe().round(2).T


# -----------------------------------------------------------------------------
# 2) Análisis de correlación
# -----------------------------------------------------------------------------
def get_correlation_matrix(data: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
    """
    Calcula la matriz de correlación de las columnas especificadas.

    :param data: DataFrame que contiene las columnas de interés (p.ej. retornos).
    :param columns: Lista de nombres de columnas para calcular correlaciones.
    :return: DataFrame con la matriz de correlación.
    """
    if data.empty:
        print("El DataFrame proporcionado está vacío.")
        return pd.DataFrame()

    # Filtramos el DataFrame a solo las columnas deseadas y luego calculamos correlación
    subset = data[columns].dropna()
    if subset.empty:
        print("No se encontraron datos válidos en las columnas solicitadas.")
        return pd.DataFrame()

    corr_matrix = subset.corr()
    return corr_matrix


def plot_correlation_heatmap(
    corr_matrix: pd.DataFrame, 
    title: str = "Matriz de Correlación",
    cmap: str = "RdBu", 
    annot: bool = True
) -> None:
    """
    Genera un mapa de calor (heatmap) a partir de una matriz de correlación.

    :param corr_matrix: DataFrame con la matriz de correlación.
    :param title: Título del gráfico.
    :param cmap: Mapa de color a utilizar (p.ej. 'RdBu', 'coolwarm', 'viridis').
    :param annot: Si True, muestra los valores de correlación en cada celda.
    """
    if corr_matrix.empty:
        print("La matriz de correlación está vacía. No se puede graficar.")
        return
    
    plt.figure(figsize=(8, 6))
    sns.set_style("white")
    sns.heatmap(corr_matrix, vmin=-1, vmax=1, cmap=cmap, annot=annot, square=True, linewidths=.5)
    plt.title(title)
    plt.xticks(rotation=45, ha='right')
    plt.yticks(rotation=0)
    plt.tight_layout()


# -----------------------------------------------------------------------------
# 3) Volatilidad móvil (rolling volatility)
# -----------------------------------------------------------------------------
def rolling_volatility(
    data: pd.DataFrame, 
    ticker: str, 
    window: int = 20, 
    return_col: Optional[str] = None
) -> pd.Series:
    """
    Calcula la volatilidad (desviación estándar) móvil de un activo en base a una ventana.
    Por defecto, se utiliza la columna de retornos logarítmicos 'R{ticker}'.

    :param data: DataFrame con las columnas de retornos.
    :param ticker: Símbolo del activo para el cual se calculará la volatilidad.
    :param window: Tamaño de la ventana móvil (en días, por defecto 20).
    :param return_col: Nombre de la columna de retornos a utilizar. Si no se especifica,
                       se asume 'R{ticker}'.
    :return: Serie con la volatilidad móvil.
    """
    if data.empty:
        print("El DataFrame proporcionado está vacío.")
        return pd.Series(dtype=float)

    col_name = return_col if return_col else f'R{ticker}'
    if col_name not in data.columns:
        print(f"No se encontró la columna de retornos '{col_name}' en el DataFrame.")
        return pd.Series(dtype=float)

    # Desviación estándar móvil
    rolling_std = data[col_name].rolling(window=window).std()
    return rolling_std


def plot_rolling_volatility(
    data: pd.DataFrame,
    ticker: str,
    window: int = 20,
    return_col: Optional[str] = None,
    title: str = "Volatilidad Móvil"
) -> None:
    """
    Grafica la volatilidad (desviación estándar) móvil de un activo en base a una ventana.

    :param data: DataFrame con las columnas de retornos.
    :param ticker: Símbolo del activo para el cual se calculará la volatilidad.
    :param window: Tamaño de la ventana móvil (en días).
    :param return_col: Columna de retornos a utilizar. Si no se especifica, se asume 'R{ticker}'.
    :param title: Título del gráfico.
    """
    vol = rolling_volatility(data, ticker, window, return_col)
    if vol.empty:
        print("No se pudo calcular la volatilidad móvil. No hay datos válidos.")
        return

    sns.set_style("whitegrid")
    plt.figure(figsize=(10, 6))
    plt.plot(vol.index, vol.values, label=f'Volatilidad ({window} días)')
    plt.xlabel("Fecha")
    plt.ylabel("Desviación Estándar")
    plt.title(f"{title} ({ticker})")
    plt.xticks(rotation=45)
    plt.legend()
    plt.tight_layout()


# -----------------------------------------------------------------------------
# 4) Histogramas de retornos
# -----------------------------------------------------------------------------
def plot_returns_histogram(
    data: pd.DataFrame,
    ticker: str,
    bins: int = 30,
    return_col: Optional[str] = None,
    title: str = "Histograma de Retornos"
) -> None:
    """
    Genera un histograma de la distribución de retornos para un activo específico.

    :param data: DataFrame con las columnas de retornos.
    :param ticker: Símbolo del activo para el cual se generará el histograma.
    :param bins: Número de bins para el histograma.
    :param return_col: Columna de retornos a utilizar. Si no se especifica, se asume 'R{ticker}'.
    :param title: Título del gráfico.
    """
    if data.empty:
        print("El DataFrame proporcionado está vacío.")
        return

    col_name = return_col if return_col else f'R{ticker}'
    if col_name not in data.columns:
        print(f"No se encontró la columna de retornos '{col_name}' en el DataFrame.")
        return

    sns.set_style("whitegrid")
    plt.figure(figsize=(10, 6))
    sns.histplot(data[col_name].dropna(), bins=bins, kde=True)
    plt.title(f"{title} - {ticker}")
    plt.xlabel("Retorno")
    plt.ylabel("Frecuencia")
    plt.tight_layout()


# -----------------------------------------------------------------------------
# Gráficos de líneas y boxplots ya existentes
# -----------------------------------------------------------------------------
def plot_line_series(
    data: pd.DataFrame,
    tickers: List[str],
    xlabel: str = 'Fecha',
    ylabel: str = 'Precio Ajustado',
    title: str = 'Precios Históricos de las Acciones'
) -> None:
    """
    Genera una figura de líneas mostrando la evolución de los precios ajustados
    para cada ticker en el tiempo, sin mostrarla en pantalla.

    :param data: DataFrame con datos de precios.
    :param tickers: Lista de símbolos de acciones presentes en data.
    :param xlabel: Etiqueta del eje X.
    :param ylabel: Etiqueta del eje Y.
    :param title: Título del gráfico.
    """
    if data.empty:
        print("El DataFrame proporcionado está vacío.")
        return

    sns.set_style("whitegrid")
    plt.figure(figsize=(10, 6))

    for ticker in tickers:
        if ticker not in data.columns:
            print(f"El ticker '{ticker}' no se encuentra en el DataFrame.")
            continue
        sns.lineplot(x=data.index, y=data[ticker], label=ticker)

    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.title(title)
    plt.legend()
    plt.xticks(rotation=45)


def show_plot_line_series(
    data: pd.DataFrame,
    tickers: List[str],
    xlabel: str = 'Fecha',
    ylabel: str = 'Precio Ajustado',
    title: str = 'Precios Históricos de las Acciones'
) -> None:
    """
    Genera y muestra el gráfico de líneas para los tickers solicitados.
    """
    plot_line_series(data, tickers, xlabel, ylabel, title)
    plt.show()


def get_fig_plot_line_series(
    data: pd.DataFrame,
    tickers: List[str],
    xlabel: str = 'Fecha',
    ylabel: str = 'Precio Ajustado',
    title: str = 'Precios Históricos de las Acciones'
) -> plt.Figure:
    """
    Genera el gráfico de líneas en un objeto Figure y lo retorna
    sin mostrarlo en pantalla.

    :param data: DataFrame con datos de precios.
    :param tickers: Lista de símbolos de acciones presentes en data.
    :param xlabel: Etiqueta del eje X.
    :param ylabel: Etiqueta del eje Y.
    :param title: Título del gráfico.
    :return: Objeto Figure de matplotlib.
    """
    if data.empty:
        print("El DataFrame proporcionado está vacío.")
        return plt.Figure()

    fig, ax = plt.subplots(figsize=(10, 6))
    sns.set_style("whitegrid")

    for ticker in tickers:
        if ticker not in data.columns:
            print(f"El ticker '{ticker}' no se encuentra en el DataFrame.")
            continue
        sns.lineplot(x=data.index, y=data[ticker], label=ticker, ax=ax)

    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    plt.xticks(rotation=45)

    return fig


def plot_box_plot(
    data: pd.DataFrame,
    tickers: List[str],
    title: str = 'Análisis de Box Plot de Retornos',
    ylabel: str = 'Retorno'
) -> None:
    """
    Genera un boxplot de los retornos logarítmicos para los tickers proporcionados,
    sin mostrarlo en pantalla.

    :param data: DataFrame con columnas de retornos logarítmicos (prefijo 'R').
    :param tickers: Lista de símbolos de acciones presentes en data.
    :param title: Título del gráfico.
    :param ylabel: Etiqueta del eje Y.
    """
    if data.empty:
        print("El DataFrame proporcionado está vacío.")
        return

    columns_for_boxplot = []
    for ticker in tickers:
        ret_col = f'R{ticker}'
        if ret_col in data.columns:
            columns_for_boxplot.append(ret_col)
        else:
            print(f"No se encontró la columna de retornos '{ret_col}' en el DataFrame.")

    if not columns_for_boxplot:
        print("No hay columnas de retornos para graficar.")
        return

    sns.set_style("whitegrid")
    plt.figure(figsize=(10, 6))
    sns.boxplot(data=data[columns_for_boxplot])
    plt.title(title)
    plt.ylabel(ylabel)


def show_plot_box_plot(
    data: pd.DataFrame,
    tickers: List[str],
    title: str = 'Análisis de Box Plot de Retornos',
    ylabel: str = 'Retorno'
) -> None:
    """
    Genera y muestra el boxplot de los retornos logarítmicos para los tickers solicitados.
    """
    plot_box_plot(data, tickers, title, ylabel)
    plt.show()


def get_fig_plot_box_plot(
    data: pd.DataFrame,
    tickers: List[str],
    title: str = 'Análisis de Box Plot de Retornos',
    ylabel: str = 'Retorno'
) -> plt.Figure:
    """
    Genera el boxplot de los retornos logarítmicos en un objeto Figure y lo retorna,
    sin mostrarlo en pantalla.

    :param data: DataFrame con columnas de retornos logarítmicos (prefijo 'R').
    :param tickers: Lista de símbolos de acciones presentes en data.
    :param title: Título del gráfico.
    :param ylabel: Etiqueta del eje Y.
    :return: Objeto Figure de matplotlib.
    """
    if data.empty:
        print("El DataFrame proporcionado está vacío.")
        return plt.Figure()

    columns_for_boxplot = []
    for ticker in tickers:
        ret_col = f'R{ticker}'
        if ret_col in data.columns:
            columns_for_boxplot.append(ret_col)
        else:
            print(f"No se encontró la columna de retornos '{ret_col}' en el DataFrame.")

    if not columns_for_boxplot:
        print("No hay columnas de retornos para graficar.")
        return plt.Figure()

    fig, ax = plt.subplots(figsize=(10, 6))
    sns.set_style("whitegrid")
    sns.boxplot(data=data[columns_for_boxplot], ax=ax)
    ax.set_title(title)
    ax.set_ylabel(ylabel)

    return fig
