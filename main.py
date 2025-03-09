"""
Programa principal para el análisis de acciones bursátiles:
 - Descarga de datos con yfinance.
 - Cálculo de estadísticas descriptivas y retornos logarítmicos.
 - Visualización de gráficos.
 - Cálculo de pesos óptimos de portafolio.
 - Generación de informe en Word y Excel.
"""

import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from typing import List
from datetime import datetime

# Módulos de utilidades
from utils.downloader import download_data 
from utils.analyse_data import add_performance, descriptive, show_plot_line_series, show_plot_box_plot
from utils.calculate_weight import calculate_mu_sigma, calculate_weight
from utils.to_generate_word import generar_informe
from utils.to_generate_excel import generar_excel
import matplotlib.pyplot as plt


# Para mostrar tablas y aplicar temas de color
from prettytable.colortable import ColorTable, Themes
import pandas as pd

def print_table(df: pd.DataFrame, theme=Themes.OCEAN) -> None:
    """
    Imprime un DataFrame en formato de tabla con un tema de color específico.

    :param df: DataFrame con la información a mostrar.
    :param theme: Tema de color a utilizar. Por defecto, Themes.OCEAN.
    """
    table = ColorTable(theme=theme)
    table.field_names = df.columns.tolist()
    table.add_rows(df.values.tolist())
    print(table)

def print_with_color_and_format(
    text: str,
    color_code: int,
    bold: bool = False,
    underline: bool = False
) -> None:
    """
    Imprime texto con formato específico (color, negrita, subrayado).

    :param text: Texto a imprimir.
    :param color_code: Código de color ANSI (ej. 31=rojo, 32=verde, 33=amarillo, 34=azul).
    :param bold: Activa el formato de negrita.
    :param underline: Activa el formato de subrayado.
    """
    start_code = "\033["
    if bold:
        start_code += "1;"
    if underline:
        start_code += "4;"
    start_code += f"{color_code}m"
    end_code = "\033[0m"
    print(f"{start_code}{text}{end_code}")

def validar_fecha(fecha_str: str) -> bool:
    """
    Valida si una cadena de texto está en el formato correcto de fecha (YYYY-MM-DD).

    :param fecha_str: Cadena de texto con la fecha a validar.
    :return: True si cumple con el formato, False en caso contrario.
    """
    try:
        datetime.strptime(fecha_str, "%Y-%m-%d")
        return True
    except ValueError:
        return False

def solicitar_fecha(mensaje: str) -> str:
    """
    Solicita al usuario que ingrese una fecha válida en formato YYYY-MM-DD.

    :param mensaje: Mensaje que se muestra al usuario.
    :return: Cadena de texto que representa la fecha validada.
    """
    while True:
        fecha = input(mensaje).strip()
        if validar_fecha(fecha):
            return fecha
        else:
            print_with_color_and_format("Formato de fecha inválido. Intente nuevamente (YYYY-MM-DD).", 31)

def solicitar_tickers() -> List[str]:
    """
    Solicita al usuario que ingrese símbolos de acciones separados por comas.

    :return: Lista de símbolos de acciones en mayúsculas.
    """
    while True:
        tickers_str = input("Ingrese los símbolos de los tickers separados por comas (ejemplo: MSFT, TSLA, NVDA): ")
        if not tickers_str.strip():
            print_with_color_and_format("Debe ingresar al menos un ticker.", 31)
            continue
        
        tickers = [t.strip().upper() for t in tickers_str.split(",") if t.strip()]
        if tickers:
            return tickers
        else:
            print_with_color_and_format("No se ingresó ningún ticker válido. Intente nuevamente.", 31)

def main() -> None:
    """
    Función principal que orquesta la interacción con el usuario:
      - Solicitar tickers y rango de fechas.
      - Descargar datos y calcular retornos.
      - Mostrar menú con distintas opciones (estadísticas, gráficos, pesos óptimos, informes, etc.).
      - Manejo de excepciones y validaciones.
    """
    print_with_color_and_format("--- Bienvenido al Informe de Análisis de Inversión ---", 32, bold=True)
    print_with_color_and_format("""
Este programa permite realizar un análisis de acciones bursátiles mediante los siguientes conceptos:
    
1) Rendimiento Diario: Cambio porcentual diario en el precio de las acciones, indicando la ganancia o pérdida relativa.
2) Volatilidad: Mide la variabilidad del precio de las acciones a lo largo del tiempo.
3) Peso Óptimo: Distribución óptima de la inversión en las acciones seleccionadas, basada en el rendimiento y la volatilidad.

A continuación, se le pedirá que ingrese los símbolos de las acciones y el rango de fechas para el análisis.
Puede consultar tickers en: 
    """, 32)
    print_with_color_and_format("Yahoo Finance:", 32)
    print_with_color_and_format("https://finance.yahoo.com/lookup/", 34)
    print_with_color_and_format("Nasdaq Symbol Directory:", 32)
    print_with_color_and_format("https://www.nasdaq.com/market-activity/stocks/screener", 34)

    # Solicitar tickers y fechas de inicio y fin
    tickers = solicitar_tickers()
    date_start = solicitar_fecha("Ingrese la fecha de inicio (YYYY-MM-DD): ")
    date_end = solicitar_fecha("Ingrese la fecha de fin (YYYY-MM-DD): ")

    print("\nDescargando datos...")

    try:
        # Descarga de datos
        data = download_data(tickers, date_start, date_end)
        if data is None or data.empty:
            raise ValueError("No se obtuvieron datos para los tickers especificados.")
        
        # Cálculo de retornos logarítmicos
        data = add_performance(data, tickers)
        if data.empty:
            raise ValueError("No fue posible calcular el rendimiento. El DataFrame está vacío.")

    except Exception as e:
        print_with_color_and_format(f"Error al descargar o procesar los datos: {e}", 31)
        return

    # Bucle principal de opciones
    fin = False
    mu = None
    sigma = None
    weights = None

    while not fin:
        print("""\nMenú de opciones:
  1) Ver estadísticas descriptivas.
  2) Generar gráficos.
  3) Calcular peso óptimo.
  4) Generar informe en documento.
  5) Salir
""")
        try:
            opc = int(input("Ingresar opción: "))
        except ValueError:
            print_with_color_and_format("Por favor, ingrese un número válido.", 31)
            continue

        if opc == 1:
            print("Mostrando análisis descriptivo...")
            try:
                stats = descriptive(data)
                if stats.empty:
                    raise ValueError("El DataFrame de estadísticas está vacío.")
                
                # Convertimos el índice en columna para la tabla
                stats.reset_index(inplace=True)
                stats.columns = ["Ticker"] + stats.columns.tolist()[1:]
                print_table(stats)

            except Exception as e:
                print_with_color_and_format(f"Error al mostrar el análisis descriptivo: {e}", 31)

        elif opc == 2:
            print("""\nEscoja el tipo de gráfico que desea visualizar:
  1) Gráfico de series temporales.
  2) Gráfico de cajas.
  3) Matriz de correlación de retornos
  4) Volatilidad movil
  5) Histograma de retornos
""")
            try:
                opc_graf = int(input("Elegir opción: "))
                if opc_graf == 1:
                    show_plot_line_series(data, tickers)
                elif opc_graf == 2:
                    show_plot_box_plot(data, tickers)
                elif opc_graf == 3:
                    from utils.analyse_data import get_correlation_matrix, plot_correlation_heatmap

                    # Elegir las columnas de retornos para correlacionar
                    ret_cols = [f'R{ticker}' for ticker in tickers if f'R{ticker}' in data.columns]
                    if not ret_cols:
                        print("No hay columnas de retornos disponibles para generar una matriz de correlación.")
                    else:
                        corr_matrix = get_correlation_matrix(data, ret_cols)
                        if not corr_matrix.empty:
                            plot_correlation_heatmap(corr_matrix, title="Matriz de Correlación de Retornos")
                            plt.show()

                # Ejemplo de volatilidad móvil
                elif opc_graf == 4:
                    from utils.analyse_data import plot_rolling_volatility
                    ticker_rol = input("Ingrese el ticker para calcular la volatilidad móvil: ").upper()
                    window = int(input("Ingrese la ventana de días (ej. 20): "))
                    plot_rolling_volatility(data, ticker_rol, window=window, title="Volatilidad Móvil")
                    plt.show()

                # Ejemplo de histograma de retornos
                elif opc_graf == 5:
                    from utils.analyse_data import plot_returns_histogram
                    ticker_hist = input("Ingrese el ticker para ver el histograma de retornos: ").upper()
                    bins = int(input("Ingrese el número de bins (ej. 30): "))
                    plot_returns_histogram(data, ticker_hist, bins=bins)
                    plt.show()

                else:
                    print_with_color_and_format("Opción no válida.", 31)
            except ValueError:
                print_with_color_and_format("Por favor, ingrese un número válido.", 31)
            except Exception as e:
                print_with_color_and_format(f"Error al generar el gráfico: {e}", 31)

        elif opc == 3:
            try:
                mu, sigma = calculate_mu_sigma(data, tickers)
                weights = calculate_weight(mu, sigma)  # Retorna dict
                if weights:
                    # Convertir dict a DataFrame para mostrar en tabla
                    weights_df = pd.DataFrame(list(weights.items()), columns=["Ticker", "Peso Óptimo"])
                    print_table(weights_df)
                else:
                    print_with_color_and_format("No se calcularon pesos óptimos.", 31)
            except Exception as e:
                print_with_color_and_format(f"Error al calcular el peso óptimo: {e}", 31)

        elif opc == 4:
            print("""\nOpciones de generación de informe:
  1) Generar informe en un archivo Word.
  2) Generar Excel con los datos obtenidos.
""")
            try:
                opc_informe = int(input("Ingresar opción: "))
            except ValueError:
                print_with_color_and_format("Por favor, ingrese un número válido.", 31)
                continue

            if opc_informe == 1:
                # Generar Word
                try:
                    if mu is None or sigma is None:
                        mu, sigma = calculate_mu_sigma(data, tickers)
                    if weights is None:
                        weights = calculate_weight(mu, sigma)  

                    generar_informe(
                        data, 
                        tickers, 
                        weights, 
                        output="informe.docx", 
                        window_volatility=30  # Por ejemplo, 30 días de ventana
                    )

                except Exception as e:
                    print_with_color_and_format(f"Error al generar el informe: {e}", 31)

            elif opc_informe == 2:
                # Generar Excel
                try:
                    generar_excel(data)
                except Exception as e:
                    print_with_color_and_format(f"Error al generar el Excel: {e}", 31)
            else:
                print_with_color_and_format("Opción no válida. Intente nuevamente.", 31)

        elif opc == 5:
            fin = True
            print("Saliendo del programa... ¡Hasta luego!")
        else:
            print_with_color_and_format("Opción no válida. Intente nuevamente.", 31)


if __name__ == "__main__":
    main()
