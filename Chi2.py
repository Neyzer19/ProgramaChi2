import pandas as pd
import numpy as np
from scipy.stats import chi2
import os

# ==============================================================
# CONFIGURACIÓN Y LECTURA DE DATOS
# ==============================================================

archivo_excel = './Datos.xlsx'

if not os.path.exists(archivo_excel):
    raise FileNotFoundError(f"❌ No se encontró el archivo: {archivo_excel}")

try:
    df = pd.read_excel(archivo_excel)
    print("✅ Archivo leído correctamente.")
except Exception as e:
    raise RuntimeError(f"Error al leer el archivo Excel '{archivo_excel}': {e}")

# ==============================================================
# LIMPIEZA Y DEFINICIÓN DE CATEGORÍAS
# ==============================================================

# Normalizar nombres de columnas
df.columns = [col.strip() for col in df.columns]

# Categorías de interés
categorias = {
    'Terror': ['Terror'],
    'Comedia': ['Comedia'],
    'Drama': ['Drama']
}

# ==============================================================
# CONTEO DE PREFERENCIAS
# ==============================================================

def contar_preferencias(df, grupo, generos):
    """Cuenta cuántos registros hay para un grupo de edad y un género."""
    return df[
        (df['Grupo de edad'] == grupo) &
        (df['Género favorito'].isin(generos))
    ].shape[0]

# Construir DataFrame con resultados
grupos = ['Adultos', 'Adultos mayores', 'Jóvenes']
resultados = [
    {
        'Grupo de edad': grupo,
        'Género': genero,
        'Total': contar_preferencias(df, grupo, generos)
    }
    for grupo in grupos
    for genero, generos in categorias.items()
]

resultados_df = pd.DataFrame(resultados)

# ==============================================================
# TABLA DE FRECUENCIAS OBSERVADAS (O)
# ==============================================================

observados = resultados_df.pivot(
    index='Grupo de edad',
    columns='Género',
    values='Total'
).fillna(0).astype(int)

# Agregar totales
observados_con_totales = observados.copy()
observados_con_totales['Total'] = observados_con_totales.sum(axis=1)
observados_con_totales.loc['Total'] = observados_con_totales.sum(axis=0).astype(int)

print("\n--- TABLA 1: Observados (con totales) ---")
print(observados_con_totales)

# ==============================================================
# TABLA DE FRECUENCIAS ESPERADAS (E)
# ==============================================================

# Calcular totales
totales_filas = observados.sum(axis=1)
totales_columnas = observados.sum(axis=0)
total_general = observados.values.sum()

# Calcular esperados manualmente
esperados = np.outer(totales_filas, totales_columnas) / total_general
esperados_df = pd.DataFrame(esperados, index=observados.index, columns=observados.columns).round(2)

# Agregar totales
esperados_con_totales = esperados_df.copy()
esperados_con_totales['Total'] = esperados_con_totales.sum(axis=1).round(2)
esperados_con_totales.loc['Total'] = esperados_con_totales.sum(axis=0).round(2)

print("\n--- TABLA 2: Esperados (bajo H₀) con totales ---")
print(esperados_con_totales)

# ==============================================================
# TABLA DE CONTRIBUCIONES ( (O - E)² / E )
# ==============================================================

with np.errstate(divide='ignore', invalid='ignore'):
    contrib = ((observados - esperados_df) ** 2) / esperados_df
    contrib = contrib.replace([np.inf, -np.inf], 0).fillna(0).round(4)

# Agregar totales
contrib_con_totales = contrib.copy()
contrib_con_totales['Total'] = contrib_con_totales.sum(axis=1).round(4)
contrib_con_totales.loc['Total'] = contrib_con_totales.sum(axis=0).round(4)

print("\n--- TABLA 3: Contribuciones ( (O - E)² / E ) ---")
print(contrib_con_totales)

# ==============================================================
# CÁLCULO ESTADÍSTICO
# ==============================================================

chi2_stat = contrib.values.sum().round(4)
gl = (observados.shape[0] - 1) * (observados.shape[1] - 1)
p_valor = chi2.sf(chi2_stat, gl)

# ==============================================================
# RESULTADOS Y CONCLUSIÓN
# ==============================================================

print("\n--- RESUMEN DE LA PRUEBA CHI-CUADRADO ---")
print(f"Estadístico χ²: {chi2_stat:.4f}")
print(f"Grados de libertad: {gl}")
print(f"P-valor: {p_valor:.6f}")

alpha = 0.05
if p_valor < alpha:
    print("➡️ RECHAZAMOS la hipótesis nula (H₀): hay asociación significativa.")
else:
    print("➡️ NO RECHAZAMOS la hipótesis nula (H₀): no hay evidencia de asociación.")

# ==============================================================
# EXPORTAR RESULTADOS A EXCEL
# ==============================================================

salida_excel = 'resultados_preferencias.xlsx'

try:
    with pd.ExcelWriter(salida_excel) as writer:
        observados_con_totales.to_excel(writer, sheet_name='Observados')
        esperados_con_totales.to_excel(writer, sheet_name='Esperados')
        contrib_con_totales.to_excel(writer, sheet_name='Contribuciones')

        resumen = pd.DataFrame({
            'Chi-cuadrado': [chi2_stat],
            'Grados_libertad': [gl],
            'P_valor': [p_valor]
        })
        resumen.to_excel(writer, sheet_name='Resumen', index=False)

    print(f"\n📊 Resultados guardados correctamente en '{salida_excel}'.")
except PermissionError:
    print(f"⚠️ No se pudo guardar '{salida_excel}'. Ciérralo si está abierto e inténtalo de nuevo.")
except Exception as e:
    print(f"❌ Error al guardar el archivo Excel: {e}")
