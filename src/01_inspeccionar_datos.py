# PASO 5 - INSPECCIONAR EL DATASET

# %%

import pandas as pd
from pathlib import Path

# Ubicación del archivo Excel
archivo = Path("data/raw/Coffee Shop Sales.xlsx")

# Identificar las hojas disponibles
excel = pd.ExcelFile(archivo)

print("Hojas disponibles:")
print(excel.sheet_names)

# Cargar la primera hoja
df = pd.read_excel(archivo, sheet_name=excel.sheet_names[0])

print("\nDimensiones del dataset:")
print(df.shape)

print("\nPrimeros 5 registros:")
print(df.head())

print("\nColumnas:")
print(df.columns.tolist())


# %%

# PASO 6 - VISUALIZACIÓN DE DATOS

# Seleccionar los primeros 1000 registros
df_vista = df.head(1000)

# Mostrar una vista previa
print(df_vista.head(10).to_string())


# %%
# PASO 7 - DIAGNÓSTICO DE CALIDAD DEL DATASET

print("\n========== INFORMACIÓN GENERAL ==========")
df.info()

print("\n========== VALORES FALTANTES ==========")
print(df.isnull().sum())

print("\n========== FILAS DUPLICADAS ==========")
print(df.duplicated().sum())

print("\n========== ESTADÍSTICAS DESCRIPTIVAS ==========")
print(df.describe().to_string())


# %%
# PASO 8 - IDENTIFICAR VALORES INCONSISTENTES

print("\n========== VALIDACIÓN DE CANTIDADES ==========")

print("Cantidad mínima:", df["transaction_qty"].min())
print("Cantidad máxima:", df["transaction_qty"].max())

# Identificar cantidades inválidas
cantidades_invalidas = df[df["transaction_qty"] <= 0]

print("Registros con cantidades inválidas:",
      len(cantidades_invalidas))


print("\n========== VALIDACIÓN DE PRECIOS ==========")

print("Precio mínimo:", df["unit_price"].min())
print("Precio máximo:", df["unit_price"].max())

# Identificar precios inválidos
precios_invalidos = df[df["unit_price"] <= 0]

print("Registros con precios inválidos:",
      len(precios_invalidos))


print("\n========== RESUMEN DE CALIDAD ==========")

print("Total de registros:", len(df))
print("Valores faltantes:", df.isnull().sum().sum())
print("Filas duplicadas:", df.duplicated().sum())
print("Cantidades inválidas:", len(cantidades_invalidas))
print("Precios inválidos:", len(precios_invalidos))


# %%
# PASO 9 - REVISAR CATEGORÍAS Y PERIODO DE LOS DATOS

print("\n========== PERIODO DEL DATASET ==========")

print("Desde:", df["transaction_date"].min())
print("Hasta:", df["transaction_date"].max())


print("\n========== LOCALES ==========")

print(df["store_location"].value_counts().to_string())


print("\n========== CATEGORÍAS DE PRODUCTOS ==========")

print(df["product_category"].value_counts().to_string())


print("\n========== TIPOS DE PRODUCTOS ==========")

print("Cantidad de tipos diferentes:",
      df["product_type"].nunique())

print(df["product_type"].value_counts().to_string())


print("\n========== TABLA RESUMEN DE TIPOS ==========")

tipos_producto = (
    df["product_type"]
    .value_counts()
    .rename_axis("Tipo de producto")
    .reset_index(name="Registros")
)

print(tipos_producto.to_string(index=False))



# %%
# PASO 10 - CALCULAR LAS VENTAS POR REGISTRO

print("\n========== CALCULO DE VENTAS ==========")

# Crear una copia para no modificar el dataset original
df_analysis = df.copy()

# Calcular el importe de cada registro
df_analysis["total_sales"] = (
    df_analysis["transaction_qty"] *
    df_analysis["unit_price"]
)

# Mostrar los primeros 10 registros
print("\nPrimeras 10 ventas calculadas:")

print(
    df_analysis[
        [
            "transaction_id",
            "transaction_qty",
            "unit_price",
            "total_sales"
        ]
    ].head(10).to_string(index=False)
)

# Calcular la facturación total
print("\n========== FACTURACION TOTAL ==========")

print("Ventas totales:", round(df_analysis["total_sales"].sum(), 2))

# Comprobar la nueva columna
print("\nCantidad de columnas:", df_analysis.shape[1])



# %%
# PASO 11 - EXTRAER INFORMACION DE FECHAS Y HORARIOS

print("\n========== ANALISIS DE FECHAS Y HORARIOS ==========")

# Extraer el día de la semana
df_analysis["day_of_week"] = (
    df_analysis["transaction_date"].dt.day_name()
)

# Extraer el número del mes
df_analysis["month"] = (
    df_analysis["transaction_date"].dt.month
)

# Extraer el nombre del mes
df_analysis["month_name"] = (
    df_analysis["transaction_date"].dt.month_name()
)

# Extraer la hora de cada transacción
df_analysis["hour"] = (
    pd.to_datetime(
        df_analysis["transaction_time"].astype(str),
        format="%H:%M:%S"
    ).dt.hour
)

# Mostrar las nuevas variables
print("\nPrimeros 10 registros:")

print(
    df_analysis[
        [
            "transaction_date",
            "day_of_week",
            "month",
            "month_name",
            "transaction_time",
            "hour"
        ]
    ].head(10).to_string(index=False)
)

# Comprobar la cantidad de columnas
print("\nTotal de columnas:", df_analysis.shape[1])


# %%
# PASO 12 - IDENTIFICAR LOS PRODUCTOS

print("\n========== IDENTIFICACION DE PRODUCTOS ==========")

resumen_productos = pd.DataFrame({
    "Campo": [
        "product_id",
        "product_detail",
        "product_type",
        "product_category"
    ],

    "Valores únicos": [
        df_analysis["product_id"].nunique(),
        df_analysis["product_detail"].nunique(),
        df_analysis["product_type"].nunique(),
        df_analysis["product_category"].nunique()
    ]
})

print("\nResumen de productos:")
print(resumen_productos.to_string(index=False))



# %%
# PASO 13 - CONSTRUIR TABLA RESUMEN POR PRODUCTO

import time

print("\n========== RESUMEN POR PRODUCTO ==========")

# Medir el tiempo de procesamiento
inicio = time.perf_counter()

# Agrupar las transacciones por producto
df_productos = (
    df_analysis
    .groupby("product_id", as_index=False)
    .agg(
        product_detail=("product_detail", "first"),
        product_type=("product_type", "first"),
        product_category=("product_category", "first"),
        unidades_vendidas=("transaction_qty", "sum"),
        facturacion=("total_sales", "sum"),
        registros=("transaction_id", "count"),
        precio_promedio=("unit_price", "mean")
    )
)

# Mostrar dimensiones y tiempo
print("Dimensiones:", df_productos.shape)

print(
    "Tiempo de procesamiento:",
    round(time.perf_counter() - inicio, 2),
    "segundos"
)

# Mostrar los primeros 10 productos
print("\nPrimeros 10 productos:")

print(df_productos.head(10).to_string(index=False))

# Comprobar que las ventas coincidan
print("\n========== VALIDACION ==========")

print("Productos analizados:", len(df_productos))

print(
    "Facturacion agrupada:",
    round(df_productos["facturacion"].sum(), 2)
)



# %%
# PASO 14 - EXAMINAR VARIABLES PARA K-MEANS

print("\n========== ESTRUCTURA DE PRODUCTOS ==========")

# Comprobar dimensiones
print("Dimensiones:", df_productos.shape)

# Comprobar productos únicos
print(
    "Productos únicos:",
    df_productos["product_id"].nunique()
)

# Comprobar valores faltantes
print("\n========== VALORES FALTANTES ==========")

print(df_productos.isnull().sum().to_string())

# Estadísticas de las variables numéricas
print("\n========== ESTADISTICAS DESCRIPTIVAS ==========")

estadisticas = df_productos[
    [
        "unidades_vendidas",
        "facturacion",
        "registros",
        "precio_promedio"
    ]
].describe().T.round(2)

print(estadisticas.to_string())



# %%
# PASO 15 - ANALIZAR CORRELACION ENTRE VARIABLES

print("\n========== ANALISIS DE CORRELACION ==========")

# Seleccionar las variables numéricas
variables_kmeans = df_productos[
    [
        "unidades_vendidas",
        "facturacion",
        "registros",
        "precio_promedio"
    ]
]

# Calcular la matriz de correlación de Pearson
correlacion = variables_kmeans.corr()

# Mostrar los resultados
print("\nMatriz de correlación:")
print(correlacion.round(2).to_string())

# Analizar específicamente demanda y precio
print("\n========== RELACION ENTRE DEMANDA Y PRECIO ==========")

relacion = correlacion.loc[
    "unidades_vendidas",
    "precio_promedio"
]

print("Correlación:", round(relacion, 2))



# %%
# PASO 15.1 - GRAFICO DE DISPERSION

import matplotlib.pyplot as plt
from pathlib import Path

print("\n========== GRAFICO DE DEMANDA Y PRECIO ==========")

# Definir la carpeta de salida
carpeta = Path("outputs/figures")

# Crear la carpeta si no existe
carpeta.mkdir(parents=True, exist_ok=True)

# Crear el gráfico
fig, ax = plt.subplots(figsize=(8, 5))

ax.scatter(
    df_productos["unidades_vendidas"],
    df_productos["precio_promedio"]
)

# Configurar títulos y ejes
ax.set_xlabel("Unidades vendidas")
ax.set_ylabel("Precio promedio")
ax.set_title("Productos según demanda y precio")

# Ajustar el gráfico
fig.tight_layout()

# Guardar como imagen PNG
ruta = carpeta / "prueba_productos.png"

fig.savefig(
    ruta,
    dpi=150,
    bbox_inches="tight"
)

plt.close(fig)

print("Gráfico generado correctamente")
print("Ubicación:", ruta.resolve())



# %%
# PASO 16 - ESTANDARIZAR VARIABLES PARA K-MEANS

from sklearn.preprocessing import StandardScaler

print("\n========== ESTANDARIZACION DE VARIABLES ==========")

# Seleccionar las dos variables para K-means
variables_kmeans = df_productos[
    [
        "unidades_vendidas",
        "precio_promedio"
    ]
]

# Crear el estandarizador
scaler = StandardScaler()

# Estandarizar las variables
X_scaled = scaler.fit_transform(variables_kmeans)

# Mostrar los primeros 5 registros
print("\nPrimeros 5 registros estandarizados:")

print(X_scaled[:5])

# Comprobar las dimensiones
print("\nDimensiones:", X_scaled.shape)

# Comprobar media y desviación estándar
print("\nMedia:", X_scaled.mean(axis=0).round(2))

print("Desviación estándar:", X_scaled.std(axis=0).round(2))
