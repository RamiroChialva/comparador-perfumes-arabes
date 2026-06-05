Comparador de Perfumes Árabes — Web Scraper

> Herramienta de scraping automático para monitorear y comparar precios de perfumes árabes en e-commerce chileno.

---

¿Qué hace?

Este script extrae en tiempo real el catálogo completo de perfumes árabes de múltiples tiendas online chilenas, normaliza los datos y genera un reporte Excel con:

- Una pestaña por tienda con todos sus productos ordenados por precio
- Una hoja de **Resumen Comparativo** que identifica el precio mínimo disponible para cada perfume y en qué tienda conseguirlo

---

Tiendas cubiertas

| Tienda           | URL                        |
|------------------|----------------------------|
| Elite Perfumes   | eliteperfumes.cl           |
| Mundo Aromas     | mundoaromas.cl             |
| Alisha Perfumes  | alishaperfumes.cl          |

> Las tres operan sobre Shopify, lo que permite usar la API pública `/products.json` con paginación automática.

---

Estructura del output

El archivo `comparador_perfumes_arabes.xlsx` contiene:

**Pestañas por tienda** (ej: `Elite Perfumes`):

| Perfume | Tienda | Precio ($ CLP) | Enlace | Fecha extracción |
|---|---|---|---|---|
| Lattafa Oud For Glory... | Elite Perfumes | 3.825 | https://... | 2026-06-04 12:11 |

**Pestaña Resumen Comparativo**:

| Perfume | Precio ($ CLP) | Tienda más barata | Enlace mejor precio | Precio Elite | Precio Mundo Aromas | Precio Alisha |
|---|---|---|---|---|---|---|
| Armaf Club de Nuit... | 2.751 | Elite Perfumes | https://... | 2.751 | 3.100 | — |

---

Tecnologías utilizadas

- **Python 3.11+**
- `requests` — llamadas HTTP a las APIs Shopify
- `pandas` — transformación y comparación de datos
- `openpyxl` — generación del archivo Excel con múltiples hojas
- `datetime` / `os` — manejo de fechas y rutas del sistema

---

Instalación y uso

### 1. Clonar el repositorio

```bash
git clone https://github.com/tu-usuario/comparador-perfumes-arabes.git
cd comparador-perfumes-arabes
```

### 2. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 3. Ejecutar

```bash
python scraper.py
```

El archivo Excel se genera automáticamente en el **escritorio del usuario**.

---

Estructura del proyecto

```
comparador-perfumes-arabes/
├── scraper.py              # Script principal
├── requirements.txt        # Dependencias
├── README.md
└── output/
    └── comparador_perfumes_arabes.xlsx   # Generado al ejecutar
```

---

Detalles técnicos

- **Paginación automática**: recorre todas las páginas del catálogo hasta recibir una respuesta vacía o un error HTTP
- **Rate limiting**: espera 2 segundos entre requests para respetar el servidor
- **Normalización de nombres**: compara perfumes entre tiendas ignorando mayúsculas y espacios extra, evitando falsos duplicados
- **Manejo de errores**: captura excepciones por página sin interrumpir el scraping completo
- **Deduplicación inteligente**: ante nombres repetidos dentro de una misma tienda, conserva el precio más bajo

---

Posibles extensiones

- [ ] Historial de precios con SQLite para detectar variaciones en el tiempo
- [ ] Alertas automáticas por email o Telegram cuando un producto baja de precio
- [ ] Dashboard interactivo con Streamlit
- [ ] Soporte para más tiendas (no necesariamente Shopify)
- [ ] Exportación a Google Sheets vía API

---

Ramiro Chialva

 
[LinkedIn](https://www.linkedin.com/in/ramiro-chialva-7b2496342/) · [GitHub](https://github.com/RamiroChialva)

---

Licencia

MIT License — libre para usar y modificar.


´´´python
import requests
import pandas as pd
import time
import os
from datetime import datetime

# ──────────────────────────────────────────────
# CONFIGURACIÓN DE TIENDAS
# ──────────────────────────────────────────────
configuracion_tiendas = [
    {
        "nombre": "Elite Perfumes",
        "url_base": "https://eliteperfumes.cl/collections/perfumes-arabes/products.json",
        "url_producto": "https://eliteperfumes.cl/products/"
    },
    {
        "nombre": "Mundo Aromas",
        "url_base": "https://mundoaromas.cl/collections/perfumes-arabes/products.json",
        "url_producto": "https://mundoaromas.cl/products/"
    },
    {
        "nombre": "Alisha Perfumes",
        "url_base": "https://alishaperfumes.cl/collections/arabes/products.json",
        "url_producto": "https://alishaperfumes.cl/products/"
    }
]

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
}

FECHA_EXTRACCION = datetime.now().strftime("%Y-%m-%d %H:%M")

# ──────────────────────────────────────────────
# RUTA DE SALIDA
# ──────────────────────────────────────────────
ruta_escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
if not os.path.exists(ruta_escritorio):
    ruta_escritorio = os.path.join(os.path.expanduser("~"), "Escritorio")

nombre_archivo = os.path.join(ruta_escritorio, "comparador_perfumes_arabes.xlsx")


# ──────────────────────────────────────────────
# FUNCIÓN: SCRAPING DE UNA TIENDA
# ──────────────────────────────────────────────
def scrape_tienda(tienda: dict) -> list[dict]:
    """Extrae todos los productos de una tienda Shopify con paginación automática."""
    productos = []
    pagina = 1

    print(f"\n--- Iniciando scraping: {tienda['nombre']} ---")

    while True:
        url_actual = f"{tienda['url_base']}?page={pagina}"
        print(f"  [{tienda['nombre']}] Página {pagina}...")

        try:
            response = requests.get(url_actual, headers=headers, timeout=15)

            if response.status_code != 200:
                print(f"  Fin de catálogo (HTTP {response.status_code})")
                break

            data = response.json()

            if not data.get("products"):
                print(f"  Catálogo completo ({pagina - 1} páginas procesadas)")
                break

            for producto in data["products"]:
                nombre = producto.get("title", "Sin nombre")
                handle = producto.get("handle", "")
                enlace = f"{tienda['url_producto']}{handle}"
                variants = producto.get("variants", [])
                precio = int(float(variants[0].get("price", 0))) if variants else 0

                productos.append({
                    "Perfume":          nombre,
                    "Tienda":           tienda["nombre"],
                    "Precio ($ CLP)":   precio,
                    "Enlace":           enlace,
                    "Fecha extracción": FECHA_EXTRACCION
                })

            pagina += 1
            time.sleep(2.0)

        except Exception as e:
            print(f"  Error en página {pagina}: {e}")
            break

    return productos


# ──────────────────────────────────────────────
# FUNCIÓN: HOJA RESUMEN COMPARATIVO (CORREGIDA)
# ──────────────────────────────────────────────
def generar_resumen(dfs_por_tienda: dict) -> pd.DataFrame:
    """
    Genera una hoja resumen con el precio mínimo de cada perfume
    entre todas las tiendas, indicando dónde comprarlo más barato.
    """
    todos = pd.concat(dfs_por_tienda.values(), ignore_index=True)

    # Normalizar nombre para comparación (minúsculas, sin espacios extra)
    todos["_nombre_norm"] = todos["Perfume"].str.lower().str.strip()

    # Ordenar por precio para asegurarnos que al borrar duplicados quede el más barato
    todos = todos.sort_values("Precio ($ CLP)", ascending=True)

    # Precio mínimo por perfume único
    resumen = todos.drop_duplicates(subset=["_nombre_norm"], keep="first").copy()

    resumen = resumen.rename(columns={
        "Tienda": "Tienda más barata",
        "Enlace": "Enlace mejor precio"
    })

    resumen = resumen[["Perfume", "Precio ($ CLP)", "Tienda más barata", "Enlace mejor precio", "Fecha extracción"]]
    resumen = resumen.sort_values("Precio ($ CLP)", ascending=True).reset_index(drop=True)

    # Agregar columnas con precio de cada tienda para comparar
    for nombre_tienda, df in dfs_por_tienda.items():
        col_nombre = f"Precio {nombre_tienda}"
        
        # Clonamos el df de la tienda y eliminamos duplicados internos usando el precio más bajo
        df_temp = df.copy()
        df_temp["_name_norm_tienda"] = df_temp["Perfume"].str.lower().str.strip()
        df_temp = df_temp.sort_values("Precio ($ CLP)", ascending=True)
        df_temp = df_temp.drop_duplicates(subset=["_name_norm_tienda"], keep="first")
        
        precios_tienda = df_temp.set_index("_name_norm_tienda")["Precio ($ CLP)"]
        resumen[col_nombre] = resumen["Perfume"].str.lower().str.strip().map(precios_tienda)

    return resumen


# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────
def main():
    todos_los_datos: dict[str, pd.DataFrame] = {}

    try:
        with pd.ExcelWriter(nombre_archivo, engine="openpyxl") as writer:

            # 1. Scraping por tienda → pestaña individual
            for tienda in configuracion_tiendas:
                productos = scrape_tienda(tienda)

                if productos:
                    df = pd.DataFrame(productos)
                    df = df.sort_values("Precio ($ CLP)", ascending=True).reset_index(drop=True)
                    df.to_excel(writer, sheet_name=tienda["nombre"], index=False)
                    todos_los_datos[tienda["nombre"]] = df
                    print(f"  ✔️ {len(productos)} productos guardados en pestaña '{tienda['nombre']}'")
                else:
                    print(f"  ✘ Sin datos para '{tienda['nombre']}'")

            # 2. Hoja resumen comparativo (solo si hay datos de al menos 2 tiendas)
            if len(todos_los_datos) >= 2:
                df_resumen = generar_resumen(todos_los_datos)
                df_resumen.to_excel(writer, sheet_name="⭐ Resumen Comparativo", index=False)
                print(f"\n  ✔️ Resumen comparativo generado ({len(df_resumen)} perfumes únicos)")

        print(f"\n✅ ÉXITO — Archivo guardado en:\n   {nombre_archivo}")

    except Exception as e:
        print(f"\n❌ Error al generar el archivo: {e}")


if _name_ == "_main_":
    main()
