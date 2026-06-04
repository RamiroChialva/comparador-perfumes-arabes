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
