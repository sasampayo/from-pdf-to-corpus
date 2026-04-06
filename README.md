[README.md](https://github.com/user-attachments/files/26512436/README.md)
# From PDF to Corpus
### A Validated Python Tool for Systematic Extraction of Parliamentary Speech Data from the Spanish Congress of Deputies

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![Google Colab](https://img.shields.io/badge/Run%20in-Google%20Colab-orange?logo=google-colab)](https://colab.research.google.com/)
[![DOI](https://zenodo.org/badge/1202730457.svg)](https://doi.org/10.5281/zenodo.19442392)

**Autora:** Sara Sampayo Sande · Centro Crímina, Universidad Miguel Hernández de Elche · ssampayo@crimina.es

---

## ¿Qué hace?

Dos herramientas para construir corpus parlamentarios temáticos a partir de los Diarios de Sesiones del Congreso de los Diputados, sin necesidad de conocimientos avanzados de programación.

| Herramienta | Descripción |
|-------------|-------------|
| `Descargar_PDF.ipynb` | Descarga PDFs del portal del Congreso filtrados por palabra clave, legislatura y cámara |
| `Buscador.ipynb` | Extrae intervenciones individuales de los PDFs y las exporta a Excel con metadatos |

---

## Inicio rápido

**Paso 1** — Descarga los PDFs:  
[![Abrir en Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/drive/1_nJk8t69URI4vhlX9KYhhLnj-V6qT_KJ)

**Paso 2** — Extrae las intervenciones:  
[![Abrir en Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/drive/1hzZebnFW0EuHJSiVMnMd0Y3T9FICM60a)

Cada notebook incluye instrucciones paso a paso. Solo necesitas una cuenta de Google.

> ⚠️ La búsqueda por texto libre en el portal del Congreso solo está disponible desde la **VI Legislatura**.

---

## Rendimiento validado

Evaluado sobre una muestra estratificada de 45 documentos y 183 intervenciones:

| Componente | Métrica | Resultado |
|------------|---------|-----------|
| Detección estructural de títulos | F1 | **96,9%** |
| Filtrado temático | F1 | **87,8%** |
| Identificación de orador | κ | **0,984** |
| Grupo parlamentario | κ | **0,979** |
| Extracción de texto | κ | **0,942** |
| Tipos de intervención | κ | **0,701** |

---

## Estructura del repositorio

```
from-pdf-to-corpus/
├── README.md
├── Descargar_PDF.ipynb
├── Buscador.ipynb
└── diccionarios/          ← grupos parlamentarios por legislatura (1977–2026)
```

---

## Citar

> Sampayo-Sande, S. (2026). *From PDF to Corpus: A Validated Python Tool for Systematic Extraction of Parliamentary Speech Data from the Spanish Congress of Deputies*. Centro Crímina, Universidad Miguel Hernández de Elche. [https://github.com/sasampayo/from-pdf-to-corpus](https://github.com/sasampayo/from-pdf-to-corpus)

---

## Licencia

[MIT](https://opensource.org/licenses/MIT)
