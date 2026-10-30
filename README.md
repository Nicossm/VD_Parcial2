# PideYa · Dashboard de desempeño de delivery

Evaluación Parcial 2 de ADY1104 Visualización de Datos · Duoc UC · sección ADY1104-001D

Equipo: Javier Sagredo, Cristian Romero, Rolando Paredes y Nicolás Osses.

PideYa es una app de delivery de comida que opera en seis comunas de Santiago. Recibimos sus datos de usuarios, restaurantes y pedidos en tres archivos sin limpiar, exportados por sistemas distintos. Los limpiamos con Python, los cargamos en Power BI y construimos un dashboard de tres páginas para la gerencia. El informe ejecutivo y la presentación cuentan lo que encontramos y qué recomendamos hacer.

## Entregables

| Entregable | Archivo |
|---|---|
| Informe Ejecutivo (PDF) | `docs/Informe_PideYa_EP2v2.pdf` |
| Dashboard interactivo (.pbix) | `dashboard/PideYa_dashboard.pbix` |
| Presentación Ejecutiva (PPT) | `docs/PideYa_Presentación_ejecutivav3.pptx` |
| Repositorio del proyecto | Esta carpeta |

## Estructura

```
VD_Parcial2-main/
├── README.md
├── data/
│   ├── raw/                         Datos originales, sin modificar
│   │   ├── pideya_datos.xlsx
│   │   ├── usuarios.csv
│   │   ├── restaurantes.csv
│   │   └── pedidos.csv
│   └── processed/                   Datos limpios (los genera src/limpieza_pideya.py)
│       ├── pideya_limpio.xlsx       Fuente del dashboard
│       ├── *_limpio.csv
│       ├── log_calidad_datos.csv    Bitácora: cada corrección y cuántas filas tocó
│       └── log_calidad_datos.xlsx   Bitácora y resumen de filas, en Excel
├── src/
│   └── limpieza_pideya.py           Limpieza, validación y exportación
├── notebooks/
│   └── analisis_pideya.ipynb        Retención, captación, comunas, pruebas de diferencia y figuras del informe
├── dashboard/
│   ├── PideYa_dashboard.pbix        Dashboard en Power BI
│   ├── medidas_dax.dax              Medidas DAX y tabla Calendario
│   ├── modelo_pideya.tmdl           Script TMDL: tabla Calendario, medidas y relaciones del Calendario
│   ├── tema_pideya.json             Tema de colores
│   ├── guia_dashboard.md            Cómo se construyó cada página
│   └── valores_control.csv          Indicadores calculados en Python para validar el dashboard
├── images/
│   ├── dashboard_*.png              Capturas de las tres páginas del dashboard
│   ├── boceto_pagina_*.png          Bocetos previos al dashboard
│   ├── eda/                         Gráficos del primer análisis exploratorio
│   └── informe/                     Figuras del informe (las genera el notebook)
└── docs/
    ├── Informe_PideYa_EP2v2.pdf               Informe ejecutivo (entregable)
    ├── Informe_PideYa_EP2v2.docx              Mismo informe en Word, editable
    ├── PideYa_Presentación_ejecutivav3.pptx   Presentación para la exposición
    └── src/
        └── build_informe_docx.js              Script que genera el informe en Word
```

## Cómo reproducir

Desde la carpeta raíz del proyecto:

```bash
pip install pandas openpyxl matplotlib jupyter
python src/limpieza_pideya.py
jupyter nbconvert --to notebook --execute --inplace notebooks/analisis_pideya.ipynb
```

El primer comando deja los datos limpios en `data/processed/`. El segundo recalcula las cifras del informe y regenera las figuras en `images/informe/`.

Después abre `dashboard/PideYa_dashboard.pbix` en Power BI Desktop y usa **Inicio → Actualizar**. Si la ruta del Excel cambió: **Transformar datos → Configuración de origen de datos → Cambiar origen**.

Para regenerar el informe en Word (requiere Node.js):

```bash
npm install docx
node docs/src/build_informe_docx.js
```

El PDF del informe se obtiene abriendo el `.docx` en Word: **Archivo → Exportar → Crear PDF**.

## Datos

| Tabla | Filas originales | Filas limpias | Clave |
|---|---|---|---|
| Usuarios | 308 | 300 | id_usuario |
| Restaurantes | 30 | 30 | id_restaurante |
| Pedidos | 1.515 | 1.488 | id_pedido (se relaciona por id_usuario e id_restaurante) |

Los pedidos van del 1 de mayo al 19 de octubre de 2026. Octubre está incompleto.

El script corrige fechas escritas en tres formatos, nombres con mayúsculas, tildes o espacios distintos (incluido el error "Tajeta"), montos con texto ("$11.200", "12600 CLP"), filas duplicadas, pedidos de usuarios o restaurantes que no existen y tiempos de entrega imposibles (-5, 0, 500 y 999 minutos). Los tiempos imposibles quedan en blanco en vez de reemplazarse por un promedio. Los 107 pedidos con fecha anterior al registro del usuario se conservan y quedan marcados en la columna `pedido_antes_registro`. El detalle está en `data/processed/log_calidad_datos.csv`.

Las categorías de restaurante no siempre calzan con el nombre del local (por ejemplo, Pasta Nostra figura como Asiática). No las cambiamos porque no hay forma de saber la categoría correcta.

## Indicadores

Ventas totales, total de pedidos, ticket promedio, tasa de cancelación, tiempo promedio de entrega, entregas en 30 minutos, calificación promedio, clientes satisfechos (notas 4 y 5), clientes activos, pedidos y gasto por cliente, ventas Premium y nuevos usuarios.

Ventas, ticket, tiempo y calificación se calculan solo con pedidos entregados, porque un pedido cancelado no genera ingreso ni recibe nota.

Las metas las propuso el equipo, ya que PideYa no entregó objetivos formales: cancelación de 8% o menos, entrega en 30 minutos o menos y calificación de 4,2 o más.

## Herramientas

Python 3 (pandas, matplotlib, Jupyter) · Power BI Desktop · Node.js (librería docx, para generar el informe)
