# Guía de construcción del dashboard PideYa (Power BI)

Mira primero los bocetos en `images/boceto_pagina_*.png`. Esta guía explica cómo construir cada uno.

```
Orden de trabajo
1. Cargar datos ─► 2. Modelo ─► 3. Calendario ─► 4. Medidas ─► 5. Tema ─► 6. Páginas ─► 7. Interacción ─► 8. Verificar ─► 9. Guardar
```

---

## 1. Cargar los datos

1. **Inicio → Obtener datos → Libro de Excel** → `data/processed/pideya_limpio.xlsx`
2. Marca solo **Usuarios**, **Restaurantes** y **Pedidos** (NO Log_calidad ni Resumen).
3. Clic en **Transformar datos** y revisa los tipos de columna (ícono a la izquierda del encabezado):

| Tabla | Columna | Tipo |
|---|---|---|
| Usuarios | fecha_registro | Fecha |
| Pedidos | fecha_pedido | Fecha |
| Pedidos | monto_clp | Número entero |
| Pedidos | tiempo_entrega_min, calificacion | Número entero |
| todas | id_* | Número entero |

4. **Cerrar y aplicar**.

## 2. Modelo (vista Modelo, ícono de 3 cajas a la izquierda)

```
 ┌──────────────┐                         ┌──────────────────┐
 │  Usuarios    │ 1                     1 │  Restaurantes    │
 │  id_usuario  ├───────┐       ┌─────────┤  id_restaurante  │
 └──────┬───────┘       │       │         └──────────────────┘
        ┆ (inactiva)    * ▼     ▼ *
        ┆             ┌─────────────────┐
        ┆             │    Pedidos      │
        ┆             │  (tabla hechos) │
        ┆             └────────▲────────┘
        ┆                      │ * 
        ┆                      │ 1 (activa)
        ┆             ┌────────┴────────┐
        └┄┄┄┄┄┄┄┄┄┄┄┄┄┤   Calendario    │
                      └─────────────────┘
```

- Usuarios[id_usuario] 1 → * Pedidos[id_usuario]
- Restaurantes[id_restaurante] 1 → * Pedidos[id_restaurante]
- Dirección de filtro cruzado: **Única** (por defecto). Es un modelo estrella.

## 3. Tabla Calendario

Copia la sección 0 de `medidas_dax.dax` en **Modelado → Nueva tabla**. Después:
- Renombra la columna `Date` a `Fecha`.
- Relación **Calendario[Fecha] → Pedidos[fecha_pedido]** (activa).
- Relación **Calendario[Fecha] → Usuarios[fecha_registro]**: después de crearla, haz doble clic sobre ella y desmarca **"Activar esta relación"** (queda punteada). La medida *Nuevos Usuarios* la activa solo cuando se necesita.
- Clic derecho en Calendario → **Marcar como tabla de fechas** → `Fecha`.
- Selecciona la columna `Mes` → **Ordenar por columna → MesNum**. Igual `Mes Año` → `AñoMes` y `Dia Semana` → `DiaSemanaNum`.
- Usa SIEMPRE los campos de Calendario (no `fecha_pedido` ni `anio_mes`) en ejes y filtros de fecha.

## 4. Medidas

1. **Inicio → Especificar datos** → nombre `_Medidas` → Cargar (queda una tabla vacía arriba de todas).
2. Selecciona `_Medidas` → **Nueva medida** → pega cada medida de `medidas_dax.dax` y aplica el formato indicado.
3. Cuando ya tengas medidas, borra la columna "Columna1" de `_Medidas`: la tabla se convierte en carpeta de medidas.

## 5. Tema

**Vista → Temas → Examinar temas** → `dashboard/tema_pideya.json`. Colores, fondos y bordes quedan aplicados a todo.

Lienzo: **Formato de página → Configuración del lienzo → 16:9**.

---

## 6. Páginas

Estructura común en las 3 páginas principales (ver bocetos):
- **Panel izquierdo oscuro** (forma rectángulo `#1F2937`) con logo, botones de navegación y segmentaciones.
- **Arriba**: título + fila de tarjetas KPI.
- **Centro**: 3 gráficos.

### Página 1 · Resumen ejecutivo — `images/boceto_pagina_1_resumen.png`
Pregunta que responde: *¿cómo va el negocio?*

| Visual | Tipo en Power BI | Campos |
|---|---|---|
| 6 KPI | **Tarjeta** (o Tarjeta nueva con varias medidas) | Ventas Totales · Total Pedidos · Ticket Promedio · Tasa Cancelacion · Tiempo Promedio Entrega · Calificacion Promedio |
| Color de KPIs | fx en color del valor | Color Cancelacion / Color Tiempo / Color Calificacion |
| Ventas y pedidos por mes | **Gráfico de líneas y columnas agrupadas** | Eje X: Calendario[Mes Año] · Columnas: Ventas Totales · Línea: Total Pedidos |
| Ventas por categoría | **Barras agrupadas** (horizontal) | Eje Y: Restaurantes[categoria] · X: Ventas Totales · ordenar desc. |
| Método de pago | **Barras apiladas 100%** | Leyenda: Pedidos[metodo_pago] · Valor: Total Pedidos |
| Subtítulo dinámico | **Tarjeta** sin etiqueta | Titulo Filtros |
| Segmentaciones | **Segmentación** (estilo lista desplegable) | Calendario[Mes Año] · Usuarios[comuna] · Restaurantes[categoria] · Usuarios[tipo_cuenta] |

### Página 2 · Operación y calidad — `images/boceto_pagina_2_operacion.png`
Pregunta: *¿dónde fallamos en el servicio?*

| Visual | Tipo | Campos |
|---|---|---|
| 4 KPI | Tarjetas | Tasa Cancelacion · % Entregas en 30 min · Tiempo Promedio Entrega · % Clientes Satisfechos |
| Cancelación por categoría | Barras agrupadas | Y: categoria · X: Tasa Cancelacion · **Panel Análisis → Línea constante = 0,08** (meta) · color de barra con fx = Color Cancelacion |
| Cancelación por comuna del restaurante | Barras agrupadas | Y: Restaurantes[comuna] · X: Tasa Cancelacion · orden desc. |
| Ventas vs calificación | **Gráfico de dispersión** | Valores (detalle): nombre_restaurante · X: Ventas Totales · Y: Calificacion Promedio · Tamaño: Total Pedidos · Línea constante Y = 4,2 |
| Tabla de restaurantes | **Matriz** | Filas: nombre_restaurante, categoria · Valores: Ventas Totales, Total Pedidos, Tasa Cancelacion, Tiempo Promedio Entrega, Calificacion Promedio · Formato condicional: barras de datos en Ventas, color de fondo en Tasa Cancelacion |

### Página 3 · Clientes — `images/boceto_pagina_3_clientes.png`
Pregunta: *¿quiénes compran, dónde y cuándo?*

| Visual | Tipo | Campos |
|---|---|---|
| 4 KPI | Tarjetas | Clientes Activos · Pedidos por Cliente · Gasto por Cliente · % Ventas Premium |
| Ventas por comuna | Barras agrupadas | Y: Usuarios[comuna] · X: Ventas Totales |
| Día × comuna | **Matriz** | Filas: Usuarios[comuna] · Columnas: Calendario[Dia Semana] · Valores: Total Pedidos · Formato condicional → Color de fondo (escala) |
| Estándar vs Premium | Columnas agrupadas | Eje X: Usuarios[tipo_cuenta] · Valores: Ticket Promedio (o una por gráfico: Ticket, Calificación, Tasa Cancelación) |
| Nuevos usuarios | **Gráfico de áreas** | Eje X: Calendario[Mes Año] · Valores: Nuevos Usuarios |

### Página 4 · Detalle restaurante (OCULTA) — `images/boceto_pagina_4_detalle_restaurante.png`
1. Crea la página y en el panel **Visualizaciones → Obtener detalles** arrastra `Restaurantes[nombre_restaurante]`. Power BI agrega solo el botón "← Volver".
2. Título: tarjeta con la medida **Titulo Restaurante**.
3. KPIs: Ventas Totales, Total Pedidos, Tasa Cancelacion, Tiempo Promedio Entrega, Calificacion Promedio.
4. Columnas: Calendario[Mes Año] vs Ventas Totales. Columnas: Pedidos[calificacion] vs Total Pedidos.
5. Clic derecho en la pestaña de la página → **Ocultar página**.

Para usarla: en la matriz o la dispersión de la página 2, clic derecho sobre un restaurante → **Obtener detalles → Detalle restaurante**.

---

## 7. Interactividad (IE14 = 15 % de la nota)

| Funcionalidad | Cómo |
|---|---|
| **Navegación entre páginas** | Insertar → Botones → En blanco. Formato → Acción → Tipo: *Navegación de páginas* → Destino. Pon el botón de la página actual en naranja `#E4572E`. |
| **Segmentaciones sincronizadas** | Vista → **Sincronizar segmentaciones** → marca Mes en las 3 páginas. Así el filtro de mes se mantiene al cambiar de página. |
| **Botón "Limpiar filtros"** | Insertar → Botones → **Borrar todas las segmentaciones**. |
| **Filtro cruzado** | Por defecto: clic en una barra filtra el resto. Si un visual no debe filtrarse: Formato → **Editar interacciones**. |
| **Drill-through** | Página 4 (arriba). |
| **Información sobre herramientas** | En la dispersión, arrastra Tasa Cancelacion y Tiempo Promedio Entrega a "Información sobre herramientas". |
| **Jerarquía (drill-down)** | Opcional: en el gráfico mensual pon Año > Mes en el eje y activa las flechas de explorar. |

## 8. Verificar los números

Compara tus tarjetas (sin filtros) con `dashboard/valores_control.csv`:

| KPI | Valor esperado |
|---|---|
| Total Pedidos | 1.488 |
| Pedidos Entregados | 1.327 |
| Ventas Totales | $14.809.700 |
| Ticket Promedio | $11.160 |
| Tasa Cancelación | 10,8 % |
| Tiempo Promedio Entrega | 31,5 min |
| % Entregas en 30 min | 45,0 % |
| Calificación Promedio | 4,14 |
| % Clientes Satisfechos | 79,3 % |
| Clientes Activos | 298 |

Si alguno no coincide, casi siempre es un tipo de dato mal puesto o una relación faltante.

## 9. Guardar

**Archivo → Guardar como** → `dashboard/PideYa_dashboard.pbix`. Toma capturas de cada página y guárdalas en `images/` (sirven para el informe y la presentación).

---

## Hallazgos que el dashboard deja ver (base para el storytelling)

1. **La cancelación (10,8 %) supera la meta de 8 %**. Ninguna categoría cumple la meta, pero el problema se concentra en comida **Asiática (13,6 %)** y **Saludable (11,5 %)**, y en restaurantes de **Maipú (17,3 %)** y **Las Condes (13,8 %)**.
2. **Solo el 45 % de las entregas llega en 30 min o menos**; el promedio es 31,5 min.
3. **El tiempo de entrega casi no cambia la calificación** (≈4,1–4,2 en todos los rangos): la satisfacción depende más del restaurante que de la rapidez.
4. **Dulce Tentación** es el 3.er restaurante en ventas pero tiene la **peor calificación (3,77)**: riesgo para la marca.
5. **Saludable** es la categoría que más vende ($5,15 M), y **Santiago** concentra casi un tercio de las ventas por comuna de usuario.
6. Los usuarios Premium (27 % de las ventas) **no** tienen mejor ticket ni mejor calificación que los Estándar, y además **cancelan más (12,4 % vs 10,8 %)**: la cuenta Premium no está fidelizando.
7. Octubre tiene datos solo hasta el día 19, así que no conviene leer su caída como una baja real.

## Limitaciones de los datos (para la evaluación crítica)

- 107 pedidos tienen fecha anterior al registro del usuario (marcados con `pedido_antes_registro = 1`).
- Algunas categorías no calzan con el nombre del restaurante (ej. *Pasta Nostra* y *Trattoria Roma* figuran como "Asiática"): problema de origen que no se corrigió para no inventar datos.
- 3 restaurantes no tienen comuna y 15 usuarios no tienen tipo de cuenta ("Sin información").
- Las metas (8 % de cancelación, 30 min, 4,2★) son supuestos del equipo, no vienen de la empresa.
