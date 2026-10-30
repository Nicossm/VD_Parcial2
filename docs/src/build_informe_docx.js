// Genera docs/Informe_PideYa_EP2v2.docx
// Requiere: npm install docx   |   Uso: node docs/src/build_informe_docx.js
// Las figuras salen de notebooks/analisis_pideya.ipynb (images/informe/) y de las capturas del dashboard.

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, Footer,
  AlignmentType, WidthType, ShadingType, BorderStyle, LevelFormat, PageNumber,
  TabStopType,
  VerticalAlign, HeadingLevel,
} = require("docx");

const IMG = path.resolve(__dirname, "../../images");
const OUT = path.resolve(__dirname, "../Informe_PideYa_EP2v2.docx");

const OSCURO = "1F2937", NARANJO = "E4572E", ROJO = "C2410C", VERDE = "15803D";
const GRIS = "6B7280", BORDE = "D1D5DB", CLARO = "F3F4F6";
const ANCHO = 9360; // ancho útil en DXA (carta con márgenes de 1")

// ---------- utilidades ----------
// Texto con **negrita** opcional
function runs(texto, base = {}) {
  return texto.split("**").map((t, i) => new TextRun({ text: t, bold: i % 2 === 1 || base.bold, ...base }));
}
const p = (texto, opts = {}) => new Paragraph({ children: runs(texto, opts.run), spacing: { after: 100, line: 259 }, ...opts.par });
const h1 = (texto) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(texto)] });
const h2 = (texto) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(texto)] });
const vineta = (texto) => new Paragraph({ numbering: { reference: "vinetas", level: 0 }, children: runs(texto), spacing: { after: 60, line: 264 } });
const leyenda = (texto) => new Paragraph({ children: [new TextRun({ text: texto, size: 16, color: GRIS })], spacing: { before: 40, after: 160 } });

function imagen(archivo, anchoPx, altoPx) {
  return new ImageRun({ type: "png", data: fs.readFileSync(path.join(IMG, archivo)), transformation: { width: anchoPx, height: altoPx } });
}
function figura(archivo, w0, h0, anchoPx, texto) {
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80 }, keepNext: true, children: [imagen(archivo, anchoPx, Math.round(anchoPx * h0 / w0))] }),
    leyenda(texto),
  ];
}

const bordeCelda = { style: BorderStyle.SINGLE, size: 4, color: BORDE };
const bordes = { top: bordeCelda, bottom: bordeCelda, left: bordeCelda, right: bordeCelda };

function celda(contenido, ancho, { encabezado = false, fondo, alinear, color, negrita, junto = false } = {}) {
  const parrafos = (Array.isArray(contenido) ? contenido : [contenido]).map((t) => new Paragraph({
    alignment: alinear || AlignmentType.LEFT,
    spacing: { after: 0, line: 240 },
    keepNext: junto,
    children: runs(String(t), { size: 17, color: encabezado ? "FFFFFF" : (color || OSCURO), bold: encabezado || negrita }),
  }));
  return new TableCell({
    width: { size: ancho, type: WidthType.DXA },
    borders: bordes,
    verticalAlign: VerticalAlign.CENTER,
    shading: encabezado ? { fill: OSCURO, type: ShadingType.CLEAR, color: "auto" } : (fondo ? { fill: fondo, type: ShadingType.CLEAR, color: "auto" } : undefined),
    margins: { top: 45, bottom: 45, left: 90, right: 90 },
    children: parrafos,
  });
}

// filas: arreglo de arreglos; num: índices de columnas numéricas (alineadas a la derecha)
function tabla(encabezados, filas, anchos, { num = [], resaltar = {}, junto = false } = {}) {
  const alin = (i) => (num.includes(i) ? AlignmentType.RIGHT : AlignmentType.LEFT);
  return new Table({
    width: { size: anchos.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: anchos,
    rows: [
      new TableRow({ tableHeader: true, children: encabezados.map((e, i) => celda(e, anchos[i], { encabezado: true, alinear: alin(i), junto })) }),
      ...filas.map((f, r) => new TableRow({
        cantSplit: true,
        children: f.map((v, i) => celda(v, anchos[i], { alinear: alin(i), fondo: r % 2 ? CLARO : undefined, color: (resaltar[`${r},${i}`]), junto: junto && r < filas.length - 1 })),
      })),
    ],
  });
}
const espacio = (after = 120) => new Paragraph({ spacing: { after }, children: [] });

// ---------- portada ----------
const portada = [
  new Paragraph({ spacing: { before: 1800, after: 600 }, children: [new TextRun({ text: "PideYa", bold: true, size: 40, color: NARANJO })] }),
  new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "INFORME EJECUTIVO · EVALUACIÓN PARCIAL 2", size: 20, color: GRIS, characterSpacing: 20 })] }),
  new Paragraph({ spacing: { after: 300, line: 240 }, children: [new TextRun({ text: "PideYa retiene a sus clientes, pero cada vez suma menos clientes nuevos", bold: true, size: 52, color: OSCURO })] }),
  new Paragraph({ spacing: { after: 2400, line: 300 }, children: [new TextRun({ text: "Hallazgos y recomendaciones a partir del dashboard de ventas, servicio y clientes. Pedidos del 1 de mayo al 19 de octubre de 2026.", size: 24, color: "4B5563" })] }),
  new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Para: gerencia de PideYa", size: 21, color: OSCURO, bold: true })] }),
  new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Equipo consultor: Javier Sagredo, Cristian Romero, Rolando Paredes y Nicolás Osses", size: 21, color: OSCURO })] }),
  new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "ADY1104 Visualización de Datos · Sección ADY1104-001D · Duoc UC", size: 21, color: GRIS })] }),
  new Paragraph({ children: [new TextRun({ text: "Octubre de 2026", size: 21, color: GRIS })] }),
];

// ---------- 1. Contexto y problema ----------
const kpi = (valor, etiqueta, nota, colorValor = OSCURO) => new TableCell({
  width: { size: 1560, type: WidthType.DXA },
  borders: { top: { style: BorderStyle.SINGLE, size: 12, color: "FFFFFF" }, bottom: { style: BorderStyle.SINGLE, size: 12, color: "FFFFFF" }, left: { style: BorderStyle.SINGLE, size: 12, color: "FFFFFF" }, right: { style: BorderStyle.SINGLE, size: 12, color: "FFFFFF" } },
  shading: { fill: CLARO, type: ShadingType.CLEAR, color: "auto" },
  margins: { top: 100, bottom: 100, left: 120, right: 80 },
  children: [
    new Paragraph({ spacing: { after: 20 }, children: [new TextRun({ text: valor, bold: true, size: 30, color: colorValor })] }),
    new Paragraph({ spacing: { after: 10 }, children: [new TextRun({ text: etiqueta, size: 16, color: OSCURO })] }),
    new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: nota, size: 15, color: GRIS })] }),
  ],
});
const filaKpi = new Table({
  width: { size: ANCHO, type: WidthType.DXA },
  columnWidths: Array(6).fill(1560),
  rows: [new TableRow({ children: [
    kpi("$14,8 M", "Ventas", "1.327 pedidos entregados"),
    kpi("1.488", "Pedidos", "may a oct 2026"),
    kpi("10,8%", "Cancelación", "meta 8%", ROJO),
    kpi("45,0%", "Entregas en 30 min", "promedio 31,5 min; meta 30", ROJO),
    kpi("4,14", "Calificación", "meta 4,2", ROJO),
    kpi("97,7%", "Clientes que repiten", "298 clientes activos"),
  ] })],
});

const seccion1 = [
  h1("1. Contexto y problema"),
  p("PideYa es una app de delivery de comida que opera en seis comunas de Santiago: Santiago, Providencia, Las Condes, Ñuñoa, Maipú y La Florida. Conecta a 300 usuarios registrados con 30 restaurantes de cinco categorías. Entre el 1 de mayo y el 19 de octubre de 2026 recibió 1.488 pedidos; se entregaron 1.327 y sumaron $14,8 millones en ventas."),
  p("La empresa creció sin consolidar sus sistemas: usuarios, restaurantes y pedidos estaban en tres bases separadas, con formatos propios, y la gerencia no tenía una cifra única de ventas, cancelaciones o satisfacción. Nos pidió responder cuatro preguntas para decidir en qué comunas invertir y qué alianzas con restaurantes fortalecer. Estas son nuestras respuestas:"),
  tabla(
    ["Pregunta de la gerencia", "Respuesta corta", "Detalle"],
    [
      ["¿Qué restaurantes y categorías generan más pedidos y mejores calificaciones?", "Saludable lidera en ventas (35%) y en nota (4,22). Panadería del Sol, Café Central y La Baguette reciben más pedidos. Dulce Tentación vende mucho con la peor nota.", "3.4"],
      ["¿En qué comunas hay mejor oportunidad de crecimiento?", "La Florida y Ñuñoa. Maipú suma usuarios, pero sus restaurantes cancelan demasiado. Las Condes casi no suma usuarios.", "3.5"],
      ["¿Qué tan bien retiene PideYa a sus usuarios?", "Muy bien: 97,7% de los clientes volvió a pedir. El problema es que cada vez entran menos clientes nuevos.", "3.1"],
      ["¿Cómo es la experiencia de entrega: tiempos, cancelaciones, medios de pago?", "Se cancela el 10,8% de los pedidos y una cuarta parte de esas cancelaciones viene de cinco restaurantes. El 55% llega después de 30 minutos, pero eso no baja la nota. El medio de pago no cambia nada.", "3.2 y 3.3"],
    ],
    [3100, 5260, 1000],
  ),
  espacio(100),
  filaKpi,
  leyenda("En rojo, los indicadores que no cumplen la meta. Las metas las propuso el equipo como referencia, porque PideYa no tenía objetivos formales; conviene validarlas con la gerencia."),
];

// ---------- 2. Tratamiento de datos ----------
const seccion2 = [
  h1("2. Tratamiento de datos"),
  p("Revisamos las tres tablas antes de construir el dashboard. Un script en Python (src/limpieza_pideya.py) corrige los problemas, valida el resultado y anota cada corrección en una bitácora (data/processed/log_calidad_datos.csv). Si llegan datos nuevos, basta con volver a ejecutarlo."),
  tabla(
    ["Problema encontrado", "Filas", "Qué hicimos"],
    [
      ["Fechas escritas en tres formatos (aaaa-mm-dd, dd/mm/aaaa y dd-mm-aaaa)", "743", "Las convertimos a un solo formato de fecha."],
      ["Comunas, categorías, métodos de pago y estados escritos de distintas formas (ÑUÑOA, Nunoa, Tajeta, espacios de más)", "313", "Los llevamos a una lista oficial de valores."],
      ["Montos con texto ($11.200, 12600 CLP)", "149", "Los convertimos a número entero en pesos."],
      ["Filas duplicadas", "23", "Las eliminamos (8 en usuarios y 15 en pedidos)."],
      ["Pedidos de usuarios o restaurantes que no existen (id 9999 y 999)", "12", "Los eliminamos: no se pueden relacionar con ninguna tabla."],
      ["Tiempos de entrega imposibles (-5, 0, 500 y 999 minutos)", "12", "Los dejamos en blanco; reemplazarlos por un promedio sería inventar datos. Otros 59 ya venían sin tiempo."],
      ["Comuna de restaurante o tipo de cuenta vacíos", "18", "Los marcamos como Sin información para que sigan visibles."],
      ["Pedidos con fecha anterior al registro del usuario", "107", "Los conservamos y los marcamos en una columna de control."],
    ],
    [4300, 700, 4360],
    { num: [1] },
  ),
  espacio(80),
  p("Quedaron 300 usuarios (de 308), 30 restaurantes y 1.488 pedidos (de 1.515). En Power BI, los pedidos se conectan con usuarios y restaurantes por sus identificadores, y un calendario permite filtrar por mes. Comparamos cada indicador del dashboard con un cálculo independiente en Python (dashboard/valores_control.csv) y todos coinciden. Ventas, ticket, tiempo y calificación se calculan solo con pedidos entregados, porque un pedido cancelado no genera ingreso ni recibe nota."),
  p("Hay un problema que no corregimos. Varias categorías no calzan con el nombre del local: Pasta Nostra figura como Asiática, Dulce Tentación como Italiana y Wok Express como Saludable. No tenemos cómo saber la categoría correcta, así que la dejamos como viene y nuestras recomendaciones apuntan a restaurantes, no a categorías."),
];

// ---------- 3. Hallazgos ----------
const seccion3 = [
  h1("3. Hallazgos"),

  h2("3.1 Los clientes vuelven, pero cada vez entran menos clientes nuevos"),
  p("De los 300 usuarios registrados, 298 pidieron al menos una vez en el período (los clientes activos del dashboard) y 291 (97,7%) volvieron a pedir. La mediana es de 5 pedidos por cliente, con 19 días entre un pedido y el siguiente. Cada mes vuelve a comprar entre el 54% y el 59% de los clientes del mes anterior. Una cancelación tampoco los aleja: quienes tuvieron cancelado su primer pedido hicieron en promedio 4,94 pedidos, casi lo mismo que el resto (5,00)."),
  p("La captación, en cambio, va a la baja. Entre enero y mayo de 2026 se registraban 13,8 usuarios nuevos por mes; entre junio y septiembre, 8,8, y en octubre no hay registros. Como casi toda la base ya compra, las ventas se mantienen planas, entre $2,5 y $2,8 millones al mes: para crecer, PideYa necesita sumar clientes. Además, 73 clientes (24,5%) no piden hace más de 45 días."),
  ...figura("informe/fig_retencion.png", 2400, 800, 620, "Figura 1. Clientes según la cantidad de pedidos (izquierda) y usuarios nuevos por mes en 2026 (derecha). Fuente: notebooks/analisis_pideya.ipynb; la serie completa de registros está en la página Clientes del dashboard (anexo B)."),
  p("La cuenta Premium no cambia este comportamiento. Los clientes Premium piden con la misma frecuencia que los Estándar (4,94 contra 4,98 pedidos) y gastan lo mismo por pedido ($11.081 contra $11.173). Cancelan algo más (12,4% contra 10,8%), pero la diferencia es tan pequeña que puede deberse al azar (p = 0,39)."),

  h2("3.2 Cinco restaurantes explican una cuarta parte de las cancelaciones"),
  p("La tasa de cancelación es 10,8%, sobre la meta de 8%, y no está repartida de forma pareja. Cinco de los 30 restaurantes reúnen 42 de las 161 cancelaciones (26%) con solo el 15,5% de los pedidos. Si esos cinco cancelaran al 8%, la tasa total bajaría a cerca de 9,2%."),
  tabla(
    ["Restaurante", "Comuna", "Pedidos", "Cancelación"],
    [
      ["Sushi Mar", "Las Condes", "44", "22,7%"],
      ["Trattoria Roma", "Maipú", "43", "18,6%"],
      ["Mar y Tierra", "Ñuñoa", "35", "17,1%"],
      ["La Esquina Criolla", "Santiago", "48", "16,7%"],
      ["Ramen Ya", "Maipú", "61", "16,4%"],
    ],
    [3400, 2400, 1600, 1960],
    { num: [2, 3], junto: true },
  ),
  espacio(60),
  p("Esta lista explica lo que muestra el dashboard por comuna y por categoría (figura 2). Maipú cancela 17,3% porque sus dos únicos restaurantes, Trattoria Roma y Ramen Ya, están en la lista; Las Condes llega a 13,8% sobre todo por Sushi Mar. La categoría Asiática (13,6%) incluye a tres de los cinco locales, pero como las categorías vienen mal asignadas, conviene actuar por restaurante."),
  ...figura("informe/fig_cancelacion.png", 781, 391, 480, "Figura 2. Tasa de cancelación por categoría y por comuna del restaurante. La línea punteada marca la meta de 8%. Fuente: dashboard, página Operación."),
  p("La meta se puede alcanzar: en agosto la cancelación bajó a 7,0%, mientras que junio (14,3%) y julio (13,6%) fueron los peores meses. Los datos no registran el motivo de cancelación, así que no sabemos qué cambió en agosto. Vale la pena preguntarlo al equipo de operaciones."),

  h2("3.3 Más de la mitad de las entregas llega tarde, pero eso no baja la nota"),
  p("El tiempo promedio de entrega es 31,5 minutos y la mediana, 32. Solo el 45,0% de los pedidos llega en 30 minutos o menos. Aun así, la calificación casi no cambia con el tiempo: va de 4,12 a 4,22 en todos los rangos, y la correlación entre ambas variables es prácticamente cero. Acelerar el reparto, por sí solo, no subiría la satisfacción. La nota depende más del restaurante, y la tabla de restaurantes del dashboard (página Operación) lo muestra: Tacos y Punto tiene el reparto más lento (33,8 min) y la segunda mejor nota (4,41), mientras que Dulce Tentación entrega antes que el promedio (30,9 min) y tiene la peor (3,77)."),
  ...figura("informe/fig_tiempo_calificacion.png", 1600, 640, 450, "Figura 3. Calificación promedio según los minutos de entrega; n es la cantidad de pedidos calificados. Fuente: notebooks/analisis_pideya.ipynb."),
  p("El medio de pago tampoco marca diferencias. La tarjeta reúne el 58% de los pedidos, el efectivo el 26% y la transferencia el 16%, y los tres tienen tasas de cancelación parecidas (9,4% a 11,1%) y casi el mismo ticket (cerca de $11.100)."),

  h2("3.4 Saludable lidera; Dulce Tentación vende mucho con la peor nota"),
  p("Saludable es la categoría que más vende (35% de las ventas, con 11 restaurantes) y la mejor evaluada: 4,22 de nota y 82,6% de clientes satisfechos. Le siguen Asiática (27%) e Italiana (22%). Postres y café tiene un solo local y aporta menos del 2%."),
  p("Los restaurantes con más pedidos son Panadería del Sol (68), Café Central (63) y La Baguette (62). Los mejor evaluados son Mar y Tierra (4,44), Tacos y Punto (4,41) y Ramen Ya (4,28). Mar y Tierra tiene la mejor nota, pero cancela uno de cada seis pedidos."),
  p("El caso que más preocupa es Dulce Tentación. Es el tercer restaurante en ventas ($692.300) y el peor evaluado (3,77). El 43% de sus notas es de 3 estrellas o menos, y la diferencia con el resto no se explica por azar (p = 0,006). En la dispersión del dashboard aparece aislado, abajo a la derecha (figura 4). Otros tres locales están bajo 4,0: Pollo al Carbón, Café Luna y Verde Bowl."),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80 }, keepNext: true, children: [
    imagen("informe/fig_ventas_categoria.png", Math.round(215 * 341 / 529), 215),
    new TextRun("      "),
    imagen("informe/fig_dispersion_restaurantes.png", Math.round(215 * 343 / 391), 215),
  ] }),
  leyenda("Figura 4. Ventas por categoría (página Resumen) y ventas contra calificación por restaurante (página Operación). El punto naranjo es Dulce Tentación, el único local con nota bajo 3,9. Fuente: dashboard."),

  h2("3.5 La Florida y Ñuñoa son las comunas con más espacio para crecer"),
  tabla(
    ["Comuna", "Clientes activos", "Nuevos desde mayo", "Pedidos por cliente", "% de ventas", "Satisfechos", "Locales", "Cancelación de sus locales"],
    [
      ["Santiago", "94", "18", "4,97", "31,3%", "76,5%", "9", "8,4%"],
      ["La Florida", "44", "8", "5,25", "15,5%", "83,0%", "4", "11,8%"],
      ["Providencia", "46", "6", "4,63", "14,5%", "79,4%", "3", "7,3%"],
      ["Ñuñoa", "41", "7", "5,15", "14,5%", "79,6%", "6", "12,8%"],
      ["Las Condes", "43", "2", "5,12", "14,5%", "80,1%", "3", "13,8%"],
      ["Maipú", "30", "9", "4,87", "9,7%", "80,8%", "2", "17,3%"],
    ],
    [1300, 1000, 1100, 1150, 1000, 1150, 1000, 1660],
    { num: [1, 2, 3, 4, 5, 6, 7], junto: true },
  ),
  leyenda("Tabla 1. Indicadores por comuna del cliente; las dos últimas columnas describen los locales ubicados en esa comuna. Fuente: dashboard, página Clientes filtrada por comuna (clientes activos, nuevos usuarios, pedidos por cliente y ventas) y página Operación filtrada por comuna del restaurante (locales y cancelación). Satisfechos usa la fórmula del indicador del dashboard, calculada por comuna del cliente en notebooks/analisis_pideya.ipynb."),
  p("La Florida tiene los clientes más activos (5,25 pedidos), la mayor satisfacción (83,0%) y 41% de cuentas Premium. Ñuñoa tiene la mayor oferta por usuario (6 locales para 41 usuarios) y sus clientes son los que menos cancelan (8,1%). Ahí una campaña tiene más probabilidades de traer clientes que se queden. Maipú crece más que ninguna en proporción (9 de 30 usuarios llegaron desde mayo), pero sus locales cancelan el 17,3%: conviene arreglar esa oferta primero. Las Condes sumó solo 2 usuarios desde mayo, y Santiago tiene casi un tercio de las ventas con la satisfacción más baja (76,5%). Solo el 18,7% de los pedidos va a un local de la misma comuna, así que mejorar un restaurante beneficia a clientes de toda la ciudad."),
];

// ---------- 4. Recomendaciones ----------
const seccion4 = [
  h1("4. Recomendaciones"),
  p("Las cuatro primeras acciones tienen un indicador que se sigue mes a mes en el dashboard; la quinta se sigue en la bitácora de calidad de datos. Las metas son propuestas del equipo y deben validarse con la gerencia."),
  tabla(
    ["#", "Acción", "Qué hacer", "Indicador y meta"],
    [
      ["1", "**Bajar las cancelaciones en cinco restaurantes**", "Revisar disponibilidad, stock y tiempos de preparación con Sushi Mar, Trattoria Roma, Mar y Tierra, La Esquina Criolla y Ramen Ya. Averiguar qué cambió en agosto.", "Cancelación bajo 9,2% en 3 meses y bajo 8% en 6."],
      ["2", "**Mejorar a los restaurantes peor evaluados**", "Plan de mejora con Dulce Tentación, Pollo al Carbón, Café Luna y Verde Bowl; si no mejoran, menos visibilidad en la app.", "Calificación de 4,0 o más en cada local."],
      ["3", "**Captar clientes en La Florida y Ñuñoa**", "Campañas de registro en esas comunas y contacto con los 73 clientes sin pedidos en 45 días.", "14 registros nuevos por mes (nivel de enero a mayo)."],
      ["4", "**Redefinir la cuenta Premium**", "Ligar el beneficio a la frecuencia de compra (por ejemplo, despacho prioritario) y medir si cambia el comportamiento.", "Más pedidos por cliente en Premium que en Estándar."],
      ["5", "**Ordenar la captura de datos**", "Listas cerradas en los formularios, catálogo de categorías corregido con el área comercial, y registro del motivo de cancelación y la hora del pedido.", "Cero correcciones manuales en la carga mensual."],
    ],
    [400, 2300, 4360, 2300],
  ),
  espacio(100),
  p("Las dos primeras acciones protegen las ventas que PideYa ya tiene. La tercera y la cuarta son las que pueden hacerlas crecer, porque casi todos los usuarios registrados ya compran. La quinta permite que el próximo análisis explique por qué se cancela un pedido, no solo dónde."),
];

// ---------- Anexos ----------
const anexos = [
  new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, children: [new TextRun("Anexos")] }),
  h2("A. Indicadores y metas"),
  tabla(
    ["Indicador", "Cómo se calcula", "Valor", "Meta propuesta"],
    [
      ["Ventas totales", "Suma del monto de pedidos entregados", "$14.809.700", "Sin meta"],
      ["Total de pedidos", "Cantidad de pedidos, entregados y cancelados", "1.488", "Sin meta"],
      ["Ticket promedio", "Ventas divididas por pedidos entregados", "$11.160", "Sin meta"],
      ["Tasa de cancelación", "Cancelados divididos por el total de pedidos", "10,8%", "8% o menos"],
      ["Tiempo promedio de entrega", "Promedio de minutos de los pedidos entregados", "31,5 min", "30 min o menos"],
      ["Entregas en 30 minutos", "Entregas de 30 min o menos sobre entregas con tiempo", "45,0%", "Sin meta"],
      ["Calificación promedio", "Promedio de notas de pedidos entregados", "4,14", "4,2 o más"],
      ["Clientes satisfechos", "Notas de 4 o 5 sobre el total de notas", "79,3%", "Sin meta"],
      ["Clientes activos", "Usuarios distintos con al menos un pedido", "298", "Sin meta"],
      ["Pedidos y gasto por cliente", "Pedidos o ventas divididos por clientes activos", "5,0 y $49.697", "Sin meta"],
      ["Ventas Premium", "Ventas de cuentas Premium sobre el total", "26,9%", "Sin meta"],
    ],
    [2500, 3960, 1500, 1400],
    { num: [2] },
  ),
  espacio(),
  p("Las medidas están escritas en DAX (dashboard/medidas_dax.dax) y se agrupan en una tabla de medidas por tema: ventas, servicio, clientes, metas y textos."),

  h2("B. El dashboard"),
  p("El dashboard tiene tres páginas, una por audiencia. Todas comparten la misma estructura: un panel lateral con navegación y filtros, el título con un subtítulo que indica los filtros activos, una fila de indicadores y los gráficos de análisis."),
  vineta("Navegación entre páginas con botones; la página activa se marca en naranjo."),
  vineta("Segmentadores por mes, comuna, categoría y tipo de cuenta, y un botón para limpiarlos todos. En Resumen y Clientes la comuna es la del cliente; en Operación, la del restaurante."),
  vineta("Resaltado cruzado: al hacer clic en una barra, un punto o una celda, el resto de la página se resalta o se filtra por ese elemento."),
  vineta("Colores de indicadores y barras que cambian según la meta, para cualquier combinación de filtros. En la tabla de restaurantes, la cancelación va de verde (más baja) a rojo (más alta)."),
  ...figura("dashboard_1_resumen.png", 1332, 749, 560, "Figura 5. Página Resumen ejecutivo, para la gerencia general: indicadores con semáforo, evolución mensual, ventas por categoría y medios de pago."),
  ...figura("dashboard_2_operacion.png", 1352, 768, 560, "Figura 6. Página Operación y calidad de servicio, para operaciones: cancelación por categoría y comuna, restaurantes según ventas y nota, y tabla de detalle."),
  ...figura("dashboard_3_clientes.png", 1331, 765, 560, "Figura 7. Página Clientes y comportamiento de compra, para marketing: ventas por comuna, demanda por día, cancelación por tipo de cuenta y nuevos usuarios."),

  h2("C. Por qué elegimos cada gráfico"),
  tabla(
    ["Gráfico", "Pregunta que responde", "Por qué este formato"],
    [
      ["Tarjetas con semáforo", "¿Cómo estamos hoy?", "Muestran la cifra exacta y el color indica si cumple la meta sin tener que leer el número."],
      ["Columnas y línea por mes", "¿Cómo evolucionan las ventas y los pedidos?", "Dos magnitudes distintas sobre el mismo eje de tiempo."],
      ["Barras horizontales ordenadas", "¿Quién vende o cancela más?", "La longitud es lo que el ojo compara con más precisión. Horizontales porque los nombres son largos."],
      ["Línea de meta sobre las barras", "¿Se cumple la meta?", "Toda barra que cruza la línea punteada está fuera de meta."],
      ["Anillo de medios de pago", "¿Cómo se reparte el pago?", "Parte de un todo con solo tres categorías, donde el anillo se lee bien."],
      ["Dispersión por restaurante", "¿Hay locales que venden mucho con mala nota?", "Cruza dos variables y deja ver los casos atípicos, como Dulce Tentación."],
      ["Mapa de calor día y comuna", "¿Cuándo y dónde se concentra la demanda?", "La intensidad del color muestra patrones que en una tabla de números no se ven."],
      ["Área de nuevos usuarios", "¿Cómo evoluciona la captación?", "Muestra la tendencia de registros en el tiempo."],
    ],
    [2500, 2700, 4160],
  ),
  espacio(),
  p("El rojo se reserva para lo que no cumple la meta y el verde para lo que sí. El naranjo de la marca se usa para datos sin juicio; en la dispersión, en cambio, marca a los locales con nota bajo 3,9. Evitamos el 3D, los ejes truncados en barras y las tortas con muchas categorías, porque distorsionan la comparación."),

  h2("D. Limitaciones"),
  vineta("Las categorías no siempre corresponden al restaurante (sección 2), lo que debilita cualquier conclusión por categoría."),
  vineta("107 pedidos tienen fecha anterior al registro de su usuario. Los conservamos y los marcamos."),
  vineta("Son menos de seis meses de pedidos y octubre llega solo hasta el día 19, así que no se puede analizar estacionalidad."),
  vineta("No hay hora del pedido, costo de despacho, margen ni motivo de cancelación, por lo que podemos decir dónde están los problemas pero no siempre por qué."),
  vineta("Las metas son propuestas del equipo y no fueron validadas con la empresa."),
  vineta("La tabla de usuarios no tiene registros en octubre. Puede que no haya habido registros o que la exportación se haya hecho antes; conviene confirmarlo con el área que la entregó."),

  h2("E. Archivos del proyecto"),
  tabla(
    ["Carpeta", "Contenido"],
    [
      ["data/raw", "Datos originales, sin modificar."],
      ["data/processed", "Datos limpios que usa el dashboard y bitácora de calidad."],
      ["src/", "limpieza_pideya.py: limpieza, validación y exportación."],
      ["notebooks/", "analisis_pideya.ipynb: retención, captación, comunas, pruebas de diferencia y figuras de este informe."],
      ["dashboard/", "PideYa_dashboard.pbix, medidas DAX, modelo, tema de colores, guía de construcción y valores de control."],
      ["images/", "Capturas del dashboard, bocetos y figuras del informe."],
      ["docs/", "Este informe (PDF y Word) y la presentación ejecutiva."],
    ],
    [2200, 7160],
  ),
];

// ---------- documento ----------
const pie = new Footer({ children: [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: ANCHO }], children: [
  new TextRun({ text: "PideYa · Informe ejecutivo · Evaluación Parcial 2", size: 16, color: GRIS }),
  new TextRun({ children: ["\t", PageNumber.CURRENT], size: 16, color: GRIS }),
] })] });

const pagina = { size: { width: 12240, height: 15840 }, margin: { top: 1152, bottom: 1152, left: 1440, right: 1440 } };

const doc = new Document({
  creator: "Equipo consultor PideYa",
  title: "Informe ejecutivo PideYa · Evaluación Parcial 2",
  styles: {
    default: { document: { run: { font: "Calibri", size: 21, color: OSCURO } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Calibri", size: 30, bold: true, color: OSCURO }, paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 0, keepNext: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Calibri", size: 24, bold: true, color: NARANJO }, paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 1, keepNext: true } },
    ],
  },
  numbering: { config: [{ reference: "vinetas", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
  sections: [
    { properties: { page: pagina }, children: portada },
    { properties: { page: pagina }, footers: { default: pie }, children: [...seccion1, ...seccion2, ...seccion3, ...seccion4, ...anexos] },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("OK", OUT);
});
