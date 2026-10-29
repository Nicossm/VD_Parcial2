const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/2e960819-a9ad-4364-af96-e96237836b88_2f1d191a-5248-4062-81a4-53356b5b7fa5/pptx/scripts/apply_theme.js");
const path = require("path");
const IMG = path.resolve(__dirname, "../../images");
const OUT = path.resolve(__dirname, "../Presentacion_Ejecutiva_PideYa.pptx");

const THEME = {
  name: "PideYa",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "1F2937", lt1: "FFFFFF", dk2: "4B5563", lt2: "F3F4F6",
    accent1: "E4572E", accent2: "2E5EAA", accent3: "C2410C", accent4: "15803D",
    accent5: "9CA3AF", accent6: "17A398", hlink: "2E5EAA", folHlink: "6C4E97",
  },
};
const HEX = { ink: "1F2937", mut: "4B5563", soft: "6B7280", line: "E5E7EB", or: "E4572E", red: "C2410C", grn: "15803D", bl: "2E5EAA", gray: "9CA3AF", tint: "FFF4EF", lt2: "F3F4F6" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "PideYa · Presentación ejecutiva";
pres.author = "Javier Sagredo, Cristian Romero, Rolando Paredes, Nicolás Osses";
const C = pres.SchemeColor;

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "DARK",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.3, w: 11.2, h: 1.9, fontSize: 44, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.25, w: 10.5, h: 0.9, fontSize: 18, color: C.accent5, valign: "top", align: "left", margin: 0 }, text: "" } },
    { text: { text: "PideYa", options: { x: 0.8, y: 0.6, w: 3, h: 0.5, fontSize: 24, bold: true, color: C.accent1, margin: 0 } } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  slideNumber: { x: 12.3, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 0.9, fontSize: 28, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: "PideYa · Presentación ejecutiva · ADY1104-001D", options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
});

// ---------- helpers ----------
function card(slide, x, y, w, h, fill) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill || C.background2 }, line: { color: fill || C.background2 }, objectName: "card" });
}
function stat(slide, x, y, w, value, label, color, sub) {
  card(slide, x, y, w, 1.55);
  slide.addText(value, { x: x + 0.25, y: y + 0.15, w: w - 0.5, h: 0.8, fontSize: 36, bold: true, color: color || C.text1, margin: 0, isTextBox: true, fit: "shrink" });
  slide.addText(label, { x: x + 0.25, y: y + 0.92, w: w - 0.5, h: 0.32, fontSize: 13, color: C.text2, margin: 0, isTextBox: true });
  if (sub) slide.addText(sub, { x: x + 0.25, y: y + 1.2, w: w - 0.5, h: 0.28, fontSize: 11, color: C.accent5, margin: 0, isTextBox: true });
}
const chartBase = {
  catAxisLabelColor: HEX.mut, valAxisLabelColor: HEX.soft, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, dataLabelFontFace: "+mn-lt", dataLabelFontSize: 12, dataLabelColor: HEX.ink,
  valGridLine: { color: HEX.line, size: 0.75 }, catGridLine: { style: "none" }, showLegend: false,
  titleFontFace: "+mn-lt", titleFontSize: 14, titleColor: HEX.ink,
};

// ---------- 1. Portada ----------
pres.addSection({ title: "Inicio" });
let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Inicio" });
s.addText("Desempeño comercial y calidad de servicio", { placeholder: "title" });
s.addText("Resumen ejecutivo · Dashboard de visualización de datos en Power BI", { placeholder: "body" });
s.addText([
  { text: "ADY1104 · Visualización de Datos · Sección ADY1104-001D", options: { breakLine: true } },
  { text: "Javier Sagredo · Cristian Romero · Rolando Paredes · Nicolás Osses", options: { breakLine: true } },
  { text: "Duoc UC · Octubre 2026" },
], { x: 0.8, y: 5.7, w: 11, h: 1.1, fontSize: 14, color: C.background2, margin: 0, isTextBox: true, paraSpaceAfter: 4 });
s.addNotes("Presentamos el análisis de PideYa, app de delivery que opera en seis comunas de Santiago. Integramos tres fuentes de datos y construimos un dashboard en Power BI para apoyar decisiones de operación, servicio y clientes.");

// ---------- 2. Mensaje central ----------
pres.addSection({ title: "Resumen" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Resumen" });
s.addText("Ventas estables, pero el servicio no cumple sus metas", { placeholder: "title" });
stat(s, 0.6, 1.6, 2.85, "$14,8 M", "Ventas (may–oct 2026)", C.text1, "1.327 pedidos entregados");
stat(s, 3.68, 1.6, 2.85, "10,8%", "Tasa de cancelación", C.accent3, "Meta: 8%");
stat(s, 6.76, 1.6, 2.85, "45%", "Entregas en ≤ 30 min", C.accent3, "Promedio: 31,5 min");
stat(s, 9.84, 1.6, 2.85, "4,14", "Calificación promedio", C.accent3, "Meta: 4,2");
s.addText("Lo que encontramos", { x: 0.6, y: 3.55, w: 6, h: 0.45, fontSize: 20, bold: true, color: C.text1, margin: 0, isTextBox: true });
s.addText([
  { text: "Ninguna categoría cumple la meta de cancelación; Maipú llega a 17,3%", options: { bullet: true, breakLine: true } },
  { text: "Más de la mitad de las entregas llega después de 30 minutos", options: { bullet: true, breakLine: true } },
  { text: "La calificación depende del restaurante, no de la rapidez", options: { bullet: true, breakLine: true } },
  { text: "El segmento Premium cancela más que el Estándar", options: { bullet: true } },
], { x: 0.6, y: 4.05, w: 7.2, h: 2.6, fontSize: 16, color: C.text2, paraSpaceAfter: 8, valign: "top", isTextBox: true });
card(s, 8.3, 3.6, 4.4, 3.0, C.text1);
s.addText([
  { text: "Mensaje clave", options: { fontSize: 13, color: C.accent1, bold: true, breakLine: true } },
  { text: "El crecimiento no está limitado por la demanda, sino por la calidad del servicio.", options: { fontSize: 20, color: C.background1, bold: true } },
], { x: 8.6, y: 3.85, w: 3.85, h: 2.5, margin: 0, valign: "top", isTextBox: true, paraSpaceAfter: 8 });
s.addNotes("Tres de los seis indicadores principales están en rojo. El mensaje central: PideYa no tiene un problema de demanda, tiene un problema de servicio.");

// ---------- 3. Datos ----------
pres.addSection({ title: "Contexto" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Contexto" });
s.addText("Integramos tres sistemas que nunca se habían consolidado", { placeholder: "title" });
const src = [["Usuarios", "308 → 300 filas", "comuna · tipo de cuenta"], ["Restaurantes", "30 → 30 filas", "categoría · comuna"], ["Pedidos", "1.515 → 1.488 filas", "monto · tiempo · calificación"]];
src.forEach((r, i) => {
  const y = 1.65 + i * 1.3;
  card(s, 0.6, y, 3.6, 1.1);
  s.addText(r[0], { x: 0.85, y: y + 0.12, w: 3.2, h: 0.4, fontSize: 18, bold: true, color: C.text1, margin: 0, isTextBox: true });
  s.addText(r[1] + "  ·  " + r[2], { x: 0.85, y: y + 0.55, w: 3.3, h: 0.4, fontSize: 12, color: C.text2, margin: 0, isTextBox: true });
});
s.addShape(pres.shapes.RIGHT_ARROW, { x: 4.45, y: 2.95, w: 0.7, h: 0.5, fill: { color: C.accent5 }, line: { color: C.accent5 }, objectName: "arrow1" });
card(s, 5.4, 1.65, 3.4, 3.7, C.accent1);
s.addText([
  { text: "Limpieza en Python", options: { fontSize: 18, bold: true, color: C.background1, breakLine: true } },
  { text: "743 fechas en 3 formatos", options: { bullet: true, breakLine: true } },
  { text: "313 textos inconsistentes", options: { bullet: true, breakLine: true } },
  { text: "149 montos con texto", options: { bullet: true, breakLine: true } },
  { text: "23 duplicados", options: { bullet: true, breakLine: true } },
  { text: "12 claves inexistentes", options: { bullet: true, breakLine: true } },
  { text: "12 tiempos imposibles", options: { bullet: true } },
], { x: 5.65, y: 1.85, w: 3.0, h: 3.3, fontSize: 14, color: C.background1, valign: "top", margin: 0, paraSpaceAfter: 5, isTextBox: true });
s.addShape(pres.shapes.RIGHT_ARROW, { x: 9.05, y: 2.95, w: 0.7, h: 0.5, fill: { color: C.accent5 }, line: { color: C.accent5 }, objectName: "arrow2" });
card(s, 10.0, 1.65, 2.7, 3.7);
s.addText([
  { text: "Modelo estrella", options: { fontSize: 18, bold: true, color: C.text1, breakLine: true } },
  { text: "Pedidos al centro, unidos a Usuarios, Restaurantes y un Calendario", options: { fontSize: 13, color: C.text2, breakLine: true } },
  { text: "26 medidas DAX", options: { fontSize: 13, color: C.accent1, bold: true } },
], { x: 10.2, y: 1.85, w: 2.35, h: 3.3, valign: "top", margin: 0, paraSpaceAfter: 10, isTextBox: true });
s.addText("Proceso reproducible: el script registra cada corrección en una bitácora de calidad y los KPIs se validaron contra cálculos independientes.", { x: 0.6, y: 5.75, w: 12.1, h: 0.6, fontSize: 14, italic: true, color: C.text2, margin: 0, isTextBox: true });
s.addNotes("Los datos venían con fechas en tres formatos, nombres escritos de distintas formas, montos como texto, duplicados y valores imposibles. Todo se corrigió con un script de Python reproducible que deja una bitácora. Luego armamos un modelo estrella en Power BI.");

// ---------- 4. Ventas ----------
pres.addSection({ title: "Hallazgos" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Hallazgos" });
s.addText("Las ventas son estables: entre $2,5 y $2,8 millones al mes", { placeholder: "title" });
s.addChart(pres.charts.BAR, [{ name: "Ventas (M$)", labels: ["May", "Jun", "Jul", "Ago", "Sep", "Oct*"], values: [2.82, 2.55, 2.72, 2.50, 2.68, 1.55] }], {
  ...chartBase, x: 0.6, y: 1.5, w: 7.6, h: 5.1, barDir: "col", chartColors: [HEX.or], showValue: true, dataLabelPosition: "outEnd",
  dataLabelFormatCode: "$0.0 \"M\"", valAxisLabelFormatCode: "$0.0", valAxisMinVal: 0, valAxisMaxVal: 3.5, barGapWidthPct: 60,
  showTitle: true, title: "Ventas mensuales de pedidos entregados (millones de pesos)", objectName: "ventas_mes",
});
card(s, 8.6, 1.6, 4.1, 4.9);
s.addText([
  { text: "Qué vende PideYa", options: { fontSize: 18, bold: true, color: C.text1, breakLine: true } },
  { text: "Saludable lidera con el 35% de las ventas", options: { bullet: true, breakLine: true } },
  { text: "Santiago concentra el 31% de las ventas", options: { bullet: true, breakLine: true } },
  { text: "58% de los pedidos se paga con tarjeta", options: { bullet: true, breakLine: true } },
  { text: "Ticket promedio estable: $11.160", options: { bullet: true, breakLine: true } },
  { text: "Sábado es el día de más pedidos", options: { bullet: true } },
], { x: 8.85, y: 1.85, w: 3.65, h: 4.4, fontSize: 14, color: C.text2, valign: "top", margin: 0, paraSpaceAfter: 9, isTextBox: true });
s.addText("*Octubre con datos solo hasta el día 19", { x: 0.6, y: 6.6, w: 6, h: 0.3, fontSize: 10, color: C.accent5, margin: 0, isTextBox: true });
s.addNotes("Las ventas mensuales son estables, sin tendencia de crecimiento. La baja de octubre no es real: los datos llegan solo hasta el 19. Saludable es la categoría líder y Santiago la comuna con más ventas.");

// ---------- 5. Cancelación ----------
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Hallazgos" });
s.addText("Solo Providencia cumple la meta de cancelación de 8%", { placeholder: "title" });
s.addChart(pres.charts.BAR, [{ name: "Cancelación", labels: ["Maipú", "Las Condes", "Ñuñoa", "La Florida", "Sin información", "Santiago", "Providencia"], values: [0.173, 0.138, 0.128, 0.118, 0.092, 0.084, 0.073] }], {
  ...chartBase, x: 0.6, y: 1.5, w: 7.6, h: 5.2, barDir: "bar", chartColors: [HEX.red, HEX.red, HEX.red, HEX.red, HEX.red, HEX.red, HEX.grn], invertedColors: [HEX.red],
  showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0%", valAxisLabelFormatCode: "0%", valAxisMinVal: 0, valAxisMaxVal: 0.2,
  catAxisOrientation: "maxMin", barGapWidthPct: 45, showTitle: true, title: "Tasa de cancelación por comuna del restaurante (meta 8%)", objectName: "cancel_comuna",
});
stat(s, 8.6, 1.6, 4.1, "17,3%", "Cancelación en Maipú", C.accent3, "Más del doble de la meta");
stat(s, 8.6, 3.35, 4.1, "13,6%", "Cancelación en comida Asiática", C.accent3, "La categoría más alta");
card(s, 8.6, 5.1, 4.1, 1.55, C.accent4);
s.addText("Agosto bajó a 7,0%: la meta es alcanzable", { x: 8.85, y: 5.2, w: 3.65, h: 1.35, fontSize: 16, bold: true, color: C.background1, margin: 0, valign: "middle", isTextBox: true });
s.addNotes("La cancelación promedio es 10,8%. Ninguna categoría cumple la meta de 8% y, por comuna, solo Providencia lo logra. Maipú llega a 17,3%. Agosto bajó a 7%, lo que demuestra que la meta es alcanzable.");

// ---------- 6. Entregas y satisfacción ----------
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Hallazgos" });
s.addText("El 55% de las entregas llega tarde, pero eso no baja la nota", { placeholder: "title" });
stat(s, 0.6, 1.6, 3.6, "55%", "Entregas sobre 30 minutos", C.accent3, "Mediana: 32 minutos");
stat(s, 0.6, 3.35, 3.6, "≈ 0", "Correlación tiempo–calificación", C.text1, "La rapidez no explica la nota");
card(s, 0.6, 5.1, 3.6, 1.55, C.text1);
s.addText("Satisfacción depende del restaurante, no del reparto", { x: 0.85, y: 5.2, w: 3.1, h: 1.35, fontSize: 16, bold: true, color: C.background1, margin: 0, valign: "middle", isTextBox: true });
s.addChart(pres.charts.BAR, [{ name: "Calificación", labels: ["≤20 min", "21–30", "31–40", "41–50", ">50 min"], values: [4.19, 4.12, 4.13, 4.22, 4.16] }], {
  ...chartBase, x: 4.6, y: 1.5, w: 8.1, h: 5.2, barDir: "col", chartColors: [HEX.bl], showValue: true, dataLabelPosition: "outEnd",
  dataLabelFormatCode: "0.00", valAxisMinVal: 0, valAxisMaxVal: 5, valAxisMajorUnit: 1, barGapWidthPct: 55,
  showTitle: true, title: "Calificación promedio (1–5) según tiempo de entrega", objectName: "calif_tiempo",
});
s.addNotes("Solo el 45% de las entregas llega en 30 minutos o menos. Sin embargo, la calificación es prácticamente igual en todos los rangos de tiempo. Acelerar las entregas no subiría la satisfacción por sí solo.");

// ---------- 7. Restaurante y Premium ----------
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Hallazgos" });
s.addText("Dos alertas: un restaurante clave y el segmento Premium", { placeholder: "title" });
card(s, 0.6, 1.6, 5.9, 5.0);
s.addText("Dulce Tentación", { x: 0.9, y: 1.8, w: 5.3, h: 0.5, fontSize: 22, bold: true, color: C.text1, margin: 0, isTextBox: true });
s.addText("3.er restaurante en ventas, con la peor calificación", { x: 0.9, y: 2.3, w: 5.3, h: 0.4, fontSize: 14, color: C.text2, margin: 0, isTextBox: true });
[["$692 mil", "Ventas", C.text1], ["3,77", "Calificación", C.accent3], ["43%", "Notas ≤ 3★", C.accent3]].forEach((v, i) => {
  const x = 0.9 + i * 1.85;
  s.addText(v[0], { x, y: 3.1, w: 1.75, h: 0.8, fontSize: 26, bold: true, color: v[2], margin: 0, isTextBox: true, fit: "shrink" });
  s.addText(v[1], { x, y: 3.9, w: 1.7, h: 0.35, fontSize: 12, color: C.text2, margin: 0, isTextBox: true });
});
s.addText("Un local con alto volumen y mala experiencia daña la marca en cada pedido.", { x: 0.9, y: 4.8, w: 5.3, h: 1.4, fontSize: 15, italic: true, color: C.text2, margin: 0, valign: "top", isTextBox: true });
s.addChart(pres.charts.BAR, [{ name: "Cancelación", labels: ["Premium", "Estándar"], values: [0.124, 0.108] }], {
  ...chartBase, x: 6.9, y: 1.5, w: 5.8, h: 3.3, barDir: "bar", chartColors: [HEX.red], showValue: true, dataLabelPosition: "outEnd",
  dataLabelFormatCode: "0.0%", valAxisLabelFormatCode: "0%", valAxisMinVal: 0, valAxisMaxVal: 0.15, catAxisOrientation: "maxMin", barGapWidthPct: 50,
  showTitle: true, title: "Tasa de cancelación por tipo de cuenta", objectName: "cancel_premium",
});
s.addText([
  { text: "Premium cancela más y no gasta más por pedido ", options: { bold: true, color: C.text1 } },
  { text: "($11.081 vs $11.173). Representa el 27% de las ventas: son los clientes que más conviene retener.", options: { color: C.text2 } },
], { x: 6.9, y: 5.0, w: 5.8, h: 1.5, fontSize: 15, margin: 0, valign: "top", isTextBox: true });
s.addNotes("Dulce Tentación vende mucho pero tiene la peor nota: 43% de sus calificaciones son de 3 estrellas o menos. Además, los usuarios Premium cancelan más que los Estándar y no gastan más, por lo que la cuenta Premium no está fidelizando.");

// ---------- 8. Dashboard ----------
pres.addSection({ title: "Solución" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Solución" });
s.addText("Un dashboard con tres páginas, una por audiencia", { placeholder: "title" });
const pages = [["dashboard_1_resumen.png", "1 · Resumen ejecutivo", "Gerencia general: KPIs con semáforo y evolución"], ["dashboard_2_operacion.png", "2 · Operación", "Operaciones: cancelación, tiempos y restaurantes"], ["dashboard_3_clientes.png", "3 · Clientes", "Marketing: comunas, días y segmento Premium"]];
pages.forEach((p, i) => {
  const x = 0.6 + i * 4.1;
  s.addImage({ path: path.join(IMG, p[0]), x, y: 1.6, w: 3.85, h: 3.85 * 749 / 1332, objectName: "shot" + i, altText: "Captura de la página " + p[1] });
  s.addText(p[1], { x, y: 3.9, w: 3.85, h: 0.4, fontSize: 16, bold: true, color: C.text1, margin: 0, isTextBox: true });
  s.addText(p[2], { x, y: 4.3, w: 3.85, h: 0.6, fontSize: 13, color: C.text2, margin: 0, valign: "top", isTextBox: true });
});
card(s, 0.6, 5.2, 12.1, 1.45);
s.addText([
  { text: "Interactividad: ", options: { bold: true, color: C.accent1 } },
  { text: "navegación entre páginas · filtros por mes, comuna, categoría y tipo de cuenta · botón limpiar filtros · filtro cruzado al hacer clic · subtítulo que indica los filtros activos · colores que cambian según la meta", options: { color: C.text1 } },
], { x: 0.85, y: 5.3, w: 11.6, h: 1.25, fontSize: 15, margin: 0, valign: "middle", isTextBox: true });
s.addNotes("La solución es un dashboard de tres páginas. Cada una responde la pregunta de una audiencia distinta. Los KPIs usan semáforos según la meta y todos los filtros actualizan la página completa.");

// ---------- 9. Recomendaciones ----------
pres.addSection({ title: "Cierre" });
s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Cierre" });
s.addText("Cuatro acciones para mejorar el servicio", { placeholder: "title" });
const recs = [
  ["1", "Plan focalizado de cancelaciones", "Maipú, Las Condes y comida Asiática", "Cancelación ≤ 8% en 3 meses"],
  ["2", "Programa de calidad", "Dulce Tentación y locales bajo 4,0★", "Calificación por local ≥ 4,0"],
  ["3", "Rediseñar Premium", "Prioridad de despacho y compensación por demoras", "Cancelación Premium ≤ Estándar"],
  ["4", "Mejorar la captura de datos", "Listas cerradas y motivo de cancelación", "Cero correcciones manuales"],
];
recs.forEach((r, i) => {
  const x = 0.6 + (i % 2) * 6.15, y = 1.6 + Math.floor(i / 2) * 2.55;
  card(s, x, y, 5.95, 2.35);
  s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: "num" + r[0] });
  s.addText(r[0], { x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7, fontSize: 20, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addText(r[1], { x: x + 1.25, y: y + 0.3, w: 4.5, h: 0.5, fontSize: 19, bold: true, color: C.text1, margin: 0, isTextBox: true });
  s.addText(r[2], { x: x + 1.25, y: y + 0.85, w: 4.5, h: 0.6, fontSize: 14, color: C.text2, margin: 0, valign: "top", isTextBox: true });
  s.addText([{ text: "Meta: ", options: { bold: true, color: C.accent4 } }, { text: r[3], options: { color: C.text1 } }], { x: x + 1.25, y: y + 1.6, w: 4.5, h: 0.45, fontSize: 14, margin: 0, isTextBox: true });
});
s.addNotes("Proponemos cuatro acciones, cada una con un indicador que se puede seguir en el dashboard: reducir cancelaciones donde se concentran, mejorar la calidad de los restaurantes críticos, rediseñar los beneficios Premium y mejorar la captura de datos.");

// ---------- 10. Cierre ----------
s = pres.addSlide({ masterName: "DARK", sectionTitle: "Cierre" });
s.addText("Retener mejor a los clientes que ya tenemos", { placeholder: "title" });
s.addText("El dashboard entrega a cada área los indicadores para monitorear ese objetivo de forma continua.", { placeholder: "body" });
s.addText("Gracias · ¿Preguntas?", { x: 0.8, y: 6.0, w: 8, h: 0.6, fontSize: 20, bold: true, color: C.accent1, margin: 0, isTextBox: true });
s.addNotes("En síntesis: PideYa no necesita más clientes para crecer, sino retener mejor a los que ya tiene. Quedamos atentos a preguntas.");

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT);
})();
