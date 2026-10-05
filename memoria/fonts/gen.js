const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, Header, Footer, PageNumber, PageBreak, ImageRun,
  LevelFormat, TableOfContents, PageOrientation, TabStopType, VerticalAlign,
} = require('docx');

const PL = path.join(__dirname, 'plans');
const OUT = process.argv[2];

const FONT = 'Arial';
const TITOL = "PROJECTE DE SUBSTITUCIÓ DE LES PORTES I ADAPTACIÓ DE LA TANCA D’ACCÉS AL RECINTE DE LA NAU DE TURBINES DE LA CENTRAL TÈRMICA DEL BESÒS (AVINGUDA D’EDUARD MARISTANY)";
const DATA = 'Octubre 2026';

// ---------- helpers ----------
const P = (text, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.JUSTIFIED,
  spacing: { after: opts.after ?? 120, before: opts.before ?? 0, line: 300 },
  indent: opts.indent,
  keepNext: opts.keepNext,
  children: Array.isArray(text) ? text : [new TextRun({ text, bold: opts.bold, italics: opts.italics, size: opts.size, highlight: opts.hl })],
});
const R = (text, o = {}) => new TextRun({ text, ...o });
const PEND = (text) => new TextRun({ text, highlight: 'yellow' });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [R(t)], pageBreakBefore: false });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [R(t)] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [R(t)] });
const B = (text, lvl = 0) => new Paragraph({
  numbering: { reference: 'bul', level: lvl }, alignment: AlignmentType.JUSTIFIED,
  spacing: { after: 60, line: 300 },
  children: Array.isArray(text) ? text : [R(text)],
});
let LIST = 0; const NEWLIST = () => { LIST++; return []; };
const N = (text) => new Paragraph({
  numbering: { reference: 'num', level: 0, instance: LIST }, alignment: AlignmentType.JUSTIFIED,
  spacing: { after: 60, line: 300 }, children: Array.isArray(text) ? text : [R(text)],
});
const BR = () => new Paragraph({ children: [new PageBreak()] });
const SP = (n = 1) => Array.from({ length: n }, () => new Paragraph({ children: [] }));

const border = { style: BorderStyle.SINGLE, size: 4, color: '999999' };
const borders = { top: border, bottom: border, left: border, right: border };
function table(widths, rows, { header = true, fontSize = 18, boldLast = false } = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r, ri) => new TableRow({
      tableHeader: header && ri === 0,
      cantSplit: true,
      children: r.map((c, ci) => {
        const isHead = header && ri === 0;
        const isLast = boldLast && ri === rows.length - 1;
        const cell = (typeof c === 'object' && c !== null && !Array.isArray(c)) ? c : { t: c };
        return new TableCell({
          borders,
          width: { size: widths[ci], type: WidthType.DXA },
          shading: isHead ? { fill: 'D9E2EC', type: ShadingType.CLEAR, color: 'auto' } : (cell.fill ? { fill: cell.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined),
          margins: { top: 50, bottom: 50, left: 90, right: 90 },
          verticalAlign: VerticalAlign.CENTER,
          children: String(cell.t).split('\n').map(line => new Paragraph({
            alignment: cell.a || (ci > 0 && /^[\d.,\-–%€ ]+$/.test(String(cell.t)) ? AlignmentType.RIGHT : AlignmentType.LEFT),
            children: [new TextRun({ text: line, bold: isHead || isLast || cell.b, size: fontSize, highlight: cell.hl })],
          })),
        });
      }),
    })),
  });
}
const eur = (v) => v.toLocaleString('ca-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const num = (v, d = 2) => v.toLocaleString('ca-ES', { minimumFractionDigits: d, maximumFractionDigits: d });
const img = (file, w, h) => new ImageRun({ type: 'jpg', data: fs.readFileSync(path.join(PL, file)), transformation: { width: w, height: h } });

// ---------- EVACUATION FIGURES ----------
const post = 0.10;                 // muntants 10x20 cm (10 cm en el sentit de la tanca)
const opening = 15.87;             // llum total entre paraments (plànol 02)
const nPost = 5;                   // pilars dins la llum
const clearTotal = +(opening - nPost * post).toFixed(2);   // 15.12
const pedEach = +(2.50 - post).toFixed(2);                 // 2.35
const pedTotal = +(4 * pedEach).toFixed(2);                // 9.40
const vehClear = +(clearTotal - pedTotal).toFixed(2);      // 5.72
const Pint = 1805;                 // aforament interior (projecte modificat)
const Pmax = 2 * Pint;             // hipòtesi conservadora interior + exterior simultanis

// ---------- BUDGET ----------
// [codi, ud, descripció, amidament(text detall), quantitat, preu, industrial?]
const chapters = [
  { c: '01', t: 'TREBALLS PREVIS, PROTECCIONS I DETECCIÓ DE SERVEIS', items: [
    ['01.01', 'u', 'Localització i marcatge en superfície de la línia elèctrica soterrada existent i d’altres serveis, mitjançant detector de cables i georadar, per empresa especialitzada, incloent informe gràfic de traçat i fondària i replanteig en obra.', '1 u', 1, 480.00],
    ['01.02', 'u', 'Cala manual de comprovació de la posició de la línia elèctrica soterrada, de 0,60×0,60 m fins a 0,80 m de fondària, amb eines manuals, sense mitjans mecànics, i reblert posterior amb sorra garbellada i compactació manual.', '1 cala per muntant nou: 7 u', 7, 85.00],
    ['01.03', 'm', 'Tanca provisional mòbil d’obra de malla electrosoldada galvanitzada de 3,50×2,00 m sobre bases de formigó, per mantenir el tancament del recinte i la delimitació de la zona de treball durant tota l’obra, incloent muntatge, lloguer i desmuntatge.', '15,87 + 2×3,00 laterals = 21,87 m', 21.87, 14.50],
    ['01.04', 'u', 'Senyalització d’obra i de desviament de vianants a la vorera de l’Av. d’Eduard Maristany (tanques de vianants, cartells i abalisament lluminós), durant tota l’obra.', '1 u', 1, 320.00],
  ]},
  { c: '02', t: 'ENDERROCS I DESMUNTATGES', items: [
    ['02.01', 'u', 'Desmuntatge de porta metàl·lica practicable de dues fulles batents existent, de fins a 4,00×3,00 m, malmesa per impacte de vehicles, amb mitjans manuals i auxiliars (camió grua), tall de frontisses i panys, i càrrega sobre camió.', '4 portes existents', 4, 185.00],
    ['02.02', 'u', 'Arrencada de pilar metàl·lic existent de suport de portes, amb tall a ran de paviment o extracció del dau, amb mitjans manuals i càrrega sobre camió.', '5 pilars existents', 5, 48.00],
    ['02.03', 'm3', 'Enderroc de daus de fonamentació de formigó dels pilars existents, amb martell elèctric manual (prohibit l’ús de retroexcavadora a menys d’1,00 m de la línia elèctrica), i càrrega manual de runa sobre contenidor.', '5 u × 0,40×0,40×0,50 = 0,40 m3', 0.40, 190.00],
    ['02.04', 'm2', 'Arrencada de paviment de panot o formigó existent a l’entorn dels nous muntants, amb mitjans manuals, i càrrega de runa.', '7 u × 1,00×1,00 = 7,00 m2', 7.00, 16.50],
  ]},
  { c: '03', t: 'MOVIMENT DE TERRES', items: [
    ['03.01', 'm3', 'Excavació de pou per a fonamentació de pilar, de 0,60×0,60×0,80 m, en terreny compacte, amb mitjans manuals, i càrrega manual sobre contenidor.', '7 u × 0,60×0,60×0,80 = 2,02 m3', 2.02, 98.00],
  ]},
  { c: '04', t: 'FONAMENTACIÓ', items: [
    ['04.01', 'm3', 'Formigó per a daus de fonamentació HA-25/B/20/IIa, abocat des de camió o amb mitjans manuals i vibrat, inclosa capa de neteja HL-150 de 5 cm.', '7 u × 0,60×0,60×0,80 = 2,02 m3', 2.02, 168.00],
    ['04.02', 'u', 'Placa d’ancoratge d’acer S275JR de 250×350×15 mm amb 4 perns d’acer corrugat B500S de Ø16 mm i 50 cm de longitud amb patilla, galvanitzada en calent, col·locada anivellada abans del formigonat.', '7 u', 7, 92.00],
  ]},
  { c: '05', t: 'TANCAMENTS PRACTICABLES: PORTES', industrial: true, items: [
    ['05.01', 'u', 'Porta batent de dues fulles d’acer galvanitzat en calent, de 2.500×3.000 mm (2,40 m de llum de pas lliure entre muntants), amb fulles d’armadura tubular d’acer amb travessers rigiditzadors i folrat de perfils d’acer galvanitzat; gir amb cassoleta superior i pivot inferior amb rodaments de boles, allotjat en muntants laterals de perfil d’acer de 100×200 mm dimensionats segons el pes i les dimensions de les fulles. Obertura cap a l’exterior. Tipus JAM o equivalent. Subministrament i col·locació.', '4 u', 4, 3857.84, 'ind'],
    ['05.02', 'u', 'Porta corredissa d’una fulla d’acer galvanitzat en calent, de 5.770×3.000 mm, amb fulla d’armadura tubular d’acer amb travessers rigiditzadors i folrat de perfils d’acer galvanitzat laminats en fred; rodes d’acer tornejades amb rodaments registrables; guia inferior calibrada de Ø 20 mm sobre platines ancorades al paviment; torreta de suport amb rodaments de niló a la part superior; perfil en “U” de tancament. Motoritzada amb motor electromecànic de 230 V per a fulles de fins a 1.100 kg, amb maniobra incorporada, finals de cursa, cremallera, fotocèl·lula, banda de seguretat, avisador òptic i desbloqueig manual en cas de tall de subministrament elèctric. Marcat CE segons UNE-EN 13241. Tipus JAM o equivalent. Subministrament i col·locació.', '1 u', 1, 8552.41, 'ind'],
    ['05.03', 'u', 'Kit antipànic per a porta exterior de dues fulles batents, segons UNE-EN 1125: barra horitzontal d’empenta amb pany de cop a la fulla activa, barra antipànic de dos punts (tancament superior i inferior) a la fulla passiva, comandament exterior amb maneta i cilindre de clau, acabat apte per a exterior (inox o galvanitzat), inclosa l’adaptació del bastidor de la porta i la regulació. Col·locat.', '4 portes peatonals', 4, 980.00],
    ['05.04', 'u', 'Senyalització dels muntants amb bandes reflectants contra impactes de vehicles, i senyalització de “Sortida d’emergència” fotoluminescent a cada porta peatonal.', '4 u', 4, 45.00],
    ['05.05', 'PA', 'Ajudes de ram de paleta al muntatge de les portes: recepció i aplomat de muntants i torreta, fixació de la guia inferior de la corredissa amb tacs químics sobre el paviment existent i reblert de juntes.', '1 PA', 1, 650.00],
  ]},
  { c: '06', t: 'PAVIMENTS I REPOSICIONS', items: [
    ['06.01', 'm2', 'Reposició de paviment de panot igual a l’existent (o formigó raspallat HM-20 de 15 cm), a l’entorn dels nous muntants, inclosa base de morter i acord amb el paviment existent.', '7,00 m2', 7.00, 72.00],
    ['06.02', 'm', 'Segellat del contorn de pilars amb el paviment amb massilla de poliuretà.', '7 u × 0,60 m = 4,20 m', 4.20, 9.50],
  ]},
  { c: '07', t: 'GESTIÓ DE RESIDUS', items: [
    ['07.01', 't', 'Transport i lliurament a gestor autoritzat de residus metàl·lics (portes i pilars existents, LER 17 04 05), per a valorització.', '4 portes + 5 pilars ≈ 1,20 t', 1.20, 35.00],
    ['07.02', 'm3', 'Transport i deposició controlada a gestor autoritzat de residus de formigó i paviment (LER 17 01 01), amb contenidor, inclòs cànon.', '0,40 + 7,00×0,15 = 1,45 m3 × 1,30 esponjament = 1,89 m3', 1.89, 52.00],
    ['07.03', 'm3', 'Transport i deposició controlada a gestor autoritzat de terres d’excavació no especials (LER 17 05 04), amb contenidor, inclòs cànon.', '2,02 × 1,25 esponjament = 2,53 m3', 2.53, 32.00],
  ]},
  { c: '09', t: 'INSTAL·LACIÓ ELÈCTRICA DEL MOTOR DE LA PORTA CORREDISSA', items: [
    ['09.01', 'PA', 'Alimentació elèctrica monofàsica 230 V al motor de la porta corredissa (380 W), des del quadre elèctric existent a l’extrem esquerre de la tanca (element 7 dels plànols): línia de cable RZ1-K 3G2,5 mm² passada pel conducte soterrat existent Ø 50 mm (element 8), sense obrir rasa nova, protecció magnetotèrmica de 10 A i diferencial de 30 mA al quadre 7, caixa de connexió al costat del motor i connexió del quadre de maniobra. Inclou el butlletí de la instal·lació elèctrica.', '1 PA', 1, 950.00],
  ]},
  { c: '08', t: 'SEGURETAT I SALUT', items: [
    ['08.01', 'PA', 'Partida alçada de seguretat i salut segons Estudi Bàsic de Seguretat i Salut: protecció individual i col·lectiva, senyalització i instal·lacions provisionals.', '1 PA', 1, 450.00],
  ]},
];

chapters.sort((a, b) => a.c.localeCompare(b.c));
// Seguretat i salut: mínim 2 % del PEM (arrodonit a l'alça a 10 €)
{
  const ssItem = chapters.find(c => c.c === '08').items[0];
  const others = chapters.filter(c => c.c !== '08').reduce((t, c) => t + c.items.reduce((u, it) => u + it[4] * it[5], 0), 0);
  ssItem[5] = Math.ceil((0.02 * others / 0.98) / 10) * 10;
}
let pem = 0;
chapters.forEach(ch => { ch.total = 0; ch.items.forEach(it => { it.imp = Math.round(it[4] * it[5] * 100) / 100; ch.total += it.imp; }); ch.total = Math.round(ch.total * 100) / 100; pem += ch.total; });
pem = Math.round(pem * 100) / 100;
const gg = Math.round(pem * 0.13 * 100) / 100;
const bi = Math.round(pem * 0.06 * 100) / 100;
const pec = Math.round((pem + gg + bi) * 100) / 100;
const iva = Math.round(pec * 0.21 * 100) / 100;
const totalIva = Math.round((pec + iva) * 100) / 100;
const indOffer = chapters.find(c => c.industrial).items.filter(i => i[6] === 'ind').reduce((a, i) => a + i.imp, 0);
const indTotal = indOffer;

// ---------- CONTENT ----------
const cover = [
  ...SP(6),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [R(TITOL, { bold: true, size: 32 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [R('Avinguda d’Eduard Maristany, 106-193. 08930 Sant Adrià de Besòs (Barcelona)', { size: 22 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [R('Projecte per a la sol·licitud de llicència d’obres davant l’Ajuntament de Sant Adrià de Besòs', { italics: true, size: 22 })] }),
  ...SP(4),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [R('Promotor: Consorci del Besòs', { size: 22 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [R('Tècnic redactor: Ignasi Tutzó Seró, Arquitecte Tècnic (col·legiat núm. 14646)', { size: 22 })] }),
  ...SP(2),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [R(DATA, { bold: true, size: 26 })] }),
  BR(),
];

const contingut = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [R(TITOL, { bold: true, size: 24 })] }),
  P('CONTINGUT DOCUMENTAL', { bold: true, after: 200 }),
  P('MEMÒRIA', { bold: true, after: 60 }),
  P('IN. ÍNDEX DE LA MEMÒRIA', { indent: { left: 567 }, after: 40 }),
  P('DD. DADES GENERALS', { indent: { left: 567 }, after: 40 }),
  P('MD. MEMÒRIA DESCRIPTIVA', { indent: { left: 567 }, after: 40 }),
  P('MJ. JUSTIFICACIÓ DEL MANTENIMENT DE LA CAPACITAT D’EVACUACIÓ', { indent: { left: 567 }, after: 40 }),
  P('MC. MEMÒRIA CONSTRUCTIVA', { indent: { left: 567 }, after: 40 }),
  P('MN. NORMATIVA APLICABLE', { indent: { left: 567 }, after: 160 }),
  P('DOCUMENTACIÓ GRÀFICA', { bold: true, after: 60 }),
  P('DG IN. ÍNDEX DE LA DOCUMENTACIÓ GRÀFICA', { indent: { left: 567 }, after: 160 }),
  P('PRESSUPOST', { bold: true, after: 60 }),
  P('PR1. AMIDAMENTS I PRESSUPOST · PR2. RESUM DE PRESSUPOST', { indent: { left: 567 }, after: 160 }),
  P('PLANIFICACIÓ TEMPORAL', { bold: true, after: 60 }),
  P('PT. PLANNING', { indent: { left: 567 }, after: 160 }),
  P('ANNEXOS', { bold: true, after: 60 }),
  P('1. Reportatge fotogràfic de l’estat actual', { indent: { left: 567 }, after: 40 }),
  P('2. Estudi bàsic de seguretat i salut', { indent: { left: 567 }, after: 40 }),
  P('3. Estudi de gestió de residus', { indent: { left: 567 }, after: 40 }),
  P('4. Pla de control de qualitat', { indent: { left: 567 }, after: 40 }),
  BR(),
];

const index = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [R('MEMÒRIA', { bold: true, size: 28 })] }),
  P('IN. ÍNDEX DE LA MEMÒRIA', { bold: true, after: 200 }),
  new TableOfContents('Índex', { hyperlink: true, headingStyleRange: '1-2' }),
  P([R('(Per actualitzar l’índex a Word: clic dret sobre l’índex › Actualitzar camp.)', { italics: true, size: 16, color: '777777' })]),
  BR(),
];

const dd = [
  H1('DD. DADES GENERALS'),
  H2('DD1. IDENTIFICACIÓ I OBJECTE DEL PROJECTE'),
  P([R('Títol del projecte: ', { bold: true }), R('Projecte de substitució de les portes i adaptació de la tanca d’accés al recinte de la Nau de Turbines de la Central Tèrmica del Besòs, des de l’Avinguda d’Eduard Maristany.')]),
  P([R('Objecte de l’encàrrec: ', { bold: true }), R('Definició tècnica i econòmica de les obres de substitució de la tanca i de les portes d’accés al recinte des de l’Avinguda d’Eduard Maristany, actualment malmeses per l’impacte de vehicles, amb l’objectiu de sol·licitar la corresponent llicència d’obres a l’Ajuntament de Sant Adrià de Besòs, mantenint com a mínim les capacitats d’evacuació previstes als projectes originals.')]),
  P([R('Situació: ', { bold: true }), R('Avinguda d’Eduard Maristany, 106-193, 08930 Sant Adrià de Besòs (Barcelona). Accés al recinte des de l’Avinguda d’Eduard Maristany (porta PX-01 del projecte executiu).')]),
  P([R('Referència cadastral: ', { bold: true }), R('6164301DF3866C0001MR (finca on se situa l’accés). Referència del conjunt segons projectes originals: 6164301DF3866C0001MR i 6267501DF3866E0001GX.')]),
  P([R('Àmbit: ', { bold: true }), R('Front de tanca de 17,67 m (límit d’ocupació directa, a comprovar segons el conveni urbanístic de gestió), amb una llum de pas entre paraments de 15,87 m, segons plànols adjunts.')]),

  H2('DD2. AGENTS DEL PROJECTE'),
  P([R('Promotor: ', { bold: true }), R('Consorci del Besòs, NIF P-5890053A. Tel. 934 626 868, correu electrònic info@consorcibesos.cat')]),
  P([R('Projectista i redactor del projecte: ', { bold: true }), R('Ignasi Tutzó Seró, Arquitecte Tècnic, col·legiat núm. 14646 del Col·legi de l’Arquitectura Tècnica de Barcelona (CAATEEB). NIF 47710532E. C/ Escorial, 113, 1r 2a, 08024 Barcelona. Tel. 627 601 224. Correu electrònic: itutzo@gmail.com')]),
  P('No intervé cap altre tècnic en la redacció d’aquest projecte. El mateix tècnic té encarregades la direcció d’obra i la coordinació de seguretat i salut.'),
  P([R('Titular dels terrenys: ', { bold: true }), R('Generalitat de Catalunya (acta d’ocupació directa de 13 de desembre de 2023), amb conveni urbanístic de gestió que faculta el Consorci del Besòs per actuar a l’àmbit.')]),
  P([R('Encàrrec: ', { bold: true }), R('Resolució del Consorci del Besòs de 15 de setembre de 2026 (expedient 178/26, codi 26XF0019), d’adjudicació del contracte menor de serveis de redacció de projecte, direcció d’obres i coordinació de seguretat i salut per a la renovació de la porta d’accés al recinte de la Nau de Turbines des de l’Av. d’Eduard Maristany.')]),

  H2('DD3. RELACIÓ DE DOCUMENTS COMPLEMENTARIS I DE REFERÈNCIA'),
  P('Formen part d’aquest projecte els documents següents:'), ...NEWLIST(),
  N('Estudi bàsic de seguretat i salut (Annex 2).'),
  N('Estudi de gestió de residus (Annex 3).'),
  N('Pla de control de qualitat (Annex 4).'),
  P('S’han pres com a documents de referència, i se’n mantenen els criteris en tot allò que no es modifica:', { before: 120 }), ...NEWLIST(),
  N('“Projecte executiu d’adaptació parcial de la Nau de Turbines de la Central Tèrmica del Besòs, pel projecte de Catalunya Media City. Esdeveniment inicial: Manifesta 15 (setembre a novembre 2024)”, Jordi Murtra i Ferré, arquitecte, desembre de 2023.'),
  N('“Projecte modificat del Projecte executiu d’adaptació parcial de la Nau de Turbines…”, Jordi Murtra i Ferré, arquitecte, maig de 2024.'),
  N('“Projecte as-built d’adaptació parcial de la Nau de Turbines…”, Batllori & Trepat Arquitectes SLP (direcció facultativa), maig de 2024. Plànols UR-04 a UR-07 i detall de la porta d’accés PX-01.'),
  N('Projecte tècnic de comunicació prèvia d’activitat (seguretat en cas d’incendi), Jordi Oste Díaz, enginyer tècnic industrial.'),
];

const md = [
  H1('MD. MEMÒRIA DESCRIPTIVA'),
  H2('MD1. INFORMACIÓ PRÈVIA: ANTECEDENTS I CONDICIONANTS DE PARTIDA'),
  P('L’any 2024 es van executar les obres del “Projecte executiu d’adaptació parcial de la Nau de Turbines de la Central Tèrmica del Besòs, pel projecte de Catalunya Media City. Esdeveniment inicial: Manifesta 15”, i del seu projecte modificat, que van adequar una part de la Nau de Turbines i el seu entorn exterior per acollir activitats esporàdiques. Aquests projectes van definir un recinte exterior delimitat perimetralment amb una tanca de malla de torsió simple, amb l’entrada principal des de l’Avinguda d’Eduard Maristany i una sortida d’emergència cap a la platja.'),
  P('L’accés des de l’Avinguda d’Eduard Maristany es resol amb la porta PX-01, definida al projecte executiu com a “porta d’emergència de dues fulles batents de 4×3 m de llum de pas, per cobrir l’amplada total de 17,67 m, d’acer galvanitzat en calent, amb bastidor de tub de 40×40×2 mm i malla de simple torsió de 50/14 mm de pas i 2,2 mm de gruix, muntants de tub de 80×80×2 mm”. Es va executar amb quatre portes de dues fulles batents sobre pilars metàl·lics fonamentats en daus de formigó.'),
  P('Actualment aquest tancament es troba malmès per l’impacte de vehicles (vegeu l’Annex 1, reportatge fotogràfic): el pòrtic superior d’acer que lliga els muntants està deformat; diverses fulles estan desplomades i fora de pla; els muntants centrals s’han hagut d’apuntalar provisionalment amb tornapuntes; i alguna fulla s’ha hagut de tancar amb cadena perquè els panys i forrellats no funcionen. L’accés està abalisat amb cinta, i el recinte no queda tancat amb garanties ni les portes es poden obrir i tancar amb normalitat. Les portes existents obren cap a l’exterior (sentit de l’evacuació) i tenen barres antipànic, que s’han de mantenir a la nova solució. A més, la configuració de quatre portes batents de gran format no permet el pas còmode de vehicles de manteniment i de servei, que és precisament l’origen dels impactes.'),
  P('Condicionants de partida:', { bold: true, keepNext: true }),
  B([R('Línia elèctrica soterrada. ', { bold: true }), R('Al llarg de la base de la tanca discorre una línia elèctrica soterrada, que no es pot afectar. Aquest condicionant determina la posició, la mida i el procediment d’execució de la nova fonamentació (vegeu MC.3 i MC.4).')]),
  B([R('Límit d’actuació. ', { bold: true }), R('L’actuació es limita al front de 17,67 m del límit d’ocupació directa (a comprovar segons el conveni), sense modificar l’alineació de la tanca actual ni ocupar la via pública.')]),
  B([R('Capacitat d’evacuació. ', { bold: true }), R('L’accés de l’Avinguda d’Eduard Maristany és un dels dos accessos del recinte computats a l’evacuació als projectes originals (amplada superior a 15 m). La nova solució ha de mantenir aquesta capacitat (vegeu l’apartat MJ).')]),
  B([R('Continuïtat del tancament del recinte. ', { bold: true }), R('Mentre durin les obres el recinte s’ha de mantenir tancat, amb una tanca provisional.')]),

  H2('MD2. DESCRIPCIÓ DEL PROJECTE'),
  H3('MD2.1. MOTIVACIÓ I CRITERIS DE LA PROPOSTA'),
  P('Es proposa substituir íntegrament el conjunt de portes existent per un nou tancament més robust, en la mateixa alineació i amb la mateixa llum total, que separa el pas de vianants del de vehicles:'),
  B([R('Quatre portes peatonals d’accés i evacuació ', { bold: true }), R('de dues fulles batents cadascuna, de 2,40 m de llum de pas lliure entre muntants i 3,00 m d’alçada, que obren cap a l’exterior (sentit de l’evacuació), com les actuals, i mantenen les '), R('barres antipànic segons UNE-EN 1125', { bold: true }), R(', amb comandament exterior amb clau.')]),
  B([R('Una porta corredissa de vehicles ', { bold: true }), R('d’una fulla, de 5,77 m de llum i 3,00 m d’alçada, motoritzada i amb dispositius de seguretat, que en obrir-se es desplaça cap al costat de l’edifici adjacent (costat Badalona), per darrere de la línia de façana, sense envair el pas de vianants ni la via pública.')]),
  B([R('Muntants d’acer galvanitzat de 100×200 mm ', { bold: true }), R('(en lloc dels muntants originals de 80×80 mm), amb més inèrcia davant d’impactes, sobre nous daus de formigó executats fora del traçat de la línia elèctrica soterrada.')]),
  P('La porta corredissa permet l’entrada de vehicles de servei per un pas únic i ample, sense haver d’obrir i maniobrar les fulles batents, que és la causa principal dels danys actuals.'),

  H3('MD2.2. JUSTIFICACIÓ DEL COMPLIMENT DE LA NORMATIVA URBANÍSTICA I MUNICIPAL'),
  P('L’actuació és una substitució d’un element de tancament existent per un altre de característiques i dimensions anàlogues. No altera l’alineació, l’alçada (3,00 m, igual que l’existent), l’ocupació, l’edificabilitat ni l’ús del recinte. És compatible amb el planejament vigent (PDU de les Tres Xemeneies) i amb el règim d’activitats esporàdiques amb què es van tramitar els projectes originals.'),
  P('Les fulles batents de les portes peatonals obren cap a l’exterior, igual que les actuals, sobre l’espai d’accés davant de les portes (≈ 3,80 m de fondària fins a l’alineació de l’avinguda), i la porta corredissa llisca per la cara interior, darrere de la línia de façana de l’edifici adjacent. Es manté, per tant, la mateixa ocupació de l’espai d’accés que la de les portes existents.'),
  P('Es tramita com a llicència d’obres (Decret legislatiu 1/2010, Text refós de la Llei d’urbanisme de Catalunya, i ordenances municipals de Sant Adrià de Besòs), atès que inclou nova fonamentació i la substitució d’un element de tancament que dona a la via pública.'),

  H3('MD2.3. SEGURETAT D’UTILITZACIÓ I ACCESSIBILITAT (CTE DB SUA)'),
  B([R('SUA 1. ', { bold: true }), R('El pas es manté al mateix nivell que el paviment existent, sense graons ni ressalts. La guia inferior de la porta corredissa (Ø 20 mm) se situa fora dels passos de vianants. Els passadors inferiors de les barres antipànic de dos punts entren en un encaix enrasat amb el paviment, sense sobresortir.')]),
  B([R('SUA 2. ', { bold: true }), R('Les fulles batents obertes no envaeixen cap itinerari. La porta corredissa és motoritzada i compleix UNE-EN 13241, UNE-EN 12453 i UNE-EN 12445: fotocèl·lula, banda de seguretat sensible, avisador òptic, finals de cursa i desbloqueig manual. Els muntants se senyalitzen amb bandes reflectants contra els impactes de vehicles.')]),
  B([R('SUA 9 i Codi d’accessibilitat de Catalunya (Decret 209/2023). ', { bold: true }), R('Cada porta peatonal té 2,40 m de llum (cada fulla ≈ 1,20 m), molt per sobre del mínim d’0,80 m per a itineraris accessibles. Es manté l’itinerari accessible existent fins al nucli 1, sense desnivells, i a les dues bandes de les portes queda un espai lliure de Ø 1,50 m.')]),

  H3('MD2.4. RELACIÓ DE SUPERFÍCIES I DIMENSIONS'),
  table([4600, 2200, 2200], [
    ['Element', 'Estat actual', 'Proposta'],
    ['Front de tanca (límit d’ocupació directa)', '17,67 m', '17,67 m'],
    ['Llum entre paraments', '15,87 m', '15,87 m'],
    ['Alçada de tancament', '3,00 m', '3,00 m'],
    ['Portes peatonals', '4 u de dues fulles (≈ 4,00 m)', '4 u de dues fulles (2,40 m de llum lliure)'],
    ['Porta de vehicles', '—', '1 u corredissa motoritzada de 5,77 m'],
    ['Sentit d’obertura portes peatonals', 'Cap a l’exterior (evacuació)', 'Cap a l’exterior (evacuació)'],
    ['Barres antipànic (UNE-EN 1125)', 'Sí', 'Sí, a les 4 portes (es mantenen)'],
    ['Muntants / pilars', 'Tub 80×80 mm', 'Perfil d’acer 100×200 mm'],
    ['Superfície de tancament substituïda', '≈ 47,6 m2', '≈ 47,6 m2'],
  ]),
  P('', { after: 60 }),
];

const mj = [
  H1('MJ. JUSTIFICACIÓ DEL MANTENIMENT DE LA CAPACITAT D’EVACUACIÓ'),
  H2('MJ1. PREVISIÓ DELS PROJECTES ORIGINALS'),
  P('Els projectes de referència (projecte executiu i projecte modificat) estableixen per al recinte exterior els criteris d’evacuació següents (apartats MD2.2 i SUA 5 del projecte modificat):'),
  B('Recinte exterior a l’aire lliure, delimitat perimetralment, amb entrada per l’Avinguda d’Eduard Maristany i sortida d’emergència per la platja, totes dues d’amplada superior a 15 m.'),
  B('Capacitat d’evacuació considerada: 600 persones per metre lineal (CTE DB SI 3, taula 4.1, passos en zones a l’aire lliure: A ≥ P/600).'),
  B(`Aforament interior de l’edifici: 3.610,68 m2 de superfície ocupable a 2 m2/persona = ${Pint.toLocaleString('ca-ES')} persones.`),
  B('Possibilitat d’una ocupació exterior similar, no simultània (cues de control d’accés, serveis higiènics, altres serveis).'),
  B('Els accessos es poden tancar a la nit, fora de l’horari de l’activitat; durant l’activitat romanen oberts.'),
  P('L’activitat s’empara en la ITC SP 144:2023 de Bombers de la Generalitat de Catalunya (condicions de seguretat en cas d’incendi en activitats esporàdiques). La porta PX-01 del projecte executiu es va definir amb quatre unitats de dues fulles de 4,00 m de llum nominal.'),

  H2('MJ2. AMPLADES DE PAS: ESTAT ACTUAL I PROPOSTA'),
  P(`La llum de pas útil es calcula descomptant de la llum total entre paraments (15,87 m) l’amplada dels muntants que hi queden inclosos. A la proposta hi ha ${nPost} muntants de 100×200 mm, col·locats amb la cara de 0,10 m en el sentit de la tanca (4 de les portes peatonals i el de recepció de la corredissa). Cada mòdul de dues portes té 2,40 m de llum lliure (2,50 m entre eixos). Comprovació: 4 × 2,40 + 5 × 0,10 + 5,77 = 15,87 m.`),
  table([3900, 2550, 2550], [
    ['Concepte', 'Projecte executiu / estat actual', 'Proposta'],
    ['Llum total entre paraments', '15,87 m', '15,87 m'],
    ['Portes peatonals', '4 × 2 fulles batents', `4 × 2 fulles batents (4 × ${num(pedEach)} m = ${num(pedTotal)} m)`],
    ['Porta de vehicles (oberta durant l’activitat)', '—', `${num(vehClear)} m`],
    ['Llum de pas útil total', '> 15,00 m (criteri del projecte)', `${num(clearTotal)} m`],
    ['Capacitat a P/600 (aire lliure)', '> 9.000 persones', `${Math.floor(clearTotal * 600).toLocaleString('ca-ES')} persones`],
  ]),
  P('', { after: 60 }),

  H2('MJ3. COMPROVACIÓ DE LA CAPACITAT D’EVACUACIÓ (CTE DB SI 3)'),
  P('S’hi fan tres comprovacions, de menys a més exigent:'),
  P('a) Criteri dels projectes originals (pas a l’aire lliure, A ≥ P/600):', { bold: true, keepNext: true }),
  P(`Amb una llum útil de ${num(clearTotal)} m, la capacitat és de ${Math.floor(clearTotal * 600).toLocaleString('ca-ES')} persones. Supera l’amplada mínima de 15 m que fixen els projectes originals i, per tant, la capacitat de 9.000 persones que s’hi va considerar. Encara que se suposin simultanis l’aforament interior i l’exterior (P = ${Pmax.toLocaleString('ca-ES')} persones), l’amplada necessària seria de ${num(Pmax / 600)} m, molt per sota de l’existent.`),
  P('b) Hipòtesi de bloqueig de la sortida a la platja (SI 3, apartat 4.1):', { bold: true, keepNext: true }),
  P(`Si se suposa inutilitzada la sortida d’emergència a la platja, tota l’ocupació (P = ${Pmax.toLocaleString('ca-ES')} persones, hipòtesi conservadora) evacua per l’Avinguda d’Eduard Maristany. Amplada necessària: A ≥ ${Pmax.toLocaleString('ca-ES')} / 600 = ${num(Pmax / 600)} m ≤ ${num(clearTotal)} m. COMPLEIX.`),
  P('c) Comprovació addicional amb el criteri de portes (A ≥ P/200), només amb les portes peatonals:', { bold: true, keepNext: true }),
  P(`Fins i tot sense comptar la porta de vehicles i aplicant el criteri més restrictiu de portes i passos (A ≥ P/200), les quatre portes peatonals (${num(pedTotal)} m) donen una capacitat de ${Math.floor(pedTotal * 200).toLocaleString('ca-ES')} persones, superior a l’aforament interior de l’edifici (${Pint.toLocaleString('ca-ES')} persones). Amb el conjunt de la llum útil, la capacitat amb aquest criteri és de ${Math.floor(clearTotal * 200).toLocaleString('ca-ES')} persones.`),
  table([4300, 1600, 1600, 1500], [
    ['Hipòtesi', 'P (persones)', 'A necessària', 'A disponible'],
    ['a) Aforament interior, P/600', Pint.toLocaleString('ca-ES'), `${num(Pint / 600)} m`, `${num(clearTotal)} m`],
    ['b) Interior + exterior simultanis, P/600', Pmax.toLocaleString('ca-ES'), `${num(Pmax / 600)} m`, `${num(clearTotal)} m`],
    ['c) Aforament interior, P/200 (només portes peatonals)', Pint.toLocaleString('ca-ES'), `${num(Pint / 200)} m`, `${num(pedTotal)} m`],
  ]),
  P('', { after: 60 }),

  H2('MJ4. CONDICIONS DE LES PORTES EN RELACIÓ AMB L’EVACUACIÓ (CTE DB SI 3, APARTAT 6)'),
  P('Les quatre portes peatonals es dissenyen com a portes de sortida d’evacuació a l’aire lliure, de manera que la capacitat d’evacuació no depèn de deixar-les obertes:'),
  B([R('Sentit d’obertura: ', { bold: true }), R('obren en el sentit de l’evacuació, cap a l’Avinguda d’Eduard Maristany, igual que les portes existents, tal com exigeix SI 3-6.3 per a portes previstes per a més de 100 persones.')]),
  B([R('Sistema d’obertura: ', { bold: true }), R('són abatibles amb eix de gir vertical i porten dispositiu antipànic de barra horitzontal d’empenta segons UNE-EN 1125 a totes dues fulles, igual que les portes existents, com exigeix SI 3-6.1 per a l’evacuació de més de 50 persones no familiaritzades amb la porta. Des de l’interior s’obren sempre, sense clau, encara que estiguin tancades. Des de l’exterior s’obren amb maneta i clau. Els dispositius antipànic es valoren en una partida específica del pressupost (05.03).')]),
  B([R('Amplada de les fulles: ', { bold: true }), R('cada fulla ha de tenir una amplada de pas entre 0,60 i 1,23 m (SI 3-4.2). Amb 2,40 m de llum per porta, cada fulla fa ≈ 1,20 m. COMPLEIX.')]),
  B([R('Porta corredissa: ', { bold: true }), R('per ser corredissa i motoritzada, no es computa com a porta d’evacuació (comprovació c). Durant l’horari de l’activitat romandrà oberta i bloquejada, i aleshores funciona com un pas lliure addicional (comprovacions a i b). Té desbloqueig manual en cas de tall de subministrament elèctric.')]),
  B('Fora de l’horari de l’activitat les portes es poden tancar amb clau sense perdre la capacitat d’evacuació, perquè l’antipànic permet sortir des de l’interior en qualsevol moment.'),
  B('Es mantenen la senyalització de sortida i l’enllumenat d’emergència existents a l’accés (SI 3.7 i SUA 4), i s’afegeix el rètol “Sortida d’emergència” a cada porta peatonal.'),
  P([R('Conclusió: ', { bold: true }), R(`la nova configuració manté una llum útil de pas de ${num(clearTotal)} m, superior als 15 m previstos als projectes originals, i una capacitat d’evacuació de ${Math.floor(clearTotal * 600).toLocaleString('ca-ES')} persones amb el mateix criteri (P/600). Com que es manté la mateixa llum total entre paraments, el mateix nombre de portes peatonals i es mantenen el sentit d’obertura cap a l’exterior i les barres antipànic de les portes actuals, `), R('no es redueix la capacitat d’evacuació prevista als projectes originals.', { bold: true })]),
];

const mc = [
  H1('MC. MEMÒRIA CONSTRUCTIVA'),
  H2('MC1. TREBALLS PREVIS'),
  P('Abans de qualsevol enderroc s’instal·larà una tanca provisional mòbil d’obra a la cara interior del front d’actuació, per mantenir el recinte tancat durant tota l’obra, i s’abalisarà i senyalitzarà la vorera de l’Avinguda d’Eduard Maristany per protegir els vianants.'),
  H2('MC2. PROTECCIÓ DE LA LÍNIA ELÈCTRICA SOTERRADA'),
  P('Al llarg de la base de la tanca hi ha una línia elèctrica soterrada que no es pot afectar. Per això s’estableix el procediment següent, d’obligat compliment per al contractista:'), ...NEWLIST(),
  N('Sol·licitud prèvia a la companyia distribuïdora i a la propietat dels plànols de serveis afectats, i comunicació de l’inici dels treballs.'),
  N('Detecció i marcatge en superfície del traçat i de la fondària de la línia amb detector de cables i georadar, per empresa especialitzada, abans de replantejar la fonamentació.'),
  N('Cala manual de comprovació a cada posició de muntant nou, abans d’excavar.'),
  N('Replanteig dels nous daus fora del traçat detectat, amb una distància lliure mínima de 0,30 m entre la cara del dau i el tub o cable. Si la posició teòrica d’un muntant coincideix amb la línia, el dau es desplaçarà en sentit perpendicular a la tanca (cap a l’interior del recinte) i el muntant s’hi ancorarà amb una placa excèntrica, amb validació prèvia de la direcció facultativa.'),
  N('Excavació i enderroc exclusivament amb mitjans manuals i eines elèctriques lleugeres a menys d’1,00 m de la línia. Queda prohibit l’ús de retroexcavadores i de martells pneumàtics pesants en aquesta franja.'),
  N('La porta corredissa no requereix rasa. La guia inferior de la porta (Ø 20 mm soldada sobre platines) no s’encastarà al terra: les platines s’ancoraran al paviment existent amb tacs químics de longitud limitada (≤ 100 mm). Si cal formigonar una franja de suport, no farà més de 0,20 m de fondària i només es farà després de la detecció i amb l’autorització de la direcció facultativa. Els daus es concentren només als muntants i a la torreta.'),
  N('Si s’afecta accidentalment la línia, s’aturaran els treballs, s’abalisarà la zona i s’avisarà immediatament la companyia distribuïdora.'),
  H2('MC3. ENDERROCS I DESMUNTATGES'),
  P('Es desmuntaran les quatre portes de dues fulles existents i els seus pilars, amb mitjans manuals i camió grua. Es tallaran les frontisses i els panys i es carregaran els elements sobre camió per portar-los a un gestor de residus metàl·lics. Els daus de formigó existents s’enderrocaran amb mitjans manuals; si algun és a menys de 0,30 m de la línia elèctrica, es tallarà el pilar a ran del dau i aquest es deixarà soterrat, sense enderrocar-lo. Es retirarà el paviment de panot de l’entorn dels nous muntants.'),
  H2('MC4. FONAMENTACIÓ'),
  P('Els nous muntants es fonamenten en daus aïllats de formigó en massa HA-25/B/20/IIa de 0,60×0,60×0,80 m, sobre una capa de neteja de 5 cm, amb una placa d’ancoratge de 250×350×15 mm per al muntant de 100×200 mm, i quatre perns Ø16 mm de B500S, galvanitzats. Es preveuen 7 daus: 5 per als muntants de les portes peatonals (un dels quals també és el de recepció de la porta corredissa), i 2 per a la torreta de suport i el motor de la corredissa.'),
  P('El dimensionament s’ha fet considerant l’acció del vent sobre la reixa (CTE DB SE-AE, zona eòlica C, entorn costaner, grau d’aspror I), el pes propi de les fulles i una acció accidental d’ús. '),
  H2('MC5. TANCAMENTS PRACTICABLES: PORTES'),
  B([R('Muntants: ', { bold: true }), R('perfil d’acer de 100×200 mm galvanitzat en calent (UNE-EN ISO 1461), col·locat amb la cara de 100 mm en el sentit de la tanca, cargolat a la placa d’ancoratge del dau.')]),
  B([R('Portes peatonals (4 u), de dues fulles batents de 2.500×3.000 mm (2,40 m de llum lliure): ', { bold: true }), R('fulla d’armadura tubular d’acer amb travessers rigiditzadors, folrada amb perfils d’acer galvanitzat; gir amb cassoleta superior i pivot inferior amb rodaments de boles, allotjat al muntant lateral; acabat galvanitzat. Obertura cap a l’exterior, com les existents. Tipus JAM o equivalent.')]),
  B([R('Dispositius antipànic (4 jocs): ', { bold: true }), R('barra horitzontal d’empenta UNE-EN 1125 amb pany de cop a la fulla activa; barra de dos punts (tancament superior i inferior) a la fulla passiva; comandament exterior amb maneta i cilindre de clau; acabat per a exterior. Marcat CE.')]),
  B([R('Porta corredissa de vehicles (1 u), de 5.770×3.000 mm: ', { bold: true }), R('fulla folrada amb perfils d’acer galvanitzat laminats en fred, rodes d’acer amb rodaments, guia inferior calibrada de Ø 20 mm, torreta de suport amb rodaments de niló i perfil en “U” de tancament. Motor electromecànic de 230 V per a fulles de fins a 1.100 kg, amb maniobra incorporada, fotocèl·lula, banda de seguretat, avisador òptic, finals de cursa i desbloqueig manual. Marcat CE i expedient tècnic segons UNE-EN 13241, UNE-EN 12453 i UNE-EN 12445. Tipus JAM o equivalent.')]),
  B([R('Senyalització: ', { bold: true }), R('bandes reflectants als muntants contra els impactes de vehicles i rètol fotoluminescent de “Sortida d’emergència” a cada porta peatonal.')]),
  P('Les característiques i els preus de les portes s’han definit a partir d’una consulta de mercat a un fabricant especialitzat. Les referències a models comercials són orientatives: el contractista podrà oferir productes equivalents que compleixin les prestacions d’aquest projecte (dimensions, llum de pas, sentit d’obertura, antipànic, marcat CE i acabat galvanitzat).'),
  H2('MC6. PAVIMENTS I REPOSICIONS'),
  P('Es reposarà el paviment a l’entorn dels muntants amb panot igual a l’existent (o formigó raspallat), enrasat amb el paviment adjacent i sense ressalts, i se segellarà el contorn dels muntants amb massilla de poliuretà.'),
  H2('MC7. INSTAL·LACIONS'),
  P('El motor de la porta corredissa necessita alimentació monofàsica de 230 V (380 W), que es valora al capítol 09. S’alimentarà des del quadre elèctric existent a l’extrem esquerre de la tanca (element 7 dels plànols), amb una línia nova de cable RZ1-K 3G2,5 mm² protegida al quadre amb magnetotèrmic de 10 A i diferencial de 30 mA. Per arribar al motor, a l’extrem dret, la línia es passarà pel conducte soterrat existent Ø 50 mm (element 8), que ja creua tot el front de la tanca, fent-hi passar només el nou cable amb guia i sense manipular ni desconnectar els cables existents, de manera que no cal obrir cap rasa a la zona de la línia elèctrica soterrada. Es farà d’acord amb el REBT i es lliurarà el butlletí corresponent.'),
  P('No es toquen la línia elèctrica soterrada ni el quadre de telecomunicacions existent a l’extrem de la tanca (element 7 dels plànols), i es mantenen l’enllumenat d’emergència i la senyalització de l’accés.'),
];

const mn = [
  H1('MN. NORMATIVA APLICABLE'),
  H2('MN1. EDIFICACIÓ'),
  B('Codi Tècnic de l’Edificació (RD 314/2006 i modificacions), en allò que sigui d’aplicació: DB SE, DB SE-AE, DB SE-A, DB SE-C, DB SI (SI 3, evacuació d’ocupants) i DB SUA.'),
  B('Codi Estructural (RD 470/2021).'),
  B('Llei 38/1999 d’ordenació de l’edificació (LOE).'),
  B('Decret 209/2023, pel qual s’aprova el Codi d’accessibilitat de Catalunya, i Llei 13/2014 d’accessibilitat.'),
  H2('MN2. URBANISME I LLICÈNCIES'),
  B('Decret legislatiu 1/2010, Text refós de la Llei d’urbanisme de Catalunya, i Decret 64/2014, Reglament sobre protecció de la legalitat urbanística.'),
  B('Pla director urbanístic de les Tres Xemeneies i planejament derivat vigent.'),
  B('Ordenances municipals de l’Ajuntament de Sant Adrià de Besòs (obres i ocupació de la via pública).'),
  H2('MN3. ALTRES'),
  B('ITC SP 144:2023 de Bombers de la Generalitat de Catalunya. Condicions de seguretat en cas d’incendi en activitats esporàdiques.'),
  B('Reglament electrotècnic per a baixa tensió (RD 842/2002), ITC-BT-07 (xarxes subterrànies), i RD 223/2008 (línies elèctriques d’alta tensió), en relació amb la protecció de la línia soterrada.'),
  B('UNE-EN 13241 (portes industrials, comercials i de garatge), UNE-EN 12604, UNE-EN 12453 i UNE-EN 12445 (seguretat d’ús de portes motoritzades).'),
  B('UNE-EN 1125 (dispositius antipànic per a sortides d’emergència, accionats per barra horitzontal).'),
  B('UNE-EN ISO 1461 (recobriments galvanitzats en calent).'),
  B('RD 1627/1997, disposicions mínimes de seguretat i salut en obres de construcció, i Llei 31/1995 de prevenció de riscos laborals.'),
  B('RD 105/2008, producció i gestió de residus de construcció i demolició; Decret 89/2010 (Catalunya); Llei 7/2022 de residus i sòls contaminats.'),
  BR(),
];

// ---------- Documentació gràfica ----------
const plans = [
  ['0', 'Emplaçament', 'E 1:5000', 'pl-1.jpg'],
  ['01', 'Estat actual. Planta i alçat', 'E 1:100 (A3)', 'pl-2.jpg'],
  ['02', 'Proposta tanca. Planta i alçat', 'E 1:100 (A3)', 'pl-3.jpg'],
  ['03', 'Enderroc i obra nova. Planta i alçat', 'E 1:100 (A3)', 'pl-4.jpg'],
];
const dgIndex = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [R('DOCUMENTACIÓ GRÀFICA', { bold: true, size: 28 })] }),
  H1('DG IN. ÍNDEX DE LA DOCUMENTACIÓ GRÀFICA'),
  table([1200, 5300, 2500], [['Núm.', 'Plànol', 'Escala'], ...plans.map(p => [p[0], p[1], p[2]])]),
  P('', { after: 120 }),
  P('Els plànols s’adjunten a continuació, a escala, en format A3.'),
];
const planPages = plans.map((p, i) => [
  new Paragraph({ alignment: AlignmentType.CENTER, children: [img(p[3], 880, 622)], pageBreakBefore: i > 0 }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [R(`Plànol ${p[0]} – ${p[1]} (${p[2]})`, { size: 18, italics: true })] }),
]).flat();

// ---------- Pressupost ----------
const prRows = [['Codi', 'Ud', 'Descripció', 'Amidament', 'Quant.', 'Preu (€)', 'Import (€)']];
chapters.forEach(ch => {
  prRows.push([{ t: ch.c, b: true, fill: 'EEF2F6' }, { t: '', fill: 'EEF2F6' }, { t: ch.t, b: true, fill: 'EEF2F6' }, { t: '', fill: 'EEF2F6' }, { t: '', fill: 'EEF2F6' }, { t: '', fill: 'EEF2F6' }, { t: '', fill: 'EEF2F6' }]);
  ch.items.forEach(it => {
    const hl = it[6] === 'est' ? 'yellow' : undefined;
    prRows.push([it[0], it[1], it[2], it[3], { t: num(it[4]), a: AlignmentType.RIGHT }, { t: eur(it[5]), a: AlignmentType.RIGHT, hl }, { t: eur(it.imp), a: AlignmentType.RIGHT, hl }]);
  });
  prRows.push(['', '', { t: `TOTAL CAPÍTOL ${ch.c}`, b: true }, '', '', '', { t: eur(ch.total), b: true, a: AlignmentType.RIGHT }]);
});
const pressupost = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [R('PRESSUPOST', { bold: true, size: 28 })] }),
  H1('PR1. AMIDAMENTS I PRESSUPOST'),
  P('Els preus de les portes (partides 05.01 i 05.02) s’han obtingut d’una consulta de mercat a un fabricant especialitzat de portes metàl·liques. La resta de partides s’han valorat amb preus de referència de mercat (base ITeC 2026). L’obra s’adjudicarà per licitació.'),
  table([650, 450, 3700, 1950, 700, 850, 1000], prRows, { fontSize: 15 }),
  BR(),
  H1('PR2. RESUM DE PRESSUPOST'),
  table([1000, 5600, 2400], [
    ['Cap.', 'Descripció', 'Import (€)'],
    ...chapters.map(ch => [ch.c, ch.t, { t: eur(ch.total), a: AlignmentType.RIGHT, }]),
    ['', { t: 'PRESSUPOST D’EXECUCIÓ MATERIAL (PEM)', b: true }, { t: eur(pem), b: true, a: AlignmentType.RIGHT }],
    ['', '13 % Despeses generals', { t: eur(gg), a: AlignmentType.RIGHT }],
    ['', '6 % Benefici industrial', { t: eur(bi), a: AlignmentType.RIGHT }],
    ['', { t: 'PRESSUPOST D’EXECUCIÓ PER CONTRACTE (PEC)', b: true }, { t: eur(pec), b: true, a: AlignmentType.RIGHT }],
    ['', '21 % IVA', { t: eur(iva), a: AlignmentType.RIGHT }],
    ['', { t: 'PRESSUPOST TOTAL (IVA inclòs)', b: true }, { t: eur(totalIva), b: true, a: AlignmentType.RIGHT }],
  ]),
  P('', { after: 120 }),
  P(`El pressupost d’execució material (base imposable de l’ICIO) puja a ${eur(pem)} €.`, { bold: true }),
  BR(),
];

// ---------- Planning ----------
const weeks = ['S1', 'S2', 'S3', 'S4', 'S5', 'S6'];
const tasks = [
  ['Fabricació a taller de portes', [1, 1, 1, 1, 0, 0]],
  ['Detecció de serveis, cales i tanca provisional', [0, 0, 0, 1, 0, 0]],
  ['Desmuntatge de portes i enderrocs', [0, 0, 0, 1, 0, 0]],
  ['Excavació i formigonat de daus (curat)', [0, 0, 0, 1, 1, 0]],
  ['Muntatge de portes, kit antipànic i guia', [0, 0, 0, 0, 1, 1]],
  ['Alimentació elèctrica, motor i posada en marxa', [0, 0, 0, 0, 0, 1]],
  ['Reposició de paviments i acabats', [0, 0, 0, 0, 0, 1]],
  ['Residus, retirada de la tanca provisional i neteja', [0, 0, 0, 0, 0, 1]],
];
const planning = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [R('PLANIFICACIÓ TEMPORAL', { bold: true, size: 28 })] }),
  H1('PT. PLANNING'),
  P('Termini d’execució previst: 6 setmanes des de l’adjudicació, de les quals unes 4 corresponen a la fabricació a taller (termini a confirmar per l’adjudicatari). La durada de les obres in situ, amb el recinte tancat provisionalment, és d’unes 3 setmanes.'),
  table([4200, ...weeks.map(() => 800)], [
    ['Activitat', ...weeks],
    ...tasks.map(t => [t[0], ...t[1].map(v => ({ t: v ? '■' : '', a: AlignmentType.CENTER, fill: v ? '9DB4CC' : undefined }))]),
  ]),
  BR(),
];

// ---------- Annexos ----------
const photo = (file, w, h) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 }, children: [img(file, w, h)] });
const caption = (t) => P(t, { italics: true, align: AlignmentType.CENTER, after: 200 });
const annex1 = [
  H1('ANNEX 1. REPORTATGE FOTOGRÀFIC DE L’ESTAT ACTUAL'),
  P('Imatge de situació de l’accés (Google Street View, 2024) i fotografies i fotogrames del vídeo de l’estat actual de les portes d’accés al recinte des de l’Avinguda d’Eduard Maristany (juliol de 2026).'),
  photo('foto0.jpg', 480, 360),
  caption('Fotografia 1. Vista de l’accés al recinte des de l’Avinguda d’Eduard Maristany, amb les Tres Xemeneies al fons. Font: Google Maps (Street View), imatge de 2024, anterior als danys actuals.'),
  photo('foto1.jpg', 480, 360),
  caption('Fotografia 2. Vista general del conjunt de quatre portes de dues fulles. El pòrtic superior està deformat, les fulles centrals estan desplomades, els muntants estan apuntalats provisionalment amb tornapuntes i l’accés està abalisat amb cinta. Es veuen les barres antipànic de les portes existents.'),
  photo('foto2.jpg', 480, 360),
  caption('Fotografia 3. Vista obliqua de les portes, amb la deformació del pòrtic superior i de les fulles centrals per l’impacte de vehicles. A la dreta, el pilar de fàbrica de maó de l’extrem de la tanca.'),
  photo('frames.jpg', 470, 410),
  caption('Fotografies 4 i 5 (fotogrames del vídeo Acces_NdT.mp4). Esquerra: fulla tancada provisionalment amb cadena, barra antipànic existent i muntant desplomat amb tornapunta. Dreta: pòrtic superior deformat i fulles fora de pla.'),
  BR(),
];
const annex2 = [
  H1('ANNEX 2. ESTUDI BÀSIC DE SEGURETAT I SALUT'),
  P('D’acord amb l’article 4 del RD 1627/1997, cal un estudi bàsic de seguretat i salut perquè el pressupost és inferior a 450.759,08 €, la durada és inferior a 30 dies laborables amb menys de 20 treballadors simultanis, el volum de mà d’obra és inferior a 500 jornades i no es tracta d’obres de túnels, galeries, conduccions subterrànies ni preses.'),
  H2('A2.1. DADES DE L’OBRA'),
  B(`Pressupost d’execució material: ${eur(pem)} €.`),
  B('Termini d’execució in situ: 3 setmanes. Nombre màxim de treballadors simultanis: 4.'),
  B('Accés a l’obra des de l’Avinguda d’Eduard Maristany, amb ocupació puntual de la vorera per a càrrega i descàrrega.'),
  H2('A2.2. RISCOS I MESURES PREVENTIVES'),
  table([3000, 6000], [
    ['Risc', 'Mesures preventives'],
    ['Contacte elèctric amb la línia soterrada', 'Detecció prèvia amb detector i georadar, cales manuals, excavació manual a menys d’1 m, abalisament del traçat, eines aïllades, prohibició de maquinària d’excavació a la franja de protecció. Protocol d’aturada i avís a la companyia.'],
    ['Atropellaments i interferència amb el trànsit', 'Abalisament de la vorera, desviament de vianants senyalitzat, senyalista durant les maniobres del camió grua.'],
    ['Caiguda d’objectes en manipulació (fulles de porta)', 'Camió grua amb eslingues certificades, prohibició de romandre sota càrregues suspeses, apuntalament de les fulles fins a fixar-les.'],
    ['Cops i talls; projecció de partícules', 'Guants, ulleres, calçat de seguretat, protecció de radials, ús de pantalla en talls i soldadures.'],
    ['Soroll i vibracions', 'Protectors auditius; limitació de l’ús del martell elèctric.'],
    ['Incendi en treballs de tall i soldadura', 'Extintor a peu d’obra, retirada de materials combustibles, permís de treballs en calent.'],
    ['Sobreesforços', 'Manipulació mecànica d’elements de més de 25 kg; treball en equip.'],
    ['Contacte elèctric en la connexió del motor', 'Treballs sense tensió, amb consignació del circuit al quadre i comprovació d’absència de tensió; per instal·lador autoritzat.'],
  ], { fontSize: 17 }),
  P('', { after: 60 }),
  P('Equips de protecció individual: casc, calçat de seguretat, guants, ulleres, armilla d’alta visibilitat, protectors auditius i guants dielèctrics per a les cales.'),
  P('El contractista elaborarà el Pla de seguretat i salut, que haurà d’aprovar el coordinador de seguretat i salut abans de començar l’obra, i farà la comunicació d’obertura del centre de treball.'),
  BR(),
];
const annex3 = [
  H1('ANNEX 3. ESTUDI DE GESTIÓ DE RESIDUS'),
  P('Redactat d’acord amb el RD 105/2008 i el Decret 89/2010.'),
  table([3600, 1500, 1300, 2600], [
    ['Residu', 'Codi LER', 'Quantitat', 'Destí'],
    ['Metalls (acer de portes i pilars)', '17 04 05', '1,20 t', 'Valorització (gestor de ferralla)'],
    ['Formigó i paviment', '17 01 01', '1,89 m3', 'Planta de reciclatge de runes'],
    ['Terres d’excavació no especials', '17 05 04', '2,53 m3', 'Dipòsit controlat / reutilització'],
    ['Envasos i embalatges', '15 01 xx', '< 0,1 t', 'Recollida selectiva'],
  ], { fontSize: 17 }),
  P('', { after: 60 }),
  P('Es farà la separació en origen dels residus metàl·lics. La resta es recollirà en contenidor. Com que la quantitat és petita, no s’arriba als llindars de separació obligatòria per fraccions de l’article 5.5 del RD 105/2008, tret dels metalls, que se separen per valoritzar-los.'),
  P('El cost de la gestió de residus és el del capítol 07 del pressupost. Els justificants de lliurament a gestor autoritzat s’aportaran a l’Ajuntament per retornar la fiança.'),
  P('La fiança de gestió de residus la tramitarà qui faci la tramitació de la llicència, d’acord amb l’ordenança municipal de Sant Adrià de Besòs. No forma part de l’abast d’aquest projecte.'),
  BR(),
];
const annex4 = [
  H1('ANNEX 4. PLA DE CONTROL DE QUALITAT'),
  table([3200, 5800], [
    ['Material / unitat', 'Control'],
    ['Formigó HA-25/B/20/IIa', 'Control documental (full de subministrament i distintiu). Assaig de consistència (con d’Abrams) en obra a cada amassada. Control estadístic segons el Codi Estructural si ho requereix la direcció facultativa.'],
    ['Acer de perns B500S i plaques S275JR', 'Certificats de qualitat i marcat CE.'],
    ['Perfils i portes d’acer galvanitzat', 'Declaració de prestacions i marcat CE (UNE-EN 13241 a la porta corredissa). Certificat de galvanitzat en calent (UNE-EN ISO 1461) i mesura del gruix de recobriment en obra (≥ 85 µm).'],
    ['Muntatge', 'Comprovació d’aplomat (≤ 1/500), anivellament i llums de pas (≥ 2,40 m per porta peatonal i ≥ 15,00 m de llum útil total).'],
    ['Dispositius antipànic', 'Certificat UNE-EN 1125 i marcat CE. Prova d’obertura de cada porta des de l’interior amb la porta tancada amb clau: força d’accionament ≤ 80 N i obertura de les dues fulles.'],
    ['Porta corredissa motoritzada', 'Declaració CE, expedient tècnic, manual d’ús i llibre de manteniment (UNE-EN 13241). Assaig de forces segons UNE-EN 12445 i prova de fotocèl·lula, banda de seguretat i desbloqueig manual.'],
    ['Instal·lació elèctrica', 'Verificació del diferencial de 30 mA i lliurament del butlletí.'],
    ['Protecció de la línia soterrada', 'Informe de detecció de serveis abans d’excavar i acta de replanteig dels daus signada per la direcció facultativa.'],
  ], { fontSize: 17 }),
  BR(),
];

const signatura = [
  P('', { after: 400 }),
  P(`Sant Adrià de Besòs, ${DATA.toLowerCase()}`, { align: AlignmentType.RIGHT }),
  P('', { after: 600 }),
  P('Ignasi Tutzó Seró', { align: AlignmentType.RIGHT, bold: true, after: 0 }),
  P('Arquitecte Tècnic. Col·legiat núm. 14646', { align: AlignmentType.RIGHT }),
];

// ---------- Document ----------
const header = new Header({ children: [
  new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '888888', space: 2 } }, tabStops: [{ type: TabStopType.RIGHT, position: 9638 }], children: [
    R('Ignasi Tutzó Seró', { bold: true, size: 16 }), R('\tSubstitució de portes d’accés · Nau de Turbines del Besòs', { size: 16, color: '666666' }),
  ] }),
  new Paragraph({ children: [R('Arquitecte Tècnic', { size: 16 })] }),
] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [R('Pàgina ', { size: 16 }), new TextRun({ children: [PageNumber.CURRENT], size: 16 })] })] });

const A4 = { width: 11906, height: 16838 };
const margin = { top: 1418, bottom: 1134, left: 1134, right: 1134, header: 567, footer: 567 };

const doc = new Document({
  creator: 'Ignasi Tutzó Seró',
  title: 'Projecte de substitució de portes d’accés – Nau de Turbines del Besòs',
  styles: {
    default: { document: { run: { font: FONT, size: 20 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 24, bold: true, font: FONT, color: '1F3A5F' }, paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0, keepNext: true } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 21, bold: true, font: FONT }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1, keepNext: true } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 20, bold: true, italics: true, font: FONT }, paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 2, keepNext: true } },
    ],
  },
  numbering: { config: [
    { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 283 } } } }] },
    { reference: 'num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 340 } } } }] },
  ] },
  features: { updateFields: true },
  sections: [
    { properties: { page: { size: A4, margin } }, children: cover },
    { properties: { page: { size: A4, margin } }, headers: { default: header }, footers: { default: footer },
      children: [...contingut, ...index, ...dd, ...md, ...mj, ...mc, ...mn, ...signatura, BR(), ...dgIndex] },
    { properties: { page: { size: A4, margin } }, headers: { default: header }, footers: { default: footer },
      children: [...pressupost, ...planning, ...annex1, ...annex2, ...annex3, ...annex4, ...signatura] },
  ],
});

Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT, buf); console.log('written', OUT, buf.length, 'PEM', pem, 'cap05', indTotal, 'PEC', pec, 'TOTAL', totalIva, 'clear', clearTotal, pedTotal, vehClear); });
