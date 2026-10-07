// Immersia — slide unique : carte « Le problème » à gauche, solution / jalons / KPIs à droite.
// Run (depuis ce dossier, après `npm install`) : node build.js  ->  ../Immersia_slide_base.pptx
// Puis : python3 add_animations.py ../Immersia_slide_base.pptx ../Immersia_slide.pptx 'stat1-number,stat1-label;stat2-number,stat2-label' --pulse
const path = require('path');
const fs = require('fs');
const pptxgen = require('pptxgenjs');
const JSZip = require('jszip');

const BUILD = __dirname;
const { iconData } = require(path.join(BUILD, 'helpers.js'));
const { applyTheme } = require(path.join(BUILD, 'apply_theme.js'));

const OUT = process.argv[2] || path.join(BUILD, '..', 'Immersia_slide_base.pptx');

const THEME = {
  name: 'Immersia',
  headFontFace: 'Calibri',
  bodyFontFace: 'Calibri',
  colors: {
    dk1: '1C1F3A', lt1: 'FFFFFF', dk2: '4A4F6B', lt2: 'F2F3F8',
    accent1: 'E30613', accent2: '2D2D5A', accent3: '7850DC',
    accent4: '0E8A6D', accent5: 'F2A900', accent6: '8A8FA8',
    hlink: '7850DC', folHlink: '4A4F6B',
  },
};

// ---------------------------------------------------------------- geometry
// Carte « Le problème » (surface lt2, coins arrondis, marges blanches autour) : 0.4" -> 4.5"
const LX = 0.4, LW = 4.1, LPAD = 0.2;
const PX = LX + LPAD, PW = LW - 2 * LPAD;   // contenu : 0.6 -> 4.3 (3.7")
const CARD_L_TOP = 1.13, CARD_L_BOTTOM = 7.15;
const LY = -0.13;                           // décalage vertical du contenu de la carte gauche (suit l'en-tête remonté)
// Colonne droite : 4.9 -> 12.95
const RX = 4.9, RW = 8.05;
const HEAD_Y = 1.26, HEAD_H = 0.24;         // en-têtes de section « Le problème » / « La solution Immersia » sur la même ligne (encre 1.32-1.45 ; 0.19" sous le haut de la carte gauche)

// pptxgenjs écrit les puces sans police dédiée (glyphe « • » Calibri, trop petit) :
// on déclare la police Arial pour la puce, comme le fait le masque PowerPoint par défaut.
async function fixBulletFont(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const part = 'ppt/slides/slide1.xml';
  let xml = await zip.file(part).async('string');
  xml = xml.replace(/<a:buSzPct val="100000"\/><a:buChar char="&#x2022;"\/>/g,
    '<a:buSzPct val="100000"/><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="&#x2022;"/>');
  zip.file(part, xml);
  fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = 'Immersia — des labs prêts à l\'emploi, immersifs et pilotés par l\'IA';
  pres.author = 'PST&B';
  pres.lang = 'fr-FR';
  const C = pres.SchemeColor;

  // ---------------------------------------------------------------- layout
  pres.defineSlideMaster({
    title: 'IMMERSIA',
    background: { color: 'FFFFFF' },
    objects: [
      { image: { x: LX, y: 0.4, w: 1.27, h: 0.55, path: path.join(BUILD, 'pstb-logo-black.png') } },
      { image: { x: 11.74, y: 0.36, w: 1.21, h: 0.6, path: path.join(BUILD, 'opco-atlas-logo.png') } },
      { text: { text: 'Les CFA de demain', options: { x: 10.95, y: 1.06, w: 2.0, h: 0.2, fontSize: 10, color: C.text2, align: 'right', margin: 0 } } },
    ],
  });

  pres.addSection({ title: 'Immersia' });
  const slide = pres.addSlide({ masterName: 'IMMERSIA', sectionTitle: 'Immersia' });

  const T = (text, o) => slide.addText(text, Object.assign({ isTextBox: true, margin: 0 }, o));

  // ================================================================ EN-TÊTE : titre sur une ligne entre les deux logos
  T('Immersia — des labs prêts à l\'emploi, immersifs et pilotés par l\'IA',
    { x: 2.0, y: 0.35, w: 9.35, h: 0.42, fontSize: 24, bold: true, color: C.text1, valign: 'middle', objectName: 'title' });
  T('Projet PST&B soutenu par OPCO Atlas · appel à projets « Les CFA de demain »',
    { x: 2.0, y: 0.78, w: 9.35, h: 0.2, fontSize: 11, color: C.text2, valign: 'middle', objectName: 'subtitle' });   // encre -> 0.97 ; carte « Le problème » à 1.13 : 0.16" d'air

  // ================================================================ CARTE « LE PROBLÈME » (gauche)
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: LX, y: CARD_L_TOP, w: LW, h: CARD_L_BOTTOM - CARD_L_TOP, rectRadius: 0.1, fill: { color: C.background2 }, objectName: 'problem-card' });

  T('Le problème', { x: PX, y: HEAD_Y, w: PW, h: HEAD_H, fontSize: 13, bold: true, color: C.text1, valign: 'middle', objectName: 'problem-header' });

  T('Des machines et des installations différentes pour chaque étudiant : avant le premier exercice, le TD se perd en configuration et l\'enseignant dépanne au lieu d\'enseigner.',
    { x: PX, y: 1.74 + LY, w: PW, h: 0.84, fontSize: 12, color: C.text1, valign: 'top', objectName: 'problem-sentence' });

  // Chiffre héros 1 (rouge PST&B accent1 sur surface claire)
  slide.addText([
    { text: 'jusqu\'à ', options: { fontSize: 20, bold: true, color: C.accent1 } },
    { text: '30 h', options: { fontSize: 54, bold: true, color: C.accent1 } },
  ], { x: PX, y: 2.49 + LY, w: PW, h: 0.92, isTextBox: true, margin: 0, valign: 'bottom', objectName: 'stat1-number' });
  T('passées par un étudiant à installer l\'environnement d\'un projet (médiane > 2 h). Plainte n°1 : les versions de Python et des packages.',
    { x: PX, y: 3.43 + LY, w: PW, h: 0.53, fontSize: 10, color: C.text1, valign: 'top', objectName: 'stat1-label' });

  // Chiffre héros 2
  slide.addText([
    { text: '61 %', options: { fontSize: 54, bold: true, color: C.accent1 } },
  ], { x: PX, y: 3.89 + LY, w: PW, h: 0.92, isTextBox: true, margin: 0, valign: 'bottom', objectName: 'stat2-number' });
  T('des étudiants placent « aucune installation requise » parmi les fonctionnalités les plus utiles d\'un environnement de programmation.',
    { x: PX, y: 4.83 + LY, w: PW, h: 0.53, fontSize: 10, color: C.text1, valign: 'top', objectName: 'stat2-label' });

  // Trois publics : icône blanche dans un cercle accent2
  const audiences = [
    { icon: 'FaUserGraduate', title: 'Étudiants', text: ' — Zéro installation : un environnement identique pour tous, prêt en un clic.' },
    { icon: 'FaBriefcase', title: 'Apprentis', text: ' — Immersion dans des scénarios d\'entreprise réalistes : banque, conseil, assurance.' },
    { icon: 'FaChalkboardTeacher', title: 'Enseignants', text: ' — Des cas d\'usage prêts à l\'emploi et une vue de suivi de la classe.' },
  ];
  const D = 0.36, AUD_Y0 = 5.52 + LY, AUD_PITCH = 0.55;   // cercles 5.40 / 5.95 / 6.50 -> 6.86 ; 0.19" entre cercles, marge basse de la carte 0.29
  for (let i = 0; i < audiences.length; i++) {
    const a = audiences[i];
    const y = AUD_Y0 + i * AUD_PITCH;
    slide.addShape(pres.shapes.OVAL, { x: PX, y, w: D, h: D, fill: { color: C.accent2 }, objectName: `audience${i + 1}-circle` });
    slide.addImage({ data: await iconData('fa', a.icon, 'FFFFFF', 512), x: PX + 0.08, y: y + 0.08, w: 0.2, h: 0.2, objectName: `audience${i + 1}-icon` });
    slide.addText([
      { text: a.title, options: { bold: true, color: C.text1 } },
      { text: a.text, options: { color: C.text2 } },
    ], { x: PX + D + 0.12, y: y - 0.02, w: PW - D - 0.12, h: D + 0.04, fontSize: 10, isTextBox: true, margin: 0, valign: 'middle', objectName: `audience${i + 1}-text` });
  }

  // ================================================================ COLONNE DROITE
  // ---- La solution Immersia : 3 cartes reliées par des chevrons, cœur dominant au centre
  T('La solution Immersia', { x: RX, y: HEAD_Y, w: RW, h: HEAD_H, fontSize: 13, bold: true, color: C.text1, valign: 'middle', objectName: 'solution-header' });

  const CARD_TOP = HEAD_Y + 0.37, CARD2_H = 1.86, SIDE_H = 1.50, GAP = 0.3;   // carte coeur 1.63 -> 3.55 (0.17" sous l'encre de l'en-tête) ; cartes latérales 1.84 -> 3.34
  // Cartes latérales : cercle à 0.14" du haut, puces à 0.69" (même ligne de départ pour les deux) ; 4 lignes de puces -> encre
  // à 0.12-0.16" du bas : marges haute et basse équilibrées (carte 1 : 0.14/0.16, carte 3 : 0.14/0.12). Carte coeur : cercle 0.10, puces 0.56 (7 lignes, 0.14" en bas).
  const W1 = 2.16, W3 = 2.58, W2 = RW - W1 - W3 - 2 * GAP;   // 2.71
  const SIDE_Y = CARD_TOP + (CARD2_H - SIDE_H) / 2;
  const cards = [
    { x: RX, w: W1, h: SIDE_H, y: SIDE_Y, key: 'card1',
      title: ['Connexion & accès'], iconSet: 'fa', icon: 'FaKey', iconColor: 'FFFFFF',
      fill: C.background2, circle: C.accent2, titleColor: C.text1, textColor: C.text1, shadow: false, circleTop: 0.14, bulletsTop: 0.69,
      items: ['Authentification liée aux comptes étudiants', 'Espaces par promotion et par module'] },
    { x: RX + W1 + GAP, w: W2, h: CARD2_H, y: CARD_TOP, key: 'card2',
      title: ['Scénarios & Labs Immersia'], iconSet: 'fa', icon: 'FaFlask', iconColor: THEME.colors.accent2,
      fill: C.accent2, circle: C.background1, titleColor: C.background1, textColor: C.background1, shadow: true, circleTop: 0.10, bulletsTop: 0.56,
      items: ['Environnements préconfigurés, zéro installation', 'IA checker : vérification automatique des rendus', 'Extension VS Code', 'Datasets & notebooks en accès direct', 'Option no-code'] },
    { x: RX + W1 + GAP + W2 + GAP, w: W3, h: SIDE_H, y: SIDE_Y, key: 'card3',
      title: ['Vue enseignant ·', 'notre modèle LLM'], iconSet: 'lu', icon: 'LuBrainCircuit', iconColor: 'FFFFFF',
      fill: C.background2, circle: C.accent3, titleColor: C.accent3, textColor: C.text1, shadow: false, circleTop: 0.14, bulletsTop: 0.69,
      items: ['Résumé de l\'activité de la classe', 'Détection des difficultés rencontrées', 'Concepts les moins bien maîtrisés', 'Suivi des versions des soumissions'] },
  ];

  const CD = 0.4;
  for (const c of cards) {
    const opts = { x: c.x, y: c.y, w: c.w, h: c.h, fill: { color: c.fill }, rectRadius: 0.08, objectName: `${c.key}-bg` };
    if (c.shadow) opts.shadow = { type: 'outer', color: THEME.colors.dk1, blur: 3, offset: 1, angle: 90, opacity: 0.2 };
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, opts);
    slide.addShape(pres.shapes.OVAL, { x: c.x + 0.12, y: c.y + c.circleTop, w: CD, h: CD, fill: { color: c.circle }, objectName: `${c.key}-circle` });
    slide.addImage({ data: await iconData(c.iconSet, c.icon, c.iconColor, 512), x: c.x + 0.12 + 0.09, y: c.y + c.circleTop + 0.09, w: 0.22, h: 0.22, objectName: `${c.key}-icon` });
    slide.addText(
      c.title.map((t, i) => ({ text: t, options: { breakLine: i < c.title.length - 1 } })),
      { x: c.x + 0.12 + CD + 0.1, y: c.y + c.circleTop - 0.01, w: c.w - 0.72, h: 0.42, fontSize: 12, bold: true, color: c.titleColor, valign: 'middle', isTextBox: true, margin: 0, objectName: `${c.key}-title` });
    slide.addText(
      c.items.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, color: c.textColor, breakLine: i < c.items.length - 1, paraSpaceAfter: c.key === 'card2' ? 0 : 1 } })),
      { x: c.x + 0.12, y: c.y + c.bulletsTop, w: c.w - 0.24, h: c.h - c.bulletsTop - 0.05, fontSize: 10, isTextBox: true, margin: 0, valign: 'top', objectName: `${c.key}-bullets` });   // coeur : 7 lignes de 10 pt (1.27") dans 1.31" ; carte 3 : 4 lignes (0.72") dans 0.74"
  }
  // chevrons entre les cartes (centrés sur la carte cœur)
  const CHEV_W = 0.16, CHEV_H = 0.34, chevY = CARD_TOP + CARD2_H / 2 - CHEV_H / 2;
  slide.addShape(pres.shapes.CHEVRON, { x: cards[0].x + cards[0].w + (GAP - CHEV_W) / 2, y: chevY, w: CHEV_W, h: CHEV_H, fill: { color: C.accent6 }, objectName: 'chevron1' });
  slide.addShape(pres.shapes.CHEVRON, { x: cards[1].x + cards[1].w + (GAP - CHEV_W) / 2, y: chevY, w: CHEV_W, h: CHEV_H, fill: { color: C.accent6 }, objectName: 'chevron2' });

  // ---- Jalons : timeline horizontale, 5 colonnes de largeur égale
  const TL_HEAD_Y = CARD_TOP + CARD2_H + 0.13;        // 3.68 (encre 3.74-3.90 ; carte coeur 3.55 -> 0.19", ombre ~3.58 -> 0.16")
  slide.addText([
    { text: 'Jalons vers la bêta MVP', options: { bold: true, color: C.text1 } },
    { text: ' (dates indicatives)', options: { color: C.text2 } },
  ], { x: RX, y: TL_HEAD_Y, w: RW, h: HEAD_H, fontSize: 13, isTextBox: true, margin: 0, valign: 'middle', objectName: 'timeline-header' });

  const DATE_Y = TL_HEAD_Y + 0.39;    // 4.07 (pastille ambre 4.06-4.26 : 0.16" sous l'encre de l'en-tête, dont le « ) » s'arrête à x 8.00 juste avant la pastille à 8.02)
  const LINE_Y = DATE_Y + 0.45;       // 4.52 (pastille 4.26 -> point ambre 4.43 : 0.17" d'air ; encre des dates -> points : ~0.21")
  const LABEL_Y = LINE_Y + 0.22;      // 4.74 (bas des points 4.60/4.61 -> capitales des libellés 4.77 : >= 0.16") ; boîte -> 5.42
  const LABEL_H = 0.68;               // 4 lignes de 10 pt max (Bêta MVP : encre 0.62", jambages 0.65")
  const PITCH = RW / 5;               // 1.61
  const LABEL_W = PITCH - 0.15;       // >= 0.15" entre libellés voisins
  const PILL_W = 0.70, PILL_PAD = 0.095; // « Fév. 2027 » en 10 pt gras ~0.51" : ~0.095" d'ambre de chaque côté
  const milestones = [
    { date: 'Oct. 2026', name: 'Cadrage', rest: ' — scénarios cibles validés' },
    { date: 'Déc. 2026', name: 'Prototype', rest: ' — connexion +\u00A01er scénario (banque)' },
    { date: 'Fév. 2027', name: 'Bêta MVP', rest: ' — 3 scénarios (banque, conseil, assurance), IA checker, vue enseignant', hero: true },
    { date: 'Avr. 2027', name: 'Pilote', rest: ' — une promotion, enseignants volontaires' },
    { date: 'Juin 2027', name: 'Bilan KPIs', rest: ' — V1' },
  ];
  // ligne de temps (du 1er jalon jusqu'au bord droit de la colonne, flèche)
  slide.addShape(pres.shapes.LINE, { x: RX + 0.08, y: LINE_Y, w: RW - 0.08, h: 0, line: { color: C.accent6, width: 1.5, endArrowType: 'triangle' }, objectName: 'timeline-line' });
  milestones.forEach((m, i) => {
    const x = RX + i * PITCH;
    const dot = m.hero ? 0.18 : 0.16;
    slide.addShape(pres.shapes.OVAL, { x: x + 0.08 - dot / 2, y: LINE_Y - dot / 2, w: dot, h: dot, fill: { color: m.hero ? C.accent5 : C.accent2 }, line: { color: C.background1, width: 1.5 }, objectName: `milestone${i + 1}-dot` });
    if (m.hero) {
      // pastille ambre : texte centré, bord gauche du texte aligné sur le libellé du dessous ; marge [l, r, b, t] en pt :
      // valign 'top' + interligne 92 % (et non 'middle') : la ligne de 10 pt porte plus d'espace au-dessus des capitales
      // qu'en dessous de la ligne de base, un centrage géométrique fait tomber « Fév. 2027 » trop bas ; ainsi calées, les
      // capitales sont au centre optique de la pastille (~0.05" d'ambre dessus et dessous)
      slide.addText(m.date, { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, x: x - PILL_PAD, y: DATE_Y - 0.01, w: PILL_W, h: 0.2, fill: { color: C.accent5 }, fontSize: 10, bold: true, color: C.text1, valign: 'top', align: 'center', margin: 0, lineSpacingMultiple: 0.92, isTextBox: true, objectName: `milestone${i + 1}-date` });
    } else {
      T(m.date, { x, y: DATE_Y, w: LABEL_W, h: 0.22, fontSize: 10, bold: true, color: C.text1, valign: 'middle', objectName: `milestone${i + 1}-date` });
    }
    slide.addText([
      { text: m.name, options: { bold: true, color: C.text1 } },
      { text: m.rest, options: { color: C.text2 } },
    ], { x, y: LABEL_Y, w: LABEL_W, h: LABEL_H, fontSize: 10, isTextBox: true, margin: 0, valign: 'top', objectName: `milestone${i + 1}-label` });
  });

  // ---- KPIs cibles du MVP : 4 tuiles
  const KPI_HEAD_Y = LABEL_Y + LABEL_H + 0.01;   // 5.43 (encre 5.47 ; le libellé Bêta MVP, colonne 3, n'a aucun recouvrement horizontal avec l'en-tête ; colonne 1 : 0.39")
  T('KPIs cibles du MVP', { x: RX, y: KPI_HEAD_Y, w: RW, h: HEAD_H, fontSize: 13, bold: true, color: C.text1, valign: 'middle', objectName: 'kpi-header' });
  const KPI_Y = KPI_HEAD_Y + HEAD_H + 0.11;      // 5.78 -> 6.71 ; sources à 6.84 (capitales 6.87 : 0.16") ; 0.16" d'encre entre l'en-tête et les tuiles
  const KPI_H = 0.93;                            // encre : chiffre à ~0.09" du haut, ligne de base du libellé (3 lignes) à ~0.11" du bas
  const kpis = [
    ['≥ 90 %', 'étudiants connectés via leur compte dès la 1re semaine'],
    ['< 5 min', 'pour démarrer un lab (contre une séance perdue aujourd\'hui)'],
    ['≥ 2 / sem.', 'sessions par étudiant et par module'],
    ['≥ 4 / 5', 'satisfaction étudiants et enseignants pilotes'],
  ];
  const KGAP = 0.15, KW = (RW - 3 * KGAP) / 4;
  kpis.forEach(([num, label], i) => {
    const x = RX + i * (KW + KGAP);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: KPI_Y, w: KW, h: KPI_H, rectRadius: 0.06, fill: { color: C.background2 }, objectName: `kpi${i + 1}-bg` });
    T(num, { x: x + 0.1, y: KPI_Y + 0.01, w: KW - 0.2, h: 0.34, fontSize: 20, bold: true, color: C.accent4, valign: 'middle', objectName: `kpi${i + 1}-number` });
    T(label, { x: x + 0.1, y: KPI_Y + 0.35, w: KW - 0.2, h: 0.54, fontSize: 10, color: C.text2, valign: 'top', objectName: `kpi${i + 1}-label` });
  });

  // ---- Sources
  slide.addText([
    { text: 'Sources : (1) Storks, Yu, Ma & Chai, « NLP Reproducibility For All », ACL 2023 — 93 étudiants, Univ. du Michigan.', options: { breakLine: true } },
    { text: '(2) Velez et al., « Student Adoption and Perceptions of a Web IDE », ACM SIGCSE 2020 — enquête, 140 étudiants, UC Davis.' },
  ], { x: RX, y: 6.84, w: RW, h: 0.31, fontSize: 9, color: C.text2, isTextBox: true, margin: 0, valign: 'bottom', objectName: 'sources' });

  slide.addNotes('Immersia — slide unique pour la commission d\'ingénierie pédagogique de PST&B.');

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  await fixBulletFont(OUT);
  console.log('written', OUT);
})().catch((e) => { console.error(e); process.exit(1); });
