const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "PFE - Simulation OpenFOAM Barrage Ahl Souss";

// ─── PALETTE ────────────────────────────────────────────────────────────────
const C = {
  navy:    "0D2B45",   // dominant dark
  blue:    "1A6FA8",   // ocean blue
  teal:    "00A8CC",   // accent
  ltBlue:  "E4F3FB",   // light content bg
  white:   "FFFFFF",
  offwt:   "F4F8FC",
  dark:    "1B2B3A",
  muted:   "5C7A95",
  gold:    "F0A500",   // highlight accent
};

const FIGDIR = "/workspace/figures/";
const makeShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 3, angle: 45, opacity: 0.13 });

// ─── HELPERS ────────────────────────────────────────────────────────────────
function darkSlide(slide) {
  slide.background = { color: C.navy };
}
function lightSlide(slide) {
  slide.background = { color: C.offwt };
}
function addSectionTag(slide, text) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.35, y: 0.18, w: 2.1, h: 0.34,
    fill: { color: C.teal }, rectRadius: 0.05,
  });
  slide.addText(text, {
    x: 0.35, y: 0.18, w: 2.1, h: 0.34,
    fontSize: 9, bold: true, color: C.white,
    align: "center", valign: "middle", margin: 0,
  });
}
function addSlideTitle(slide, text, dark = false) {
  slide.addText(text, {
    x: 0.35, y: 0.58, w: 9.3, h: 0.65,
    fontSize: 26, bold: true,
    color: dark ? C.white : C.navy,
    align: "left", valign: "middle",
  });
}
function addDivider(slide, dark = false) {
  slide.addShape(pres.shapes.LINE, {
    x: 0.35, y: 1.28, w: 9.3, h: 0,
    line: { color: dark ? C.teal : C.blue, width: 1.5 },
  });
}
function card(slide, x, y, w, h, fillColor) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: fillColor },
    rectRadius: 0.1,
    shadow: makeShadow(),
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 1 — TITLE
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkSlide(s);

  // Background accent shape
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 3.5, w: 10, h: 2.125,
    fill: { color: "0A1F33" },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 3.48, w: 10, h: 0.06,
    fill: { color: C.teal },
  });

  // University / institution
  s.addText("École Nationale des Sciences Appliquées — Agadir", {
    x: 0.5, y: 0.3, w: 9, h: 0.4,
    fontSize: 11, color: C.teal, bold: false, align: "center",
  });

  // Main title
  s.addText("Modélisation numérique 3D de l'évacuateur de crues\ndu barrage Ahl Souss par OpenFOAM", {
    x: 0.5, y: 0.82, w: 9, h: 1.8,
    fontSize: 30, bold: true, color: C.white, align: "center", valign: "middle",
  });

  // Subtitle bar
  s.addText("Projet de Fin d'Études — Génie Civil & Hydraulique", {
    x: 0.5, y: 2.72, w: 9, h: 0.42,
    fontSize: 14, color: C.teal, align: "center",
  });

  // Bottom info
  s.addText("Présenté par :  Hiba Asmae BASTOUT", {
    x: 0.5, y: 3.65, w: 4.5, h: 0.38,
    fontSize: 12, color: C.white, bold: true,
  });
  s.addText("Encadrant :  [Nom de l'encadrant]", {
    x: 0.5, y: 4.02, w: 4.5, h: 0.35,
    fontSize: 11, color: "A8C8E0",
  });
  s.addText("Année universitaire 2025 – 2026", {
    x: 5.5, y: 3.65, w: 4, h: 0.38,
    fontSize: 12, color: "A8C8E0", align: "right",
  });
  s.addText("Session de Juin 2026", {
    x: 5.5, y: 4.02, w: 4, h: 0.35,
    fontSize: 11, color: "A8C8E0", align: "right",
  });

  s.addNotes("Slide de titre. Présenter brièvement le sujet et les intervenants.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 2 — SOMMAIRE
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "PLAN");
  addSlideTitle(s, "Sommaire");
  addDivider(s);

  const items = [
    ["01", "Contexte du projet & Présentation du barrage Ahl Souss"],
    ["02", "Objectifs et démarche méthodologique"],
    ["03", "Dimensionnement hydraulique théorique"],
    ["04", "Modélisation numérique avec OpenFOAM"],
    ["05", "Résultats des simulations"],
    ["06", "Comparaison théorie / simulation & Conclusion"],
  ];

  items.forEach(([num, label], i) => {
    const col = i < 3 ? 0 : 1;
    const row = i % 3;
    const x = col === 0 ? 0.35 : 5.15;
    const y = 1.45 + row * 1.2;

    // Number circle
    card(s, x, y, 0.7, 0.7, C.navy);
    s.addText(num, {
      x, y, w: 0.7, h: 0.7,
      fontSize: 16, bold: true, color: C.teal,
      align: "center", valign: "middle", margin: 0,
    });
    // Label box
    card(s, x + 0.8, y, 3.75, 0.7, C.white);
    s.addText(label, {
      x: x + 0.9, y: y + 0.05, w: 3.55, h: 0.6,
      fontSize: 12, color: C.dark,
      align: "left", valign: "middle",
    });
  });

  s.addNotes("Présenter le plan en 6 parties principales.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 3 — CONTEXTE
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "CONTEXTE");
  addSlideTitle(s, "Présentation du projet — Barrage Ahl Souss");
  addDivider(s);

  // Left text column
  const facts = [
    ["Localisation", "Oued Souss, province de Taroudant, Maroc"],
    ["Hauteur du barrage", "56 m (type poids en BCR)"],
    ["Volume de retenue", "480 Mm³"],
    ["Cote seuil Creager", "600,0 m NGM"],
    ["Largeur déversante", "70 m (1 passe)"],
    ["Débit millénaire laminé", "Q₁₀₀₀ = 794,93 m³/s"],
    ["Type d'évacuateur", "Déversoir Creager + coursier + cuillère saut de ski"],
  ];

  facts.forEach(([key, val], i) => {
    const y = 1.45 + i * 0.52;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 2.3, h: 0.42,
      fill: { color: C.navy },
    });
    s.addText(key, {
      x: 0.35, y, w: 2.3, h: 0.42,
      fontSize: 10, bold: true, color: C.teal,
      align: "center", valign: "middle", margin: 0,
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 2.65, y, w: 3.8, h: 0.42,
      fill: { color: i % 2 === 0 ? C.ltBlue : C.white },
    });
    s.addText(val, {
      x: 2.75, y: y + 0.01, w: 3.6, h: 0.4,
      fontSize: 10, color: C.dark, valign: "middle",
    });
  });

  // Right: location figure
  s.addImage({
    path: FIGDIR + "figure4_location.png",
    x: 6.85, y: 1.4, w: 2.8, h: 2.8,
    sizing: { type: "contain", w: 2.8, h: 2.8 },
  });
  s.addText("Localisation du barrage Ahl Souss", {
    x: 6.85, y: 4.2, w: 2.8, h: 0.3,
    fontSize: 9, color: C.muted, align: "center", italic: true,
  });

  // Bottom note
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.35, y: 5.05, w: 9.3, h: 0.42,
    fill: { color: C.navy }, rectRadius: 0.06,
  });
  s.addText("Problématique : Valider le dimensionnement hydraulique de l'évacuateur de crues par simulation CFD 3D avec OpenFOAM", {
    x: 0.45, y: 5.05, w: 9.1, h: 0.42,
    fontSize: 10, color: C.white, align: "center", valign: "middle", bold: true,
  });

  s.addNotes("Présenter le barrage Ahl Souss et poser la problématique centrale du PFE.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 4 — OBJECTIFS
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "OBJECTIFS");
  addSlideTitle(s, "Objectifs du Projet de Fin d'Études");
  addDivider(s);

  const objs = [
    { num: "01", title: "Dimensionnement théorique", body: "Calculer le profil hydraulique (seuil, coursier, cuillère), la ligne d'eau (HEC-RAS), la trajectoire du jet et l'affouillement à l'aval pour Q₁₀, Q₁₀₀ et Q₁₀₀₀.", color: C.navy },
    { num: "02", title: "Modélisation 3D OpenFOAM", body: "Construire un modèle CFD complet (géométrie SketchUp → STL, maillage snappyHexMesh, solveur interFoam) pour les trois débits de crue.", color: "0E5A8A" },
    { num: "03", title: "Analyse des résultats", body: "Extraire et analyser les champs de vitesse, de pression et de fraction volumique (α) issus des simulations pour chaque débit.", color: "065A7A" },
    { num: "04", title: "Comparaison & Validation", body: "Confronter les résultats numériques aux valeurs théoriques (hauteurs d'eau, vitesses maximales) et quantifier les écarts.", color: "1A4A6A" },
  ];

  objs.forEach((o, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = col === 0 ? 0.35 : 5.15;
    const y = 1.45 + row * 1.9;

    card(s, x, y, 4.6, 1.72, C.white);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.15, y: y + 0.15, w: 0.75, h: 0.75,
      fill: { color: o.color }, rectRadius: 0.08,
    });
    s.addText(o.num, {
      x: x + 0.15, y: y + 0.15, w: 0.75, h: 0.75,
      fontSize: 18, bold: true, color: C.gold,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(o.title, {
      x: x + 1.05, y: y + 0.15, w: 3.4, h: 0.42,
      fontSize: 13, bold: true, color: C.navy, valign: "middle",
    });
    s.addText(o.body, {
      x: x + 0.25, y: y + 0.72, w: 4.2, h: 0.85,
      fontSize: 10.5, color: C.dark, valign: "top",
    });
  });

  s.addNotes("Expliquer les 4 objectifs principaux du PFE en lien avec la problématique.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 5 — MÉTHODOLOGIE
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "MÉTHODOLOGIE");
  addSlideTitle(s, "Démarche méthodologique");
  addDivider(s);

  const steps = [
    { n: "1", label: "Hydrologie", sub: "Débits de crue\nQ₁₀ / Q₁₀₀ / Q₁₀₀₀", color: C.navy },
    { n: "2", label: "Hydraulique\nthéorique", sub: "Seuil · Coursier\nCuillère · HEC-RAS", color: "0E5A8A" },
    { n: "3", label: "Géométrie 3D", sub: "SketchUp → STL\nAutoCAD profil", color: "065A7A" },
    { n: "4", label: "Maillage\nOpenFOAM", sub: "blockMesh +\nsnappyHexMesh", color: "0A7B9C" },
    { n: "5", label: "Simulation\nCFD", sub: "interFoam\nVOF diphasique", color: C.teal },
  ];

  steps.forEach((st, i) => {
    const x = 0.35 + i * 1.88;
    // Box
    card(s, x, 1.5, 1.62, 2.3, st.color);
    s.addText(st.n, {
      x, y: 1.5, w: 1.62, h: 0.52,
      fontSize: 22, bold: true, color: C.gold,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(st.label, {
      x, y: 2.05, w: 1.62, h: 0.65,
      fontSize: 12, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(st.sub, {
      x, y: 2.72, w: 1.62, h: 0.72,
      fontSize: 9.5, color: "B8D8EA",
      align: "center", valign: "top", margin: 0,
    });
    // Arrow (not after last)
    if (i < steps.length - 1) {
      s.addShape(pres.shapes.LINE, {
        x: x + 1.62, y: 2.6, w: 0.26, h: 0,
        line: { color: C.teal, width: 2 },
      });
    }
  });

  // Bottom: outputs row
  const outputs = [
    "Ligne d'eau\nHEC-RAS",
    "Champs U & P\nOpenFOAM",
    "Trajectoire du jet\n& Affouillement",
    "Comparaison\nthéorie / CFD",
  ];
  s.addText("RÉSULTATS", {
    x: 0.35, y: 4.1, w: 1.5, h: 0.35,
    fontSize: 10, bold: true, color: C.teal, valign: "middle",
  });
  outputs.forEach((txt, i) => {
    card(s, 0.35 + i * 2.35, 4.45, 2.1, 0.75, C.ltBlue);
    s.addText(txt, {
      x: 0.35 + i * 2.35, y: 4.45, w: 2.1, h: 0.75,
      fontSize: 10, color: C.navy, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  s.addNotes("Décrire la chaîne de travail complète du PFE.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 6 — DÉBITS DE CRUE
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "THÉORIE");
  addSlideTitle(s, "Hydrogrammes de crue — Débits laminés retenus");
  addDivider(s);

  // 3 stat cards
  const debits = [
    { T: "Q₁₀", Q: "160,74 m³/s", Hd: "1,13 m", label: "Débit décennal", color: "1A6FA8" },
    { T: "Q₁₀₀", Q: "507,54 m³/s", Hd: "2,32 m", label: "Débit centennal", color: "065A7A" },
    { T: "Q₁₀₀₀", Q: "794,93 m³/s", Hd: "3,07 m", label: "Débit millénaire", color: C.navy },
  ];

  debits.forEach((d, i) => {
    const x = 0.35 + i * 3.2;
    card(s, x, 1.5, 2.9, 2.6, d.color);
    s.addText(d.T, {
      x, y: 1.5, w: 2.9, h: 0.7,
      fontSize: 28, bold: true, color: C.gold,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(d.label, {
      x, y: 2.18, w: 2.9, h: 0.38,
      fontSize: 11, color: "B0CCE0", align: "center", italic: true,
    });
    s.addText(`Q = ${d.Q}`, {
      x, y: 2.58, w: 2.9, h: 0.38,
      fontSize: 13, bold: true, color: C.white, align: "center",
    });
    s.addText(`Hd = ${d.Hd}`, {
      x, y: 2.98, w: 2.9, h: 0.35,
      fontSize: 12, color: C.teal, align: "center",
    });
    s.addText("(charge sur seuil Creager)", {
      x, y: 3.33, w: 2.9, h: 0.3,
      fontSize: 9, color: "8AB0CA", align: "center", italic: true,
    });
  });

  // Table résultats théoriques synthèse
  const tableData = [
    [
      { text: "T (ans)", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "Q (m³/s)", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "y_coursier (m)", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "H_mur (m)", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "V₀ cuillère (m/s)", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "Lx (m)", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
    ],
    ["10", "160,74", "1,10", "1,82", "17,89", "36,72"],
    ["100", "507,54", "1,60", "2,39", "16,81", "35,36"],
    ["1 000", "794,93", "2,00", "2,83", "17,91", "38,07"],
  ];

  s.addTable(tableData, {
    x: 0.35, y: 4.3, w: 9.3, h: 1.1,
    colW: [1.1, 1.4, 1.6, 1.4, 1.8, 1.4],
    border: { pt: 0.5, color: "C0D8EC" },
    align: "center",
    fontSize: 11,
  });

  s.addNotes("Présenter les trois débits laminés et la synthèse des résultats théoriques.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 7 — CHAÎNE DE CALCUL HYDRAULIQUE
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "THÉORIE");
  addSlideTitle(s, "Chaîne de calcul hydraulique — Tableur Excel");
  addDivider(s);

  const steps = [
    {
      title: "① Seuil\nCreager",
      color: C.navy,
      entree: "Q (m³/s)\nL = 70 m\nCote seuil = 600 m NGM",
      methode: "Q = C · L · Hd^1.5",
      resultat: "Hd = 1,13 / 2,32 / 3,07 m\npour Q₁₀ / Q₁₀₀ / Q₁₀₀₀",
    },
    {
      title: "② Profil\nWES",
      color: "0E5A8A",
      entree: "Hd\nk = 0,5\nn = 1,872",
      methode: "y^n = k · Hd^(n-1) · x\n(coordonnées X-Z)",
      resultat: "Profil géométrique\ndu déversoir point par point",
    },
    {
      title: "③ Courbe\nde remous",
      color: "065A7A",
      entree: "Profil, Ks = 70\npente 1,25 m/m\nSeuil + Coursier + Cuillère",
      methode: "Méthode pas-à-pas\n(Manning-Strickler)",
      resultat: "h(x) et V(x)\nsur toute la structure\nH mur = 2,83 m",
    },
    {
      title: "④ Trajectoire\ndu jet",
      color: "0A7B9C",
      entree: "V₀ ≈ 18 m/s\nθ = 40°\nR = 3,00 m",
      methode: "Équation balistique\n(Lx = f(V₀, θ, ΔZ))",
      resultat: "Lx = 36 à 38 m\nselon débit",
    },
    {
      title: "⑤ Affouille-\nment",
      color: C.teal,
      entree: "q (m²/s)\nH (dénivelée PHE-jet)\nh* (dénivelée PHE-PHE aval)",
      methode: "Véronèse · Damle\nMartins · Chian\n(4 formules empiriques)",
      resultat: "D = 8,96 m\nCote fond = 587,23 m NGM",
    },
  ];

  const cardW = 1.68;
  const cardH = 3.72;
  const startX = 0.35;
  const startY = 1.42;
  const gap = 0.22;

  steps.forEach((st, i) => {
    const x = startX + i * (cardW + gap);

    // Header
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: startY, w: cardW, h: 0.60,
      fill: { color: st.color },
      rectRadius: 0.08,
      shadow: makeShadow(),
    });
    s.addText(st.title, {
      x, y: startY, w: cardW, h: 0.60,
      fontSize: 10, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 2,
    });

    // Card body
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: startY + 0.60, w: cardW, h: cardH - 0.60,
      fill: { color: C.white },
      line: { color: "C8DDEF", width: 0.75 },
      rectRadius: 0.05,
      shadow: makeShadow(),
    });

    // ENTRÉE
    s.addText("ENTRÉE", {
      x: x + 0.06, y: startY + 0.66, w: cardW - 0.12, h: 0.20,
      fontSize: 7, bold: true, color: st.color,
    });
    s.addText(st.entree, {
      x: x + 0.06, y: startY + 0.84, w: cardW - 0.12, h: 0.68,
      fontSize: 8.5, color: C.dark,
    });

    // Separator
    s.addShape(pres.shapes.LINE, {
      x: x + 0.1, y: startY + 1.55, w: cardW - 0.2, h: 0,
      line: { color: "C8DDEF", width: 0.5 },
    });

    // MÉTHODE
    s.addText("MÉTHODE", {
      x: x + 0.06, y: startY + 1.62, w: cardW - 0.12, h: 0.20,
      fontSize: 7, bold: true, color: st.color,
    });
    s.addText(st.methode, {
      x: x + 0.06, y: startY + 1.80, w: cardW - 0.12, h: 0.60,
      fontSize: 8.5, color: C.dark, italic: true,
    });

    // RÉSULTAT box
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.06, y: startY + 2.55, w: cardW - 0.12, h: 0.82,
      fill: { color: st.color },
      rectRadius: 0.05,
    });
    s.addText(st.resultat, {
      x: x + 0.06, y: startY + 2.55, w: cardW - 0.12, h: 0.82,
      fontSize: 9, bold: true, color: C.white,
      align: "center", valign: "middle", margin: 3,
    });

    // Arrow to next
    if (i < steps.length - 1) {
      s.addShape(pres.shapes.LINE, {
        x: x + cardW + 0.02, y: startY + cardH / 2,
        w: gap - 0.04, h: 0,
        line: { color: C.teal, width: 2.5 },
      });
      // Arrowhead (small triangle simulation with text)
      s.addText("▶", {
        x: x + cardW + gap - 0.18, y: startY + cardH / 2 - 0.15,
        w: 0.2, h: 0.3,
        fontSize: 8, color: C.teal,
      });
    }
  });

  // Bottom bar
  card(s, 0.35, 5.22, 9.3, 0.34, C.navy);
  s.addText(
    "Scénario dimensionnant : Q₁₀₀₀ = 794,93 m³/s  ·  D = 8,96 m  →  Cote fond fosse = 587,23 m NGM  ·  Fondations non menacées ✓",
    {
      x: 0.45, y: 5.22, w: 9.1, h: 0.34,
      fontSize: 10, color: C.gold, bold: true,
      align: "center", valign: "middle", margin: 0,
    }
  );

  s.addNotes("Présenter la chaîne des 5 étapes de calcul Excel : seuil Creager → profil WES → courbe de remous (seuil+coursier+cuillère) → trajectoire du jet → affouillement. Insister que la courbe de remous couvre toute la structure.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 7B — RÉSULTATS DIMENSIONNEMENT
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "THÉORIE");
  addSlideTitle(s, "Résultats du dimensionnement — 3 scénarios de crue");
  addDivider(s);

  // Table
  const tbl = [
    [
      { text: "Paramètre", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "left" } },
      { text: "Q₁₀ = 161 m³/s", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "Q₁₀₀ = 508 m³/s", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "Q₁₀₀₀ = 795 m³/s", options: { bold: true, color: C.gold, fill: { color: C.navy }, align: "center" } },
    ],
    [
      "PHE (m NGM)",
      { text: "601,1", options: { align: "center" } },
      { text: "602,3", options: { align: "center" } },
      { text: "603,1", options: { align: "center", bold: true, color: C.navy } },
    ],
    [
      "Charge sur seuil Hd (m)",
      { text: "1,13", options: { align: "center" } },
      { text: "2,32", options: { align: "center" } },
      { text: "3,07", options: { align: "center", bold: true, color: C.navy } },
    ],
    [
      "Hauteur d'eau au seuil (m)",
      { text: "0,81", options: { align: "center" } },
      { text: "1,75", options: { align: "center" } },
      { text: "2,36", options: { align: "center", bold: true, color: C.navy } },
    ],
    [
      "Hauteur d'eau coursier (m)",
      { text: "1,10", options: { align: "center" } },
      { text: "1,60", options: { align: "center" } },
      { text: "2,00", options: { align: "center", bold: true, color: C.navy } },
    ],
    [
      "Hauteur mur bajoyer (m)",
      { text: "1,82", options: { align: "center" } },
      { text: "2,39", options: { align: "center" } },
      { text: "2,83", options: { align: "center", bold: true, color: C.navy } },
    ],
    [
      "Vitesse cuillère V₀ (m/s)",
      { text: "17,89", options: { align: "center" } },
      { text: "16,81", options: { align: "center" } },
      { text: "17,91", options: { align: "center", bold: true, color: C.navy } },
    ],
    [
      "Distance d'impact Lx (m)",
      { text: "36,72", options: { align: "center" } },
      { text: "35,36", options: { align: "center" } },
      { text: "38,07", options: { align: "center", bold: true, color: C.navy } },
    ],
    [
      { text: "Profondeur affouillement D (m)", options: { bold: true, fill: { color: "D4EAF7" } } },
      { text: "3,65", options: { align: "center", bold: true, fill: { color: "D4EAF7" } } },
      { text: "6,78", options: { align: "center", bold: true, fill: { color: "D4EAF7" } } },
      { text: "8,96", options: { align: "center", bold: true, color: C.navy, fill: { color: "D4EAF7" } } },
    ],
  ];

  s.addTable(tbl, {
    x: 0.35, y: 1.42, w: 9.3, h: 3.55,
    colW: [3.5, 1.85, 1.85, 2.1],
    border: { pt: 0.5, color: "C0D8EC" },
    fontSize: 11,
    align: "left",
  });

  // Conclusion box
  card(s, 0.35, 5.1, 9.3, 0.46, C.navy);
  s.addText("Scénario dimensionnant : Q₁₀₀₀   →   D = 8,96 m   →   Cote fond fosse = 587,23 m NGM   →   Fondations non menacées  ✓", {
    x: 0.45, y: 5.1, w: 9.1, h: 0.46,
    fontSize: 12, color: C.gold, bold: true,
    align: "center", valign: "middle",
  });

  s.addNotes("Présenter le tableau des résultats pour les 3 débits. Mettre en valeur Q1000 comme scénario dimensionnant. Conclure sur la sécurité des fondations.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 8A — DÉMARCHE NUMÉRIQUE (workflow)
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "MODÉLISATION");
  addSlideTitle(s, "Démarche de la simulation numérique 3D");
  addDivider(s);

  const steps = [
    { num: "①", title: "AutoCAD", sub: "Plans 2D côtés\n(profil, coupes)", color: C.navy },
    { num: "②", title: "SketchUp", sub: "Modèle 3D\n(export .STL)", color: "0E5A8A" },
    { num: "③", title: "OpenFOAM\nMaillage", sub: "blockMesh\nsnappyHexMesh", color: "065A7A" },
    { num: "④", title: "OpenFOAM\nSimulation", sub: "interFoam\nVOF diphasique", color: "0A7B9C" },
    { num: "⑤", title: "ParaView", sub: "Post-traitement\nVisualisation", color: C.teal },
  ];

  const cw = 1.68, ch = 2.8, gap = 0.12, startX = 0.35, startY = 1.5;
  steps.forEach((st, i) => {
    const x = startX + i * (cw + gap);

    // Card background
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: startY, w: cw, h: ch,
      fill: { color: st.color }, rectRadius: 0.1, shadow: makeShadow(),
    });

    // Number circle
    s.addShape(pres.shapes.OVAL, {
      x: x + cw/2 - 0.28, y: startY + 0.15, w: 0.56, h: 0.56,
      fill: { color: C.white },
    });
    s.addText(st.num, {
      x: x + cw/2 - 0.28, y: startY + 0.15, w: 0.56, h: 0.56,
      fontSize: 14, bold: true, color: st.color,
      align: "center", valign: "middle", margin: 0,
    });

    // Title
    s.addText(st.title, {
      x: x + 0.06, y: startY + 0.82, w: cw - 0.12, h: 0.75,
      fontSize: 14, bold: true, color: C.white,
      align: "center", valign: "middle",
    });

    // Sub
    s.addText(st.sub, {
      x: x + 0.06, y: startY + 1.65, w: cw - 0.12, h: 0.95,
      fontSize: 10, color: "D0EAF5",
      align: "center", valign: "top",
    });

    // Arrow
    if (i < steps.length - 1) {
      s.addText("▶", {
        x: x + cw, y: startY + ch/2 - 0.18,
        w: gap + 0.04, h: 0.36,
        fontSize: 13, color: C.teal,
        align: "center", valign: "middle",
      });
    }
  });

  // Bottom note
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.35, y: 4.55, w: 9.3, h: 0.55,
    fill: { color: C.ltBlue }, rectRadius: 0.07,
  });
  s.addText("Objectif : reproduire fidèlement l'écoulement diphasique eau/air sur l'évacuateur de crues et valider les résultats théoriques", {
    x: 0.45, y: 4.55, w: 9.1, h: 0.55,
    fontSize: 11, color: C.navy, align: "center", valign: "middle", italic: true,
  });

  s.addNotes("Présenter la chaîne de travail : AutoCAD → SketchUp → OpenFOAM (maillage + simulation) → ParaView.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 8B — PRÉPARATION GÉOMÉTRIQUE (AutoCAD + SketchUp)
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "MODÉLISATION");
  addSlideTitle(s, "Préparation du modèle géométrique");
  addDivider(s);

  // Left — AutoCAD
  card(s, 0.35, 1.42, 4.5, 0.42, C.navy);
  s.addText("① AutoCAD — Plans 2D de référence", {
    x: 0.45, y: 1.42, w: 4.3, h: 0.42,
    fontSize: 13, bold: true, color: C.teal, valign: "middle",
  });
  s.addImage({
    path: FIGDIR + "ch8_profil_clair.jpeg",
    x: 0.35, y: 1.9, w: 4.5, h: 2.4,
    sizing: { type: "contain", w: 4.5, h: 2.4 },
  });
  const autocadPoints = [
    "Profil en long de l'évacuateur coté",
    "Seuil Creager + coursier (pente 1,25 m/m) + cuillère (R=3 m, θ=40°)",
    "Dimensions issues du dimensionnement théorique",
  ];
  autocadPoints.forEach((p, i) => {
    s.addText(`•  ${p}`, {
      x: 0.45, y: 4.38 + i * 0.22, w: 4.2, h: 0.22,
      fontSize: 9.5, color: C.dark,
    });
  });

  // Right — SketchUp
  card(s, 5.15, 1.42, 4.5, 0.42, "065A7A");
  s.addText("② SketchUp — Modèle 3D & export STL", {
    x: 5.25, y: 1.42, w: 4.3, h: 0.42,
    fontSize: 13, bold: true, color: C.teal, valign: "middle",
  });
  s.addImage({
    path: FIGDIR + "ch8_maillage_grossier.jpeg",
    x: 5.15, y: 1.9, w: 4.5, h: 2.4,
    sizing: { type: "contain", w: 4.5, h: 2.4 },
  });
  const sketchupPoints = [
    "Construction 3D à partir des plans AutoCAD",
    "Structure complète : seuil + coursier + cuillère",
    "Export au format .STL pour OpenFOAM",
  ];
  sketchupPoints.forEach((p, i) => {
    s.addText(`•  ${p}`, {
      x: 5.25, y: 4.38 + i * 0.22, w: 4.2, h: 0.22,
      fontSize: 9.5, color: C.dark,
    });
  });

  s.addNotes("Présenter la préparation géométrique en deux étapes : plans 2D AutoCAD puis modèle 3D SketchUp exporté en STL.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 8C — MAILLAGE : RÉDUCTION 70 m → 20 m + RAFFINEMENT
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "MODÉLISATION");
  addSlideTitle(s, "Stratégie de maillage — Réduction du domaine");
  addDivider(s);

  // Top: two mesh images side by side
  // Left: 70m domain (initial)
  card(s, 0.35, 1.42, 4.5, 0.38, C.navy);
  s.addText("Maillage initial — Domaine 70 m  (~150 000 cellules)", {
    x: 0.45, y: 1.42, w: 4.3, h: 0.38,
    fontSize: 11, bold: true, color: C.teal, valign: "middle",
  });
  s.addImage({
    path: FIGDIR + "ch8_domaine_70m.jpeg",
    x: 0.35, y: 1.85, w: 4.5, h: 2.1,
    sizing: { type: "contain", w: 4.5, h: 2.1 },
  });
  s.addText("Domaine complet — premier test", {
    x: 0.35, y: 3.96, w: 4.5, h: 0.22,
    fontSize: 9, color: C.muted, italic: true, align: "center",
  });

  // Arrow between
  s.addText("⟹", {
    x: 4.9, y: 2.65, w: 0.45, h: 0.5,
    fontSize: 22, color: C.gold, bold: true, align: "center", valign: "middle",
  });
  s.addText("Réduction\n70 → 20 m", {
    x: 4.82, y: 3.1, w: 0.6, h: 0.5,
    fontSize: 8, color: C.navy, bold: true, align: "center",
  });

  // Right: 20m domain (refined)
  card(s, 5.15, 1.42, 4.5, 0.38, "065A7A");
  s.addText("Maillage raffiné — Domaine 20 m  (~370 000 cellules)", {
    x: 5.25, y: 1.42, w: 4.3, h: 0.38,
    fontSize: 11, bold: true, color: C.teal, valign: "middle",
  });
  s.addImage({
    path: FIGDIR + "ch8_maillage_raffine.jpeg",
    x: 5.15, y: 1.85, w: 4.5, h: 2.1,
    sizing: { type: "contain", w: 4.5, h: 2.1 },
  });
  s.addText("Zone active — maillage raffiné aux zones critiques", {
    x: 5.15, y: 3.96, w: 4.5, h: 0.22,
    fontSize: 9, color: C.muted, italic: true, align: "center",
  });

  // Bottom: justification cards
  const why = [
    { title: "Pourquoi réduire ?", body: "Le domaine 70 m incluait une grande zone de retenue inutile → temps de calcul excessif et convergence lente" },
    { title: "Domaine retenu : 20 m", body: "Centré sur la zone hydraulique active : seuil + coursier + cuillère + zone de jet aval" },
    { title: "Raffinement", body: "Densification du maillage au niveau du seuil Creager, de la cuillère et de l'interface eau/air" },
  ];
  why.forEach((w, i) => {
    const x = 0.35 + i * 3.15;
    card(s, x, 4.28, 3.0, 0.95, i === 1 ? C.navy : C.ltBlue);
    s.addText(w.title, {
      x: x + 0.1, y: 4.28, w: 2.8, h: 0.32,
      fontSize: 10, bold: true, color: i === 1 ? C.gold : C.navy, valign: "middle",
    });
    s.addText(w.body, {
      x: x + 0.1, y: 4.6, w: 2.8, h: 0.6,
      fontSize: 9, color: i === 1 ? C.white : C.dark, valign: "top",
    });
  });

  s.addNotes("Expliquer la stratégie de maillage en deux étapes : domaine 70m initial → réduction à 20m + raffinement pour optimiser temps de calcul et précision.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 8D — PARAMÈTRES DE SIMULATION (solveur, turbulence k-ω SST, CLs)
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "MODÉLISATION");
  addSlideTitle(s, "Paramètres de la simulation OpenFOAM");
  addDivider(s);

  // Left: setup table
  const setup = [
    ["Solveur", "interFoam — VOF diphasique eau/air"],
    ["Géométrie importée", "Fichier .STL (SketchUp)"],
    ["Modèle de turbulence", "k-ω SST (RAS)"],
    ["Conditions aux limites", "Inlet : vitesse imposée\nOutlet : pression libre\nParois : no-slip"],
    ["Parallélisation", "8 processeurs (decomposePar)"],
    ["Débits simulés", "Q₁₀ / Q₁₀₀ / Q₁₀₀₀ + Q exploratoire"],
  ];

  setup.forEach(([key, val], i) => {
    const y = 1.45 + i * 0.6;
    card(s, 0.35, y, 2.6, 0.5, C.navy);
    s.addText(key, {
      x: 0.35, y, w: 2.6, h: 0.5,
      fontSize: 10, bold: true, color: C.teal,
      align: "center", valign: "middle", margin: 0,
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 2.95, y, w: 3.3, h: 0.5,
      fill: { color: i % 2 === 0 ? C.ltBlue : C.white },
    });
    s.addText(val, {
      x: 3.05, y, w: 3.1, h: 0.5,
      fontSize: 9.5, color: C.dark, valign: "middle",
    });
  });

  // Right: k-ω SST justification box
  card(s, 6.6, 1.45, 3.05, 4.05, C.navy);
  s.addText("Pourquoi k-ω SST ?", {
    x: 6.7, y: 1.5, w: 2.85, h: 0.42,
    fontSize: 13, bold: true, color: C.gold,
    align: "center", valign: "middle",
  });

  // Divider
  s.addShape(pres.shapes.LINE, {
    x: 6.75, y: 1.95, w: 2.75, h: 0,
    line: { color: C.teal, width: 1 },
  });

  const reasons = [
    { icon: "✓", text: "Combine k-ω près des parois et k-ε en zone libre → meilleure précision partout" },
    { icon: "✓", text: "Adapté aux forts gradients de pression adverses (seuil Creager, courbure cuillère)" },
    { icon: "✓", text: "Capture la séparation de l'écoulement à la sortie de la cuillère (formation du jet)" },
    { icon: "✓", text: "k-ε surestime la viscosité turbulente dans les écoulements fortement courbés → inadapté ici" },
    { icon: "✓", text: "Modèle validé dans la littérature pour les évacuateurs de crues (OpenFOAM + interFoam)" },
  ];
  reasons.forEach((r, i) => {
    s.addShape(pres.shapes.OVAL, {
      x: 6.72, y: 2.05 + i * 0.68, w: 0.28, h: 0.28,
      fill: { color: C.teal },
    });
    s.addText(r.icon, {
      x: 6.72, y: 2.05 + i * 0.68, w: 0.28, h: 0.28,
      fontSize: 9, bold: true, color: C.navy,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(r.text, {
      x: 7.07, y: 2.02 + i * 0.68, w: 2.5, h: 0.44,
      fontSize: 9, color: C.white, valign: "middle",
    });
  });

  s.addNotes("Justifier le choix du modèle k-ω SST : meilleur que k-ε pour les écoulements avec gradients de pression et courbure (seuil + cuillère). Validé dans la littérature pour les évacuateurs de crues.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 9 — RÉSULTATS VITESSE
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "RÉSULTATS");
  addSlideTitle(s, "Champs de vitesse — Comparaison des trois débits");
  addDivider(s);

  const imgs = [
    { file: "ch9_q10_U_profil.jpeg",    label: "Q₁₀ = 161 m³/s",    vmax: "V_max = 20,5 m/s" },
    { file: "ch9_q100_U_profil.jpeg",   label: "Q₁₀₀ = 508 m³/s",   vmax: "V_max = 19,3 m/s" },
    { file: "ch9_q1000_U_profil.jpeg",  label: "Q₁₀₀₀ = 795 m³/s",  vmax: "V_max = 20,2 m/s" },
  ];

  imgs.forEach((img, i) => {
    const x = 0.35 + i * 3.2;
    card(s, x, 1.45, 2.9, 3.55, C.white);
    s.addImage({
      path: FIGDIR + img.file,
      x: x + 0.08, y: 1.53, w: 2.74, h: 2.55,
      sizing: { type: "contain", w: 2.74, h: 2.55 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 4.08, w: 2.9, h: 0.5,
      fill: { color: C.navy },
    });
    s.addText(img.label, {
      x, y: 4.08, w: 2.9, h: 0.25,
      fontSize: 11, bold: true, color: C.teal, align: "center", margin: 0,
    });
    s.addText(img.vmax, {
      x, y: 4.33, w: 2.9, h: 0.25,
      fontSize: 10, color: C.gold, align: "center", bold: true, margin: 0,
    });
  });

  s.addText("→  Accélération progressive de la retenue (≈ 0 m/s) jusqu'à la cuillère (≈ 20 m/s) pour les trois débits", {
    x: 0.35, y: 5.1, w: 9.3, h: 0.38,
    fontSize: 11, color: C.navy, italic: true, align: "center",
  });

  s.addNotes("Présenter les champs de vitesse pour les 3 débits et commenter l'accélération de l'écoulement.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 10 — RÉSULTATS PRESSION + JET
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "RÉSULTATS");
  addSlideTitle(s, "Pression & Formation du jet — Q₁₀₀₀ = 795 m³/s");
  addDivider(s);

  // Pressure left
  s.addImage({
    path: FIGDIR + "ch9_q1000_P_profil.jpeg",
    x: 0.35, y: 1.45, w: 4.5, h: 2.6,
    sizing: { type: "contain", w: 4.5, h: 2.6 },
  });
  card(s, 0.35, 4.05, 4.5, 0.55, C.navy);
  s.addText("Champ de pression (Pa) — Q₁₀₀₀\nSurpression en pied de seuil · Dépression en cuillère → formation du jet", {
    x: 0.45, y: 4.05, w: 4.3, h: 0.55,
    fontSize: 9.5, color: C.white, align: "left", valign: "middle",
  });

  // Jet right
  s.addImage({
    path: FIGDIR + "ch9_q1000_jet.png",
    x: 5.15, y: 1.45, w: 4.5, h: 2.6,
    sizing: { type: "contain", w: 4.5, h: 2.6 },
  });
  card(s, 5.15, 4.05, 4.5, 0.55, C.navy);
  s.addText("Jet aérien à la sortie de la cuillère — Q₁₀₀₀\nTrajectoire parabolique · Impact à Lx ≈ 36–38 m", {
    x: 5.25, y: 4.05, w: 4.3, h: 0.55,
    fontSize: 9.5, color: C.white, align: "left", valign: "middle",
  });

  s.addNotes("Commenter le champ de pression et le jet formé en sortie de cuillère pour Q1000.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 11 — AFFOUILLEMENT
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "RÉSULTATS");
  addSlideTitle(s, "Évaluation de l'affouillement à l'aval");
  addDivider(s);

  // Left: fosse image
  s.addImage({
    path: FIGDIR + "fosse_q1000.png",
    x: 0.35, y: 1.45, w: 4.5, h: 2.2,
    sizing: { type: "contain", w: 4.5, h: 2.2 },
  });
  s.addText("Calcul pour Q₁₀₀₀ — 4 formules empiriques", {
    x: 0.35, y: 3.65, w: 4.5, h: 0.28,
    fontSize: 9, color: C.muted, align: "center", italic: true,
  });

  // Right: table
  const tbl = [
    [
      { text: "Formule", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "Q₁₀ · D (m)", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "Q₁₀₀ · D (m)", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "Q₁₀₀₀ · D (m)", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
    ],
    ["Véronèse",  "5,64", "10,20", "13,48"],
    ["Damle",     "2,19",  "4,33",  "5,74"],
    ["Martins",   "3,28",  "6,46",  "8,59"],
    ["Chian",     "3,51",  "6,13",  "8,01"],
    [
      { text: "Moyenne", options: { bold: true, color: C.navy, fill: { color: "D4EAF7" } } },
      { text: "3,65",    options: { bold: true, color: C.navy, fill: { color: "D4EAF7" } } },
      { text: "6,78",    options: { bold: true, color: C.navy, fill: { color: "D4EAF7" } } },
      { text: "8,96",    options: { bold: true, color: C.navy, fill: { color: "D4EAF7" } } },
    ],
  ];

  s.addTable(tbl, {
    x: 5.1, y: 1.45, w: 4.55, h: 2.5,
    colW: [1.4, 1.0, 1.1, 1.05],
    border: { pt: 0.5, color: "C0D8EC" },
    align: "center", fontSize: 11,
  });

  // Key result
  card(s, 5.1, 4.05, 4.55, 0.75, C.navy);
  s.addText("Valeur retenue (Q₁₀₀₀) :", {
    x: 5.2, y: 4.05, w: 4.35, h: 0.32,
    fontSize: 11, color: C.teal, bold: true, valign: "middle",
  });
  s.addText("D = 8,96 m  →  Cote fond fosse = 587,23 m NGM", {
    x: 5.2, y: 4.37, w: 4.35, h: 0.35,
    fontSize: 12, color: C.gold, bold: true, valign: "middle",
  });

  // Bottom note
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.35, y: 5.05, w: 9.3, h: 0.38,
    fill: { color: C.ltBlue }, rectRadius: 0.06,
  });
  s.addText("PHE aval = 596,19 m NGM  ·  Cote fond fosse > Cote fondations barrage  →  Pas de risque pour les fondations", {
    x: 0.45, y: 5.05, w: 9.1, h: 0.38,
    fontSize: 10, color: C.navy, align: "center", valign: "middle",
  });

  s.addNotes("Présenter les résultats de l'affouillement avec les 4 formules et conclure sur la sécurité des fondations.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 12 — COMPARAISON THÉORIE / SIMULATION
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  lightSlide(s);
  addSectionTag(s, "COMPARAISON");
  addSlideTitle(s, "Comparaison théorie / simulation numérique");
  addDivider(s);

  // Hauteurs d'eau table
  s.addText("Hauteurs d'eau — Q₁₀₀₀ = 794,93 m³/s", {
    x: 0.35, y: 1.42, w: 5.5, h: 0.32,
    fontSize: 12, bold: true, color: C.navy,
  });

  const tblH = [
    [
      { text: "Grandeur", options: { bold: true, color: C.white, fill: { color: C.navy } } },
      { text: "Théorique", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "Simulé", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "Écart (%)", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
    ],
    ["Charge seuil Hd (m)",     "2,36", "2,54", "7,6 %"],
    ["Hauteur coursier (m)",    "0,82", "1,12", "36,6 %"],
    ["Hauteur cuillère (m)",    "0,61", "1,01", "65,6 %"],
  ];

  s.addTable(tblH, {
    x: 0.35, y: 1.75, w: 5.5, h: 1.5,
    colW: [2.5, 1.0, 1.0, 1.0],
    border: { pt: 0.5, color: "C0D8EC" },
    align: "center", fontSize: 11,
  });

  // Vitesses table
  s.addText("Vitesses maximales — 3 débits", {
    x: 0.35, y: 3.4, w: 5.5, h: 0.32,
    fontSize: 12, bold: true, color: C.navy,
  });

  const tblV = [
    [
      { text: "Débit", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "V_théo (m/s)", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "V_sim (m/s)", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
      { text: "Écart (%)", options: { bold: true, color: C.white, fill: { color: C.navy }, align: "center" } },
    ],
    ["Q₁₀  (161 m³/s)",   "17,99", "20,5", "13,95 %"],
    ["Q₁₀₀ (508 m³/s)",  "18,00", "19,3",  "7,22 %"],
    ["Q₁₀₀₀ (795 m³/s)", "18,50", "20,2",  "9,19 %"],
  ];

  s.addTable(tblV, {
    x: 0.35, y: 3.73, w: 5.5, h: 1.5,
    colW: [2.2, 1.1, 1.1, 1.1],
    border: { pt: 0.5, color: "C0D8EC" },
    align: "center", fontSize: 11,
  });

  // Right: interpretation cards
  const interp = [
    { title: "Hauteurs d'eau", body: "Écart Hd < 10 % ✓\nCoursier & cuillère : mesure 3D\n≠ section 1D → écart attendu", color: "0E5A8A" },
    { title: "Vitesses maximales", body: "Écarts 7–14 % ✓\nOpenFOAM surestime : V_sim\nest la valeur maximale locale\nvs V_théo moyenne de section", color: C.navy },
    { title: "Conclusion", body: "Cohérence globale confirmée\nEcarts < 14 % : validation\ndu modèle CFD acceptable", color: "065A7A" },
  ];

  interp.forEach((c, i) => {
    card(s, 6.1, 1.45 + i * 1.35, 3.55, 1.2, c.color);
    s.addText(c.title, {
      x: 6.2, y: 1.45 + i * 1.35, w: 3.35, h: 0.38,
      fontSize: 12, bold: true, color: C.gold, valign: "middle",
    });
    s.addText(c.body, {
      x: 6.2, y: 1.83 + i * 1.35, w: 3.35, h: 0.72,
      fontSize: 10, color: C.white, valign: "top",
    });
  });

  s.addNotes("Commenter les écarts et expliquer pourquoi ils sont acceptables pour une comparaison 1D vs 3D.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 13 — CONCLUSION
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkSlide(s);

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 1.1, w: 10, h: 0.05,
    fill: { color: C.teal },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.5, w: 10, h: 0.05,
    fill: { color: C.teal },
  });

  s.addText("Conclusion & Perspectives", {
    x: 0.5, y: 0.3, w: 9, h: 0.72,
    fontSize: 26, bold: true, color: C.white, align: "center",
  });

  const concl = [
    { icon: "✓", text: "Dimensionnement complet de l'évacuateur de crues : seuil Creager, coursier, cuillère saut de ski, murs bajoyers, trajectoire du jet et affouillement — pour Q₁₀, Q₁₀₀ et Q₁₀₀₀." },
    { icon: "✓", text: "Modèle CFD 3D complet sous OpenFOAM (interFoam/VOF) : deux niveaux de maillage, 3 débits simulés + 1 scénario exploratoire (Q = 1 051 m³/s)." },
    { icon: "✓", text: "Validation numérique : écarts théorie/simulation inférieurs à 14 % pour les vitesses et 10 % pour la charge sur le seuil — cohérence confirmée." },
  ];

  concl.forEach((c, i) => {
    s.addShape(pres.shapes.OVAL, {
      x: 0.4, y: 1.3 + i * 0.98, w: 0.45, h: 0.45,
      fill: { color: C.teal },
    });
    s.addText(c.icon, {
      x: 0.4, y: 1.3 + i * 0.98, w: 0.45, h: 0.45,
      fontSize: 14, bold: true, color: C.navy,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(c.text, {
      x: 1.0, y: 1.28 + i * 0.98, w: 8.6, h: 0.5,
      fontSize: 11.5, color: C.white, valign: "middle",
    });
  });

  // Perspectives
  s.addText("Perspectives", {
    x: 0.4, y: 4.62, w: 2, h: 0.38,
    fontSize: 14, bold: true, color: C.gold,
  });

  const persp = [
    "Validation expérimentale sur modèle physique réduit",
    "Simulation de scénarios de surélévation du barrage (Q = 1 051 m³/s étudié)",
    "Optimisation du profil de la cuillère pour réduire l'affouillement",
  ];
  persp.forEach((p, i) => {
    s.addText(`→  ${p}`, {
      x: 0.4, y: 5.05 + i * 0.16, w: 9.2, h: 0.22,
      fontSize: 10, color: "A8C8E0",
    });
  });

  s.addNotes("Conclure sur les apports du PFE et ouvrir sur les perspectives.");
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 14 — MERCI
// ════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkSlide(s);

  s.addShape(pres.shapes.RECTANGLE, {
    x: 2, y: 2.1, w: 6, h: 0.06,
    fill: { color: C.teal },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 2, y: 3.42, w: 6, h: 0.06,
    fill: { color: C.teal },
  });

  s.addText("Merci pour votre attention", {
    x: 0.5, y: 1.0, w: 9, h: 1.0,
    fontSize: 34, bold: true, color: C.white, align: "center", valign: "middle",
  });
  s.addText("Questions & Discussion", {
    x: 0.5, y: 2.25, w: 9, h: 0.55,
    fontSize: 18, color: C.teal, align: "center",
  });
  s.addText("Modélisation numérique 3D de l'évacuateur de crues\ndu barrage Ahl Souss par OpenFOAM", {
    x: 0.5, y: 3.55, w: 9, h: 0.85,
    fontSize: 13, color: "7BADC8", align: "center",
  });
  s.addText("Hiba Asmae BASTOUT  ·  ENSA Agadir  ·  2025–2026", {
    x: 0.5, y: 4.8, w: 9, h: 0.38,
    fontSize: 11, color: C.muted, align: "center",
  });

  s.addNotes("Slide de clôture — inviter les questions.");
}

// ─── WRITE FILE ──────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "/workspace/Presentation_Soutenance_BarrageAhlSouss_v2.pptx" })
  .then(() => console.log("✅  Présentation générée avec succès."))
  .catch(e => console.error("❌  Erreur :", e));
