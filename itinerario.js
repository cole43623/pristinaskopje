const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageBreak, LevelFormat
} = require('docx');
const fs = require('fs');

// Colors
const NAVY = "1A3A5C";
const BLUE = "2E6DA4";
const LIGHT_BLUE = "D6E8F5";
const GOLD = "C8963E";
const LIGHT_GOLD = "FDF3E3";
const GRAY = "666666";
const LIGHT_GRAY = "F5F5F5";
const WHITE = "FFFFFF";

function border(color = "CCCCCC", size = 4) {
  return { style: BorderStyle.SINGLE, size, color };
}
function noBorder() {
  return { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
}
const allBorders = (c, s) => ({ top: border(c,s), bottom: border(c,s), left: border(c,s), right: border(c,s) });
const noAllBorders = () => ({ top: noBorder(), bottom: noBorder(), left: noBorder(), right: noBorder() });

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    children: [new TextRun({ text, bold: true, size: 36, color: WHITE, font: "Arial" })]
  });
}

function heading2(text, color = NAVY) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GOLD, space: 1 } },
    children: [new TextRun({ text, bold: true, size: 28, color, font: "Arial" })]
  });
}

function heading3(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, size: 24, color: BLUE, font: "Arial" })]
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 22, font: "Arial", color: "333333", ...opts })]
  });
}

function bulletItem(text, bold = false) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "Arial", color: "333333", bold })]
  });
}

function spacer(size = 120) {
  return new Paragraph({ spacing: { before: size, after: 0 }, children: [new TextRun("")] });
}

// Banner paragraph with colored background via shading (using table trick)
function colorBanner(text, bgColor = NAVY, textColor = WHITE) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: noAllBorders(),
            shading: { fill: bgColor, type: ShadingType.CLEAR },
            margins: { top: 160, bottom: 160, left: 240, right: 240 },
            width: { size: 9360, type: WidthType.DXA },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text, bold: true, size: 32, color: textColor, font: "Arial" })]
              })
            ]
          })
        ]
      })
    ]
  });
}

function infoBox(leftLabel, leftVal, rightLabel, rightVal, bgColor = LIGHT_BLUE) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1800, 2880, 1800, 2880],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: noAllBorders(),
            shading: { fill: bgColor, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 160, right: 100 },
            width: { size: 1800, type: WidthType.DXA },
            children: [new Paragraph({ children: [new TextRun({ text: leftLabel, bold: true, size: 20, color: NAVY, font: "Arial" })] })]
          }),
          new TableCell({
            borders: noAllBorders(),
            shading: { fill: bgColor, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 100, right: 100 },
            width: { size: 2880, type: WidthType.DXA },
            children: [new Paragraph({ children: [new TextRun({ text: leftVal, size: 20, color: "333333", font: "Arial" })] })]
          }),
          new TableCell({
            borders: noAllBorders(),
            shading: { fill: bgColor, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 100, right: 100 },
            width: { size: 1800, type: WidthType.DXA },
            children: [new Paragraph({ children: [new TextRun({ text: rightLabel, bold: true, size: 20, color: NAVY, font: "Arial" })] })]
          }),
          new TableCell({
            borders: noAllBorders(),
            shading: { fill: bgColor, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 100, right: 160 },
            width: { size: 2880, type: WidthType.DXA },
            children: [new Paragraph({ children: [new TextRun({ text: rightVal, size: 20, color: "333333", font: "Arial" })] })]
          }),
        ]
      })
    ]
  });
}

function dayTable(orario, attivita, note, bgHeader = LIGHT_BLUE) {
  const headerRow = new TableRow({
    children: [
      new TableCell({
        borders: allBorders(BLUE, 4),
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 60 },
        width: { size: 1440, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: "Orario", bold: true, size: 20, color: WHITE, font: "Arial" })] })]
      }),
      new TableCell({
        borders: allBorders(BLUE, 4),
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 60 },
        width: { size: 5280, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: "Attivita'", bold: true, size: 20, color: WHITE, font: "Arial" })] })]
      }),
      new TableCell({
        borders: allBorders(BLUE, 4),
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 60 },
        width: { size: 2640, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: "Note / Costi stimati", bold: true, size: 20, color: WHITE, font: "Arial" })] })]
      }),
    ]
  });

  const dataRows = orario.map((o, i) => {
    const isEven = i % 2 === 0;
    const bg = isEven ? WHITE : LIGHT_GRAY;
    return new TableRow({
      children: [
        new TableCell({
          borders: allBorders("DDDDDD", 2),
          shading: { fill: bg, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 60 },
          width: { size: 1440, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: o, size: 20, bold: true, color: NAVY, font: "Arial" })] })]
        }),
        new TableCell({
          borders: allBorders("DDDDDD", 2),
          shading: { fill: bg, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 60 },
          width: { size: 5280, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: attivita[i], size: 20, color: "333333", font: "Arial" })] })]
        }),
        new TableCell({
          borders: allBorders("DDDDDD", 2),
          shading: { fill: bg, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 60 },
          width: { size: 2640, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: note[i], size: 20, color: GRAY, font: "Arial" })] })]
        }),
      ]
    });
  });

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1440, 5280, 2640],
    rows: [headerRow, ...dataRows]
  });
}

function budgetTable(voci, costi) {
  const headerRow = new TableRow({
    children: [
      new TableCell({
        borders: allBorders(GOLD, 4),
        shading: { fill: GOLD, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 60 },
        width: { size: 6240, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: "Voce di spesa", bold: true, size: 20, color: WHITE, font: "Arial" })] })]
      }),
      new TableCell({
        borders: allBorders(GOLD, 4),
        shading: { fill: GOLD, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 60 },
        width: { size: 3120, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: "Stima a persona (4gg)", bold: true, size: 20, color: WHITE, font: "Arial" })] })]
      }),
    ]
  });

  const dataRows = voci.map((v, i) => {
    const isLast = i === voci.length - 1;
    const bg = isLast ? LIGHT_GOLD : (i % 2 === 0 ? WHITE : LIGHT_GRAY);
    const bold = isLast;
    return new TableRow({
      children: [
        new TableCell({
          borders: allBorders("DDDDDD", 2),
          shading: { fill: bg, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 160, right: 60 },
          width: { size: 6240, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: v, size: 20, bold, color: bold ? NAVY : "333333", font: "Arial" })] })]
        }),
        new TableCell({
          borders: allBorders("DDDDDD", 2),
          shading: { fill: bg, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 160, right: 60 },
          width: { size: 3120, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: costi[i], size: 20, bold, color: bold ? NAVY : "333333", font: "Arial" })] })]
        }),
      ]
    });
  });

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [6240, 3120],
    rows: [headerRow, ...dataRows]
  });
}

function sectionCard(emoji, title, content) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: { top: border(GOLD, 8), bottom: border("DDDDDD", 2), left: border(GOLD, 8), right: border("DDDDDD", 2) },
            shading: { fill: LIGHT_GOLD, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 200, right: 200 },
            width: { size: 9360, type: WidthType.DXA },
            children: [
              new Paragraph({ children: [new TextRun({ text: `${emoji}  ${title}`, bold: true, size: 24, color: NAVY, font: "Arial" })] }),
              new Paragraph({ spacing: { before: 60 }, children: [new TextRun({ text: content, size: 20, color: "333333", font: "Arial" })] }),
            ]
          })
        ]
      })
    ]
  });
}

// ===== DOCUMENT =====
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 600, hanging: 300 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: WHITE },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0, shading: { fill: NAVY, type: ShadingType.CLEAR } }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: NAVY },
        paragraph: { spacing: { before: 280, after: 100 }, outlineLevel: 1 }
      },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children: [

      // ===== COPERTINA =====
      colorBanner("VIAGGIO BALCANI ESTATE 2026", NAVY, WHITE),
      spacer(40),
      colorBanner("PRISTINA  \u2192  SKOPJE", BLUE, WHITE),
      spacer(40),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [4680, 4680],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: noAllBorders(),
                shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
                margins: { top: 160, bottom: 160, left: 240, right: 120 },
                width: { size: 4680, type: WidthType.DXA },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "\uD83D\uDCC5  DATE OPZIONE A", bold: true, size: 22, color: NAVY, font: "Arial" })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60 }, children: [new TextRun({ text: "22 \u2013 25 Luglio 2026", size: 24, bold: true, color: BLUE, font: "Arial" })] }),
                ]
              }),
              new TableCell({
                borders: noAllBorders(),
                shading: { fill: LIGHT_GOLD, type: ShadingType.CLEAR },
                margins: { top: 160, bottom: 160, left: 120, right: 240 },
                width: { size: 4680, type: WidthType.DXA },
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "\uD83D\uDCC5  DATE OPZIONE B", bold: true, size: 22, color: NAVY, font: "Arial" })] }),
                  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60 }, children: [new TextRun({ text: "29 Luglio \u2013 1 Agosto 2026", size: 24, bold: true, color: GOLD, font: "Arial" })] }),
                ]
              }),
            ]
          })
        ]
      }),

      spacer(80),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2340, 2340, 2340, 2340],
        rows: [
          new TableRow({
            children: [
              ["4 Giorni", "\uD83C\uDF0D", "3-4 Persone", "\uD83D\uDCB6"].map((t, i) => new TableCell({
                borders: allBorders("DDDDDD", 2),
                shading: { fill: i % 2 === 0 ? LIGHT_BLUE : LIGHT_GOLD, type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 100, right: 100 },
                width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: t, bold: true, size: 22, color: NAVY, font: "Arial" })] })]
              }))
            ]
          }),
          new TableRow({
            children: [
              ["4 giorni / 3 notti", "Kosovo + Nord Macedonia", "gruppo piccolo", "Budget low-cost"].map((t, i) => new TableCell({
                borders: allBorders("DDDDDD", 2),
                shading: { fill: i % 2 === 0 ? LIGHT_BLUE : LIGHT_GOLD, type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 120, left: 100, right: 100 },
                width: { size: 2340, type: WidthType.DXA },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: t, size: 18, color: GRAY, font: "Arial" })] })]
              }))
            ]
          }),
        ]
      }),

      spacer(120),

      // ===== PANORAMICA =====
      new Paragraph({
        spacing: { before: 0, after: 0 },
        shading: { fill: NAVY, type: ShadingType.CLEAR },
        children: [new TextRun({ text: "  PANORAMICA DEL VIAGGIO", bold: true, size: 28, color: WHITE, font: "Arial" })]
      }),
      spacer(80),

      body("Un viaggio di 4 giorni attraverso due capitali balcaniche tra loro vicine ma storicamente dense: Pristina, la giovane capitale del Kosovo (il paese piu' recente d'Europa), e Skopje, la sorprendente metropoli macedone. Cultura, storia e, se avanza una giornata, natura alle porte della citta'."),
      spacer(60),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [4680, 4680],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: { top: border(BLUE, 8), bottom: border("CCCCCC", 2), left: border("CCCCCC", 2), right: border("CCCCCC", 2) },
                shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 120 },
                width: { size: 4680, type: WidthType.DXA },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "\uD83C\uDFF3\uFE0F  PRISTINA (Giorno 1-2)", bold: true, size: 22, color: NAVY, font: "Arial" })] }),
                  spacer(40),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Newborn Monument & Memorial Heroines", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Cattedrale di Madre Teresa", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Biblioteca Nazionale (architettura iconica)", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Piazza Scanderbeg + Museo Etnologico", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Vita da strada su Mother Teresa Blvd", size: 20, font: "Arial", color: "333333" })] }),
                ]
              }),
              new TableCell({
                borders: { top: border(GOLD, 8), bottom: border("CCCCCC", 2), left: border("CCCCCC", 2), right: border("CCCCCC", 2) },
                shading: { fill: LIGHT_GOLD, type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 120, right: 200 },
                width: { size: 4680, type: WidthType.DXA },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "\uD83C\uDDF2\uD83C\uDDF0  SKOPJE (Giorno 3-4)", bold: true, size: 22, color: NAVY, font: "Arial" })] }),
                  spacer(40),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Bazaar Ottomano (uno dei piu' grandi dei Balcani)", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Stone Bridge + Piazza Macedonia", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Fortezza Kale (vista panoramica, gratuita)", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Museo Indipendenza Macedone + Museo Citta'", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Matka Canyon (gita natura facoltativa)", size: 20, font: "Arial", color: "333333" })] }),
                ]
              }),
            ]
          })
        ]
      }),

      spacer(120),

      // ===== GIORNO 1 =====
      new Paragraph({
        pageBreakBefore: false,
        spacing: { before: 0, after: 0 },
        shading: { fill: NAVY, type: ShadingType.CLEAR },
        children: [new TextRun({ text: "  GIORNO 1 \u2014 PRISTINA & TRASFERIMENTO", bold: true, size: 28, color: WHITE, font: "Arial" })]
      }),
      spacer(40),
      new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "Esplorazione culturale di Pristina e viaggio serale per Skopje", size: 20, color: GRAY, font: "Arial", italics: true })] }),

      dayTable(
        ["Mattina", "Primo pomer.", "Tardo pomer.", "Sera"],
        [
          "Arrivo a Pristina, check-in. Visita del campus universitario: Biblioteca Nazionale (99 cupole), Cattedrale Madre Teresa e la Chiesa Cristo Salvatore.",
          "Passeggiata lungo il Boulevard pedonale Nënë Tereza. Visita al Newborn Monument e al Memoriale Heroinat.",
          "Ingresso nella parte ottomana: Grande Moschea (Xhamia e Madhe) e visita dello splendido Museo Etnografico (Emin Gjiku).",
          "Trasferimento serale in bus per Skopje (~1h45m, frontiera inclusa). Sistemazione in alloggio a Skopje e cena tipica sul lungofiume Vardar."
        ],
        [
          "Campanile Cattedrale: 2 EUR. Biblioteca: gratis.",
          "Gratuito. Molto vivace.",
          "Museo Etnografico: gratuito (consigliata mancia).",
          "Bus ~7 EUR. Cena macedone ~6-10 EUR. Tavče Gravče!"
        ]
      ),

      spacer(60),
      sectionCard("\uD83D\uDCA1", "Consiglio Giorno 1", "Pristina e' una citta' giovane e accogliente. Il Newborn cambia veste grafica ogni anno il 17 febbraio. Prima di salire sul bus per Skopje, non perdetevi un ottimo macchiato in uno dei tantissimi bar del Boulevard (costo ~1 EUR)."),

      spacer(120),

      // ===== GIORNO 2 =====
      new Paragraph({
        spacing: { before: 0, after: 0 },
        shading: { fill: NAVY, type: ShadingType.CLEAR },
        children: [new TextRun({ text: "  GIORNO 2 \u2014 SKOPJE MONUMENTALE & MUSEI", bold: true, size: 28, color: WHITE, font: "Arial" })]
      }),
      spacer(40),
      new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "Piazza Macedonia, memoriali e musei della capitale", size: 20, color: GRAY, font: "Arial", italics: true })] }),

      dayTable(
        ["Mattina", "11:30", "Pomeriggio", "16:00", "Sera"],
        [
          "Passeggiata in Piazza Macedonia per vedere l'imponente statua 'Guerriero a cavallo' (Alessandro Magno) e attraversamento dello storico Stone Bridge (XV sec.).",
          "Visita al toccante Memoriale di Madre Teresa (costruito dove sorgeva la chiesa in cui fu battezzata).",
          "Visita ai resti della Vecchia Stazione Ferroviaria, colpita dal sisma del 1963 (orologio fermo alle 5:17), oggi Museo della Citta'.",
          "Visita all'imponente Museo della Lotta Macedone per l'Indipendenza (statue di cera realistiche e dipinti storici giganteschi).",
          "Cena tradizionale nel quartiere bohemien di Debar Maalo, famoso per i suoi tavoli sotto le fronde, kafane storiche e musica folk dal vivo."
        ],
        [
          "Gratuito.",
          "Ingresso gratuito.",
          "Ingresso gratuito.",
          "Biglietto ~2 EUR (120 MKD). Visita guidata inclusa.",
          "Cena ~6-10 EUR. Ottimi spiedini alla griglia e birra Skopsko."
        ]
      ),

      spacer(60),
      sectionCard("\uD83C\uDF09", "Skopje di notte", "Il centro di Skopje e' illuminato in modo spettacolare di sera: lo Stone Bridge, il Museo Archeologico con le sue enormi colonne neoclassiche riflesse sul fiume Vardar, e le decine di statue dorate. Una passeggiata serale e' obbligatoria."),

      spacer(120),

      // ===== GIORNO 3 =====
      new Paragraph({
        spacing: { before: 0, after: 0 },
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        children: [new TextRun({ text: "  GIORNO 3 \u2014 MILLENNIUM CROSS & CANYON MATKA", bold: true, size: 28, color: WHITE, font: "Arial" })]
      }),
      spacer(40),
      new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "Panorama dal monte Vodno e gita nella natura selvaggia del Canyon", size: 20, color: GRAY, font: "Arial", italics: true })] }),

      dayTable(
        ["Mattina", "Pomeriggio", "Pranzo tardi", "Sera"],
        [
          "Salita sul monte Vodno (in bus 25 o taxi) e funivia per raggiungere la sommità dove si trova la Millennium Cross (croce alta 66 metri).",
          "Spostamento allo spettacolare Matka Canyon. Escursione a piedi sul sentiero scavato nella roccia o noleggio kayak / giro in barca alle grotte di Vrelo.",
          "Pranzo rilassante con vista mozzafiato sul lago Matka, circondati da alte pareti rocciose.",
          "Rientro a Skopje e cena rilassante nei pressi del fiume Vardar."
        ],
        [
          "Funivia ~2 EUR (120 MKD). Vista fantastica.",
          "Taxi a/r ~15-20 EUR (da dividere). Barca ~5 EUR.",
          "Pranzo ~8-12 EUR. Ristorante sul lago molto suggestivo.",
          "Cena ~6-10 EUR/pers. Birra Skopsko o vino locale."
        ]
      ),

      spacer(60),
      sectionCard("\uD83D\uDEA3", "Attivita' al Matka Canyon", "La gita in barca sul lago artificiale di Matka e' il modo migliore per esplorare il canyon. Consente di sbarcare all'ingresso della Grotta Vrelo, caratterizzata da suggestive stallattiti e profondi laghi interni di acqua gelida."),

      spacer(120),

      // ===== GIORNO 4 =====
      new Paragraph({
        spacing: { before: 0, after: 0 },
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        children: [new TextRun({ text: "  GIORNO 4 \u2014 SKOPJE STORICA & PARTENZA", bold: true, size: 28, color: WHITE, font: "Arial" })]
      }),
      spacer(40),
      new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "Bazar ottomano, moschee storiche, fortezza e volo per Roma", size: 20, color: GRAY, font: "Arial", italics: true })] }),

      dayTable(
        ["09:30", "11:30", "13:00", "15:00", "Fine giornata"],
        [
          "Visita al Bazar Ottomano (Čaršija): perdersi tra vicoli lastricati, caravanserragli (Suli An, Kapan An) e bere un tè turco.",
          "Visita alla splendida Moschea di Mustafa Pascià (1492) e all'adiacente Chiesa di San Salvatore (Saint Spas) con iconostasi lignea.",
          "Pranzo tradizionale nel Bazar a base di ćevapi cotti alla brace.",
          "Esplorazione della Fortezza Kale (bastioni bizantini panoramici gratuiti) e Cattedrale bizantina moderna di San Clemente d'Ocrida.",
          "Trasferimento all'aeroporto di Skopje (SKP) per il volo di rientro a Roma."
        ],
        [
          "Gratuito. Tè turco ~0.50 EUR.",
          "Chiesa ~2 EUR (120 MKD). Moschea gratis.",
          "Pranzo da Destan o Kosmos ~5-8 EUR.",
          "Tutto gratuito.",
          "Taxi per aeroporto ~20 EUR (da dividere) o bus di linea."
        ]
      ),

      spacer(60),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [4680, 4680],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: { top: border(BLUE, 6), bottom: border("CCCCCC", 2), left: border("CCCCCC", 2), right: border("CCCCCC", 2) },
                shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 200, right: 120 },
                width: { size: 4680, type: WidthType.DXA },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "⛪ Saint Spas: cosa sapere", bold: true, size: 22, color: NAVY, font: "Arial" })] }),
                  spacer(40),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Chiesa in parte sotterranea (limite altezza ottomano)", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Capolavoro di iconostasi in legno di noce intagliato", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Tomba dell'eroe nazionale Gotse Delchev nel cortile", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Costo ingresso irrisorio (~2 EUR / 120 MKD)", size: 20, font: "Arial", color: "333333" })] }),
                ]
              }),
              new TableCell({
                borders: { top: border(GOLD, 6), bottom: border("CCCCCC", 2), left: border("CCCCCC", 2), right: border("CCCCCC", 2) },
                shading: { fill: LIGHT_GOLD, type: ShadingType.CLEAR },
                margins: { top: 120, bottom: 120, left: 120, right: 200 },
                width: { size: 4680, type: WidthType.DXA },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "Moschea & Old Bazaar", bold: true, size: 22, color: NAVY, font: "Arial" })] }),
                  spacer(40),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Moschea di Mustafa Pascià (1492, imponente cupola)", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Caravanserragli storici ottomani (Suli An, Kapan An)", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Te' turco e baklava nei caffe' tradizionali del bazar", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Molti negozi e banchi sono chiusi di domenica", size: 20, font: "Arial", color: "333333" })] }),
                ]
              }),
            ]
          })
        ]
      }),

      spacer(120),

      // ===== BUDGET =====
      new Paragraph({
        spacing: { before: 0, after: 0 },
        shading: { fill: GOLD, type: ShadingType.CLEAR },
        children: [new TextRun({ text: "  BUDGET STIMATO", bold: true, size: 28, color: WHITE, font: "Arial" })]
      }),
      spacer(40),
      new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "Stime indicative per persona (4 giorni). I prezzi alloggio e trasporti sono da verificare.", size: 20, color: GRAY, font: "Arial", italics: true })] }),

      budgetTable(
        [
          "Trasporto Pristina \u2192 Skopje (bus)",
          "Trasporti locali (autobus urbani, a piedi)",
          "Ingressi musei e attrazioni (tutto incluso)",
          "Pasti: colazioni + pranzi (3-4 EUR/pasto)",
          "Pasti: cene (5-8 EUR/pasto)",
          "Caffe', snack, bibite",
          "Souvenir e extra",
          "TOTALE STIMATO (escluso alloggio)",
        ],
        [
          "Da verificare (~5-10 EUR)",
          "~3-5 EUR in 4 giorni",
          "~5-8 EUR (molti gratuiti!)",
          "~15-20 EUR",
          "~20-30 EUR",
          "~8-12 EUR",
          "~5-10 EUR",
          "~60-90 EUR / persona",
        ]
      ),

      spacer(80),
      sectionCard("\uD83D\uDCB0", "Nota sul budget", "Kosovo e Macedonia del Nord sono tra le destinazioni piu' economiche d'Europa. La maggior parte dei monumenti principali e' gratuita o costa meno di 2 EUR. Mangiare street food e burek mantiene i costi bassissimi. Un budget di 20-25 EUR/giorno (escluso alloggio) e' piu' che realistico."),

      spacer(120),

      // ===== INFO PRATICHE =====
      new Paragraph({
        spacing: { before: 0, after: 0 },
        shading: { fill: NAVY, type: ShadingType.CLEAR },
        children: [new TextRun({ text: "  INFO PRATICHE", bold: true, size: 28, color: WHITE, font: "Arial" })]
      }),
      spacer(40),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [4680, 4680],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: allBorders("DDDDDD", 4),
                shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
                margins: { top: 160, bottom: 160, left: 200, right: 120 },
                width: { size: 4680, type: WidthType.DXA },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "\uD83C\uDFF3\uFE0F KOSOVO", bold: true, size: 22, color: NAVY, font: "Arial" })] }),
                  spacer(60),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Valuta: Euro (EUR)", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Lingua: albanese, serbo", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Documenti: carta d'identita' per cittadini UE", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Luglio: caldo (~30 gradi), soleggiato", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "SIM locale: facilmente reperibile, economica", size: 20, font: "Arial", color: "333333" })] }),
                ]
              }),
              new TableCell({
                borders: allBorders("DDDDDD", 4),
                shading: { fill: LIGHT_GOLD, type: ShadingType.CLEAR },
                margins: { top: 160, bottom: 160, left: 120, right: 200 },
                width: { size: 4680, type: WidthType.DXA },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "\uD83C\uDDF2\uD83C\uDDF0 MACEDONIA DEL NORD", bold: true, size: 22, color: NAVY, font: "Arial" })] }),
                  spacer(60),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Valuta: Denar macedone (MKD). 1 EUR ~ 61 MKD", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Lingua: macedone (cirillico), albanese", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Documenti: carta d'identita' per cittadini UE", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Luglio: caldo (~32 gradi), clima secco", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Cambio: uffici cambio nel Bazar (verificare commissioni)", size: 20, font: "Arial", color: "333333" })] }),
                ]
              }),
            ]
          })
        ]
      }),

      spacer(80),

      // Tips finali
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: { top: border(BLUE, 8), bottom: border("DDDDDD", 2), left: border("DDDDDD", 2), right: border("DDDDDD", 2) },
                shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
                margins: { top: 140, bottom: 140, left: 200, right: 200 },
                width: { size: 9360, type: WidthType.DXA },
                children: [
                  new Paragraph({ children: [new TextRun({ text: "\uD83D\uDCA1 Consigli pratici finali", bold: true, size: 22, color: NAVY, font: "Arial" })] }),
                  spacer(60),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Sia Pristina che Skopje sono citta' molto sicure e accoglienti per i turisti.", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Portare contanti: non tutti i locali small accettano carta.", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Luglio e' piena estate: portare crema solare, bottiglia d'acqua riutilizzabile.", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Prenotare Airbnb/alloggio con anticipo per luglio (alta stagione).", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Per il bus Pristina-Skopje: verificare orari e prezzi aggiornati su Gjirafa.com o direttamente alla stazione.", size: 20, font: "Arial", color: "333333" })] }),
                  new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "Nel bazar di Skopje la domenica molti negozi sono chiusi: meglio visitarlo sabato o nei feriali.", size: 20, font: "Arial", color: "333333" })] }),
                ]
              })
            ]
          })
        ]
      }),

      spacer(120),

      // Footer
      colorBanner("Documento preparato come proposta di viaggio \u2014 Luglio 2026", NAVY, WHITE),
      spacer(20),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 40, after: 40 },
        children: [new TextRun({ text: "Prezzi alloggi, orari bus e tariffe Airbnb da verificare al momento della prenotazione", size: 18, color: GRAY, font: "Arial", italics: true })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/Itinerario_Pristina_Skopje_Luglio2026.docx', buffer);
  console.log('DONE');
});
