// Ferro Tilbud — DOCX generator (runs in browser via docx IIFE)

function b64ToUint8(b64) {
  const bin = atob(b64);
  const u = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
  return u;
}

function fmtE(n) { return Math.round(n).toLocaleString('nb-NO') + ',-'; }
function fmt(n)  { return (Math.round(n/1000)*1000).toLocaleString('nb-NO') + ',-'; }
function fmtD(s) { if (!s) return ''; const [y,m,d] = s.split('-'); return d+'.'+m+'.'+y; }

const FORUTSETNINGER_LINES = [
  'Tiltakshaver st\u00e5r selv for:',
  '- Grunnarbeid og komprimering av byggetomt',
  '- Offentlige gebyrer og s\u00f8knadskostnader',
  '- Fiber inn til bygning',
  '- Dokumentasjon av grunnforhold og uavhengig kontroll',
  '- Brannisolering - tilbys separat ved behov',
  '- Klimagass- og bygningsfysikkrapport - tilbys separat ved behov',
  '- All rivning og klargj\u00f8ring av byggeplass',
  '- Oljeutskiller og tilknyttet r\u00f8rarbeid',
  '- Tr\u00e6r, busker og inventar',
  '',
  'Generelle betingelser:',
  '- Ved tilleggsarbeid: 750,- pr. time for mont\u00f8r, 1\u00a0200,- pr. time for prosjektleder, 15\u00a0% materialp\u00e5slag.',
  '- Mengder gjeldende og reguleres f\u00f8r kontrakt i samr\u00e5d med tiltakshaver.',
  '- Tilbudet skal skriftlig bestilles av kunde.',
  '- Medf\u00f8lgende tegning p\u00e5 st\u00e5lkonstruksjon gjelder for pris. Det forutsettes kontinuerlig montasje.',
  '- Prisene er basert p\u00e5 anbudsdagens materialpriser og l\u00f8nninger med vanlig adgang til forh\u00e5ndsmessig prisjustering.',
  '- Fundamentering dimensjonert for monteringslaster. Setninger fra ferdigstilt grunnarbeid er ikke Ferro sitt ansvar.',
  '- Tiltakshaver s\u00f8rger for fremkommelig vei rundt bygget - minimum 4\u00a0m bredde for kran og transport.',
  '- Tilbudet er gyldig i 14 dager.',
  '- Forutsetter tiltaksklasse 2. Seismikk utelates. Direkte fundamentering 250\u00a0kN/m\u00b2 bruddgrense (fjell og sprengstein). Bygget forutsettes ikke pelet.',
  '- U-verdi tak 0,18 - U-verdi vegg 0,18 - U-verdi glass 1,2.',
  '- Ryddet ut etter eget arbeid - ikke vasket. Offentlig vann/avl\u00f8p og str\u00f8mtilf\u00f8rsel til byggegrop er med i pris.',
];

async function generateTilbudDocx(info, pris, scope, opsList) {
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    ImageRun, Header, AlignmentType, WidthType, BorderStyle, ShadingType,
    VerticalAlign, TabStopType, LevelFormat
  } = window.docx;

  const NAVY = '1B3A6B', BLUE = '4A90C4', HBLUE = '156082', GRAY = '888888', LBKG = 'EEF4FA';
  const nb = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const nbs = { top: nb, bottom: nb, left: nb, right: nb };

  const banner = b64ToUint8(BANNER_B64);
  const logo2  = b64ToUint8(LOGO2_B64);

  const sumEks  = parseFloat((pris.sumEks||'').replace(/\s/g,'').replace(',','.')) || 0;
  const sumEksR = Math.round(sumEks/1000)*1000;
  const mva     = Math.round(sumEksR*0.25);
  const sumInk  = sumEksR + mva;
  const opsVal  = parseFloat((pris.opsjoner||'').replace(/\s/g,'').replace(',','.')) || 0;
  const opsMva  = Math.round(opsVal*0.25);
  const opsInk  = opsVal + opsMva;
  const dato    = fmtD(info.dato);
  const prosjekt = info.prosjekt || '[Prosjektnavn]';

  const header = new Header({ children: [
    new Paragraph({ spacing: { after: 60 }, children: [
      new ImageRun({ type: 'png', data: banner, transformation: { width: 393, height: 43 },
        altText: { title: 'b', description: 'b', name: 'b' } })
    ]}),
    new Table({ width: { size: 8640, type: WidthType.DXA }, columnWidths: [2400, 6240],
      rows: [new TableRow({ children: [
        new TableCell({ width: { size: 2400, type: WidthType.DXA }, borders: nbs,
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [new Paragraph({ children: [new ImageRun({
            type: 'png', data: logo2, transformation: { width: 86, height: 45 },
            altText: { title: 'l', description: 'l', name: 'l' }
          })]})]
        }),
        new TableCell({ width: { size: 6240, type: WidthType.DXA }, borders: nbs,
          margins: { top: 60, bottom: 0, left: 200, right: 0 },
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({ alignment: AlignmentType.RIGHT, children: [
              new TextRun({ text: 'PRISTILBUD', bold: true, color: HBLUE, size: 28 })
            ]}),
            new Paragraph({ alignment: AlignmentType.RIGHT, children: [
              new TextRun({ text: prosjekt + '  \u2022  ' + dato, color: HBLUE, size: 16 })
            ]})
          ]
        })
      ]})]
    })
  ]});

  const lbl = t => new Paragraph({ spacing: { after: 80 }, children: [
    new TextRun({ text: t, bold: true, color: BLUE, size: 14 })
  ]});

  const sh = t => new Paragraph({
    spacing: { before: 200, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 4 } },
    children: [new TextRun({ text: t.toUpperCase(), bold: true, color: NAVY, size: 22 })]
  });

  const st = t => t.split('\n').map((ln, i, a) => new Paragraph({
    spacing: { after: i === a.length - 1 ? 80 : 20 },
    children: [new TextRun({ text: ln, size: 20 })]
  }));

  const forutParagraphs = FORUTSETNINGER_LINES.map(ln => {
    if (ln.startsWith('-')) {
      return new Paragraph({ numbering: { reference: 'fbullets', level: 0 },
        children: [new TextRun({ text: ln.slice(1).trim(), size: 19 })] });
    }
    return new Paragraph({ spacing: { after: ln === '' ? 60 : 20 },
      children: [new TextRun({ text: ln, bold: ln.endsWith(':'), size: 19 })] });
  });

  const scopeContent = (scope || []).filter(s => s.text && s.text.trim())
    .flatMap(s => [sh(s.label), ...st(s.text)]);

  let opsContent = [];
  if ((opsList && opsList.length > 0) || opsVal > 0) {
    opsContent = [
      new Paragraph({ spacing: { before: 200, after: 60 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 4 } },
        children: [new TextRun({ text: 'OPSJONER (EKS. MVA)', bold: true, color: NAVY, size: 22 })]
      }),
      ...(opsList || []).map(o => new Paragraph({ spacing: { after: 30 },
        children: [new TextRun({ text: o, size: 19 })] })),
      ...(opsVal > 0 ? [
        new Paragraph({ spacing: { before: 60, after: 20 }, children: [new TextRun({ text: 'Sum opsjoner eks. mva:   ' + fmtE(opsVal), size: 19 })] }),
        new Paragraph({ spacing: { after: 20 }, children: [new TextRun({ text: 'Mva (25 %):              ' + fmtE(opsMva), size: 19 })] }),
        new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: 'Sum opsjoner ink. mva:   ' + fmtE(opsInk), bold: true, size: 19 })] }),
      ] : [])
    ];
  }

  const doc = new Document({
    numbering: { config: [{
      reference: 'fbullets',
      levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2013', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 480, hanging: 240 } } } }]
    }]},
    sections: [{
      properties: { page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1600, right: 1100, bottom: 1100, left: 1100, header: 708, footer: 708 }
      }},
      headers: { default: header },
      children: [
        new Table({ width: { size: 8640, type: WidthType.DXA }, columnWidths: [3800, 4840],
          rows: [new TableRow({ children: [
            new TableCell({ width: { size: 3800, type: WidthType.DXA }, borders: nbs,
              margins: { top: 160, bottom: 160, left: 0, right: 200 }, children: [
              lbl('TILBUD TIL'),
              new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: info.kunde || '[Kunde]', bold: true, color: NAVY, size: 28 })] }),
              ...(info.kontakt ? [new Paragraph({ spacing: { after: 30 }, children: [new TextRun({ text: 'v/ ' + info.kontakt, size: 20 })] })] : []),
              new Paragraph({ spacing: { after: 30 }, children: [new TextRun({ text: info.adresse || '', size: 20 })] }),
              new Paragraph({ spacing: { after: 30 }, children: [new TextRun({ text: ((info.postnr || '') + ' ' + (info.sted || '')).trim(), size: 20 })] })
            ]}),
            new TableCell({ width: { size: 4840, type: WidthType.DXA },
              borders: { top: { style: BorderStyle.SINGLE, size: 18, color: BLUE }, bottom: nb, left: nb, right: nb },
              shading: { fill: LBKG, type: ShadingType.CLEAR },
              margins: { top: 160, bottom: 160, left: 280, right: 280 }, children: [
              lbl('PRISSAMMENDRAG'),
              new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 4500 }], spacing: { before: 50, after: 50 },
                children: [new TextRun({ text: 'Sum uten opsjoner eks. mva' }), new TextRun({ text: '\t' + (sumEksR ? fmt(sumEksR) : '000 000,-'), color: '333333' })] }),
              new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 4500 }], spacing: { before: 50, after: 50 },
                children: [new TextRun({ text: 'Mva (25 %)' }), new TextRun({ text: '\t' + (sumEksR ? fmtE(mva) : '000 000,-'), color: '333333' })] }),
              new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 6 } },
                tabStops: [{ type: TabStopType.RIGHT, position: 4100 }], spacing: { before: 80, after: 40 },
                children: [new TextRun({ text: 'Sum ink. mva', bold: true, color: NAVY, size: 26 }), new TextRun({ text: '\t' + (sumEksR ? fmtE(sumInk) : '000 000,-'), bold: true, color: NAVY, size: 26 })] }),
              new Paragraph({ spacing: { after: 20 } }),
              new Paragraph({ children: [new TextRun({ text: 'Gyldig i 14 dager fra ' + dato, italics: true, color: GRAY, size: 17 })] })
            ]})
          ]})]
        }),
        new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: 'Vi takker for deres foresp\u00f8rsel og har gleden av \u00e5 tilby dere f\u00f8lgende:', size: 20 })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: 'I vedleggene ligger v\u00e5r beskrivelse av arbeidet. Vi h\u00e5per tilbudet er tilfredsstillende og at deres \u00f8nsker har blitt ivaretatt. H\u00e5per dere aksepterer tilbudet og tar kontakt ved sp\u00f8rsm\u00e5l.', size: 20 })] }),
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: 'Med vennlig hilsen', size: 20 })] }),
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: 'Vegard Landsdal', bold: true, size: 20 })] }),
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: 'Daglig leder, Ferro St\u00e5lentrepren\u00f8r AS', size: 20 })] }),
        new Paragraph({ spacing: { after: 320 }, children: [new TextRun({ text: 'vegard@ferrostal.no  \u2022  Tlf: 95 83 71 23', size: 20 })] }),
        new Paragraph({ spacing: { before: 0, after: 120 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 4 } },
          children: [new TextRun({ text: 'Arbeidsbeskrivelse', bold: true, color: NAVY, size: 28 })] }),
        ...scopeContent,
        ...opsContent,
        new Paragraph({ spacing: { before: 300, after: 120 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 4 } },
          children: [new TextRun({ text: 'Generelle Forutsetninger', bold: true, color: NAVY, size: 24 })] }),
        ...forutParagraphs
      ]
    }]
  });

  return Packer.toBlob(doc);
}
