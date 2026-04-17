const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak
} = require("docx");

// Colors
const PRIMARY = "1B4F72";
const ACCENT = "2E86C1";
const LIGHT_BG = "EBF5FB";
const DARK_BG = "D4E6F1";
const TEXT_DARK = "2C3E50";
const GRAY = "7F8C8D";

const border = { style: BorderStyle.SINGLE, size: 1, color: "BDC3C7" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0 };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, bold: true, size: 32, font: "Arial", color: PRIMARY })],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, bold: true, size: 26, font: "Arial", color: ACCENT })],
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
    children: [new TextRun({ text, size: 22, font: "Arial", color: TEXT_DARK, ...opts })],
  });
}

function bulletItem(text, boldPrefix) {
  const children = [];
  if (boldPrefix) {
    children.push(new TextRun({ text: boldPrefix, size: 22, font: "Arial", color: TEXT_DARK, bold: true }));
    children.push(new TextRun({ text, size: 22, font: "Arial", color: TEXT_DARK }));
  } else {
    children.push(new TextRun({ text, size: 22, font: "Arial", color: TEXT_DARK }));
  }
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 60 },
    children,
  });
}

function tableRow(cells, isHeader = false) {
  return new TableRow({
    children: cells.map((cell, i) => {
      const widths = [6000, 3360];
      return new TableCell({
        borders,
        width: { size: widths[i], type: WidthType.DXA },
        margins: cellMargins,
        shading: isHeader
          ? { fill: PRIMARY, type: ShadingType.CLEAR }
          : { fill: i === 0 ? LIGHT_BG : "FFFFFF", type: ShadingType.CLEAR },
        children: [new Paragraph({
          children: [new TextRun({
            text: cell, size: 20, font: "Arial",
            color: isHeader ? "FFFFFF" : TEXT_DARK,
            bold: isHeader,
          })],
        })],
      });
    }),
  });
}

function divider() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: DARK_BG, space: 1 } },
    children: [],
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: PRIMARY },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: ACCENT },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  sections: [
    // --- COVER PAGE ---
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children: [
        new Paragraph({ spacing: { before: 3000 } }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: "PROPOSTA COMERCIAL", size: 44, bold: true, font: "Arial", color: PRIMARY })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: ACCENT, space: 8 } },
          children: [],
        }),
        new Paragraph({ spacing: { before: 200 } }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({ text: "Ferramenta de Diagn\u00F3stico e A\u00E7\u00E3o Operacional", size: 34, font: "Arial", color: ACCENT })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "Gest\u00E3o Inteligente de Estoque e Performance por Loja", size: 24, font: "Arial", color: GRAY })],
        }),
        new Paragraph({ spacing: { before: 2400 } }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "Preparado por: Lucas Lopes", size: 22, font: "Arial", color: TEXT_DARK })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "Data: Abril de 2026", size: 22, font: "Arial", color: GRAY })],
        }),
      ],
    },

    // --- CONTENT ---
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: DARK_BG, space: 4 } },
            children: [new TextRun({ text: "Proposta \u2014 Ferramenta de Diagn\u00F3stico Operacional", size: 16, font: "Arial", color: GRAY, italics: true })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 1, color: DARK_BG, space: 4 } },
            children: [
              new TextRun({ text: "P\u00E1gina ", size: 16, font: "Arial", color: GRAY }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: GRAY }),
            ],
          })],
        }),
      },
      children: [
        // 1. CONTEXTO
        heading1("1. O que \u00E9 esta proposta"),
        para("Esta proposta apresenta o desenvolvimento de uma ferramenta que vai ajudar a sua opera\u00E7\u00E3o a tomar decis\u00F5es melhores e mais r\u00E1pidas sobre o estoque de cada loja."),
        para("Hoje, muitas empresas t\u00EAm os dados de vendas e estoque, mas n\u00E3o conseguem transform\u00E1-los em a\u00E7\u00F5es pr\u00E1ticas com agilidade. A ferramenta proposta resolve exatamente isso: ela analisa os n\u00FAmeros e entrega recomenda\u00E7\u00F5es claras e prontas para serem executadas pelo time de cada loja."),
        para("Em vez de apenas mostrar gr\u00E1ficos e relat\u00F3rios, a ferramenta vai dizer, de forma direta, o que precisa ser feito em cada loja e com cada produto."),

        divider(),

        // 2. PROBLEMA
        heading1("2. Qual problema ela resolve"),
        para("No dia a dia do varejo, \u00E9 comum enfrentar situa\u00E7\u00F5es como:"),
        bulletItem("Produto vendendo bem em uma loja, mas faltando na prateleira (ruptura)"),
        bulletItem("Produto parado no estoque de uma loja sem ningu\u00E9m comprar (excesso)"),
        bulletItem("Falta de visibilidade sobre quais lojas precisam de aten\u00E7\u00E3o urgente"),
        bulletItem("Dificuldade em decidir onde redirecionar estoque ou aplicar promo\u00E7\u00F5es"),
        bulletItem("Recomenda\u00E7\u00F5es que chegam tarde demais para o time de loja agir"),
        new Paragraph({ spacing: { before: 120 } }),
        para("A ferramenta elimina esse \u201Cachismo\u201D e transforma dados que a empresa j\u00E1 tem em a\u00E7\u00F5es concretas e priorizadas."),

        divider(),

        // 3. COMO FUNCIONA
        heading1("3. Como a ferramenta funciona"),
        para("De forma simples, a ferramenta funciona em 5 passos:"),

        heading2("Passo 1 \u2014 Leitura dos dados"),
        para("A ferramenta recebe duas informa\u00E7\u00F5es b\u00E1sicas que a empresa j\u00E1 possui:"),
        bulletItem("Vendas por produto e por loja", "Vendas: "),
        bulletItem("Estoque atual por produto e por loja", "Estoque: "),
        para("Com apenas esses dois dados, j\u00E1 \u00E9 poss\u00EDvel fazer uma an\u00E1lise muito rica."),

        heading2("Passo 2 \u2014 Diagn\u00F3stico autom\u00E1tico"),
        para("A ferramenta cruza vendas com estoque e classifica cada produto em cada loja automaticamente. Por exemplo:"),

        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [6000, 3360],
          rows: [
            tableRow(["Situa\u00E7\u00E3o identificada", "O que significa"], true),
            tableRow(["Ruptura", "O produto acabou e est\u00E1 vendendo bem"]),
            tableRow(["Estoque baixo", "Est\u00E1 prestes a acabar"]),
            tableRow(["Estoque saud\u00E1vel", "Equil\u00EDbrio entre venda e estoque"]),
            tableRow(["Estoque alto / excesso", "Muito estoque parado sem vender"]),
            tableRow(["Sem venda / sem giro", "O produto n\u00E3o est\u00E1 saindo"]),
          ],
        }),

        new Paragraph({ spacing: { before: 160 } }),

        heading2("Passo 3 \u2014 Prioriza\u00E7\u00E3o"),
        para("Nem tudo precisa ser resolvido ao mesmo tempo. A ferramenta organiza as situa\u00E7\u00F5es por ordem de urg\u00EAncia e impacto, destacando:"),
        bulletItem("Quais lojas precisam de aten\u00E7\u00E3o imediata"),
        bulletItem("Quais produtos t\u00EAm maior impacto no resultado"),
        bulletItem("Onde est\u00E1 o maior risco de perda de vendas"),

        heading2("Passo 4 \u2014 Recomenda\u00E7\u00F5es de a\u00E7\u00E3o"),
        para("A ferramenta gera tr\u00EAs tipos de recomenda\u00E7\u00F5es pr\u00E1ticas:"),

        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2800, 6560],
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders, width: { size: 2800, type: WidthType.DXA }, margins: cellMargins,
                  shading: { fill: PRIMARY, type: ShadingType.CLEAR },
                  children: [new Paragraph({ children: [new TextRun({ text: "Tipo de a\u00E7\u00E3o", size: 20, font: "Arial", color: "FFFFFF", bold: true })] })],
                }),
                new TableCell({
                  borders, width: { size: 6560, type: WidthType.DXA }, margins: cellMargins,
                  shading: { fill: PRIMARY, type: ShadingType.CLEAR },
                  children: [new Paragraph({ children: [new TextRun({ text: "Quando \u00E9 recomendada", size: 20, font: "Arial", color: "FFFFFF", bold: true })] })],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders, width: { size: 2800, type: WidthType.DXA }, margins: cellMargins,
                  shading: { fill: LIGHT_BG, type: ShadingType.CLEAR },
                  children: [new Paragraph({ children: [new TextRun({ text: "Reposi\u00E7\u00E3o", size: 20, font: "Arial", color: TEXT_DARK, bold: true })] })],
                }),
                new TableCell({
                  borders, width: { size: 6560, type: WidthType.DXA }, margins: cellMargins,
                  children: [new Paragraph({ children: [new TextRun({ text: "Quando o produto vende bem mas o estoque est\u00E1 acabando. Sugere enviar mais unidades para a loja.", size: 20, font: "Arial", color: TEXT_DARK })] })],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders, width: { size: 2800, type: WidthType.DXA }, margins: cellMargins,
                  shading: { fill: LIGHT_BG, type: ShadingType.CLEAR },
                  children: [new Paragraph({ children: [new TextRun({ text: "Promo\u00E7\u00E3o / Markdown", size: 20, font: "Arial", color: TEXT_DARK, bold: true })] })],
                }),
                new TableCell({
                  borders, width: { size: 6560, type: WidthType.DXA }, margins: cellMargins,
                  children: [new Paragraph({ children: [new TextRun({ text: "Quando h\u00E1 muito estoque parado e as vendas est\u00E3o fracas. Sugere reduzir pre\u00E7o para girar o produto.", size: 20, font: "Arial", color: TEXT_DARK })] })],
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  borders, width: { size: 2800, type: WidthType.DXA }, margins: cellMargins,
                  shading: { fill: LIGHT_BG, type: ShadingType.CLEAR },
                  children: [new Paragraph({ children: [new TextRun({ text: "Transfer\u00EAncia", size: 20, font: "Arial", color: TEXT_DARK, bold: true })] })],
                }),
                new TableCell({
                  borders, width: { size: 6560, type: WidthType.DXA }, margins: cellMargins,
                  children: [new Paragraph({ children: [new TextRun({ text: "Quando uma loja tem excesso e outra est\u00E1 com falta do mesmo produto. Sugere mover estoque entre lojas.", size: 20, font: "Arial", color: TEXT_DARK })] })],
                }),
              ],
            }),
          ],
        }),

        new Paragraph({ spacing: { before: 160 } }),

        heading2("Passo 5 \u2014 Sa\u00EDda pronta para o time de loja"),
        para("A ferramenta entrega tudo organizado e pronto para a\u00E7\u00E3o. O time de loja recebe instru\u00E7\u00F5es claras e diretas, como:"),
        bulletItem("\u201CRepor o produto X nesta loja \u2014 estoque cr\u00EDtico\u201D"),
        bulletItem("\u201CAplicar promo\u00E7\u00E3o no produto Y \u2014 estoque parado h\u00E1 muito tempo\u201D"),
        bulletItem("\u201CTransferir 50 unidades do produto Z da Loja A para a Loja B\u201D"),
        bulletItem("\u201CManter posi\u00E7\u00E3o atual \u2014 estoque saud\u00E1vel\u201D"),
        para("Antes de qualquer a\u00E7\u00E3o ser executada, o gestor pode revisar e aprovar as recomenda\u00E7\u00F5es."),

        divider(),

        // 4. O QUE A FERRAMENTA ENTREGA
        heading1("4. O que a ferramenta entrega"),
        para("A cada an\u00E1lise, a ferramenta gera os seguintes relat\u00F3rios:"),

        heading2("a) Resumo Executivo"),
        para("Uma vis\u00E3o geral r\u00E1pida com os principais achados: quais lojas est\u00E3o mais cr\u00EDticas, onde h\u00E1 maior risco e quais s\u00E3o as oportunidades priorit\u00E1rias."),

        heading2("b) Ranking de Oportunidades"),
        para("Uma lista ordenada por prioridade, mostrando onde agir primeiro para ter o maior resultado com o menor esfor\u00E7o."),

        heading2("c) Lista de A\u00E7\u00F5es por Loja e por Produto"),
        para("Para cada loja e cada produto, a ferramenta informa:"),
        bulletItem("Qual a\u00E7\u00E3o tomar (repor, promover, transferir, manter)"),
        bulletItem("Por que essa a\u00E7\u00E3o \u00E9 recomendada"),
        bulletItem("Qual a prioridade (alta, m\u00E9dia, baixa)"),

        divider(),

        // 5. PERGUNTAS QUE A FERRAMENTA RESPONDE
        heading1("5. Perguntas que a ferramenta responde"),
        para("Com a ferramenta, ser\u00E1 poss\u00EDvel responder rapidamente perguntas como:"),
        bulletItem("Qual loja est\u00E1 com mais produtos em falta?"),
        bulletItem("Quais lojas t\u00EAm estoque encalhado?"),
        bulletItem("Quais produtos precisam ser repostos com urg\u00EAncia?"),
        bulletItem("Quais itens deveriam entrar em promo\u00E7\u00E3o?"),
        bulletItem("De onde posso transferir estoque para onde est\u00E1 faltando?"),
        bulletItem("Quais a\u00E7\u00F5es t\u00EAm maior impacto agora?"),

        divider(),

        // 6. FASES DO PROJETO
        heading1("6. Fases do projeto"),

        heading2("Fase 1 \u2014 Vers\u00E3o Inicial (MVP)"),
        para("A primeira vers\u00E3o da ferramenta j\u00E1 entrega valor imediato, trabalhando com:"),
        bulletItem("Dados de vendas por produto e loja"),
        bulletItem("Dados de estoque por produto e loja"),
        bulletItem("C\u00E1lculo autom\u00E1tico de cobertura e giro"),
        bulletItem("Classifica\u00E7\u00E3o de cada produto (ruptura, excesso, saud\u00E1vel, etc.)"),
        bulletItem("Ranking de oportunidades por prioridade"),
        bulletItem("Recomenda\u00E7\u00F5es de a\u00E7\u00E3o prontas para valida\u00E7\u00E3o"),

        new Paragraph({ spacing: { before: 200 } }),

        heading2("Fase 2 \u2014 Evolu\u00E7\u00F5es Futuras"),
        para("Ap\u00F3s a primeira vers\u00E3o estar rodando, a ferramenta pode ser ampliada com funcionalidades adicionais, como:"),
        bulletItem("An\u00E1lise de margem de lucro por produto"),
        bulletItem("Tempo que o produto est\u00E1 parado no estoque (aging)"),
        bulletItem("Agrupamento de lojas por perfil de venda (clusters)"),
        bulletItem("Sensibilidade do produto a mudan\u00E7as de pre\u00E7o"),
        bulletItem("Hist\u00F3rico de transfer\u00EAncias anteriores"),
        bulletItem("Taxa de convers\u00E3o de estoque em venda (sell-through)"),

        divider(),

        // 7. BENEFÍCIOS
        heading1("7. Benef\u00EDcios esperados"),

        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [4680, 4680],
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders, width: { size: 4680, type: WidthType.DXA }, margins: cellMargins,
                  shading: { fill: PRIMARY, type: ShadingType.CLEAR },
                  children: [new Paragraph({ children: [new TextRun({ text: "Antes (sem a ferramenta)", size: 20, font: "Arial", color: "FFFFFF", bold: true })] })],
                }),
                new TableCell({
                  borders, width: { size: 4680, type: WidthType.DXA }, margins: cellMargins,
                  shading: { fill: PRIMARY, type: ShadingType.CLEAR },
                  children: [new Paragraph({ children: [new TextRun({ text: "Depois (com a ferramenta)", size: 20, font: "Arial", color: "FFFFFF", bold: true })] })],
                }),
              ],
            }),
            ...[
              ["An\u00E1lise manual e demorada", "Diagn\u00F3stico autom\u00E1tico e instant\u00E2neo"],
              ["Decis\u00F5es baseadas em intui\u00E7\u00E3o", "Decis\u00F5es baseadas em dados"],
              ["Rupturas descobertas tarde demais", "Alertas antecipados de ruptura"],
              ["Estoque parado sem visibilidade", "Identifica\u00E7\u00E3o autom\u00E1tica de excesso"],
              ["Transfer\u00EAncias sem crit\u00E9rio claro", "Sugest\u00F5es inteligentes de redistribui\u00E7\u00E3o"],
              ["Time de loja sem dire\u00E7\u00E3o clara", "A\u00E7\u00F5es prontas e priorizadas para cada loja"],
            ].map(([before, after]) =>
              new TableRow({
                children: [
                  new TableCell({
                    borders, width: { size: 4680, type: WidthType.DXA }, margins: cellMargins,
                    children: [new Paragraph({ children: [new TextRun({ text: before, size: 20, font: "Arial", color: TEXT_DARK })] })],
                  }),
                  new TableCell({
                    borders, width: { size: 4680, type: WidthType.DXA }, margins: cellMargins,
                    shading: { fill: LIGHT_BG, type: ShadingType.CLEAR },
                    children: [new Paragraph({ children: [new TextRun({ text: after, size: 20, font: "Arial", color: TEXT_DARK })] })],
                  }),
                ],
              })
            ),
          ],
        }),

        divider(),

        // 8. PRÓXIMOS PASSOS
        heading1("8. Pr\u00F3ximos passos"),
        para("Para darmos in\u00EDcio ao projeto, os passos sugeridos s\u00E3o:"),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 80 },
          children: [
            new TextRun({ text: "Alinhamento inicial: ", size: 22, font: "Arial", color: TEXT_DARK, bold: true }),
            new TextRun({ text: "reuni\u00E3o para entender o formato atual dos dados (vendas e estoque) e validar as prioridades.", size: 22, font: "Arial", color: TEXT_DARK }),
          ],
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 80 },
          children: [
            new TextRun({ text: "Acesso aos dados: ", size: 22, font: "Arial", color: TEXT_DARK, bold: true }),
            new TextRun({ text: "recebimento de uma amostra dos dados de vendas e estoque por produto e loja.", size: 22, font: "Arial", color: TEXT_DARK }),
          ],
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 80 },
          children: [
            new TextRun({ text: "Desenvolvimento da Fase 1 (MVP): ", size: 22, font: "Arial", color: TEXT_DARK, bold: true }),
            new TextRun({ text: "constru\u00E7\u00E3o da primeira vers\u00E3o funcional da ferramenta.", size: 22, font: "Arial", color: TEXT_DARK }),
          ],
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 80 },
          children: [
            new TextRun({ text: "Valida\u00E7\u00E3o conjunta: ", size: 22, font: "Arial", color: TEXT_DARK, bold: true }),
            new TextRun({ text: "apresenta\u00E7\u00E3o dos resultados para ajustes e aprova\u00E7\u00E3o.", size: 22, font: "Arial", color: TEXT_DARK }),
          ],
        }),
        new Paragraph({
          numbering: { reference: "numbers", level: 0 },
          spacing: { after: 80 },
          children: [
            new TextRun({ text: "Implanta\u00E7\u00E3o: ", size: 22, font: "Arial", color: TEXT_DARK, bold: true }),
            new TextRun({ text: "coloca\u00E7\u00E3o da ferramenta em uso com o time de opera\u00E7\u00E3o.", size: 22, font: "Arial", color: TEXT_DARK }),
          ],
        }),

        new Paragraph({ spacing: { before: 400 } }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          border: { top: { style: BorderStyle.SINGLE, size: 2, color: ACCENT, space: 8 } },
          children: [],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "Estou \u00E0 disposi\u00E7\u00E3o para alinharmos os detalhes e darmos in\u00EDcio ao projeto.", size: 22, font: "Arial", color: TEXT_DARK, italics: true })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 40 },
          children: [new TextRun({ text: "Lucas Lopes", size: 24, font: "Arial", color: PRIMARY, bold: true })],
        }),
      ],
    },
  ],
});

const OUTPUT = "/Users/lucaslopes/VS Code/Proposta - Rafael/Proposta - Ferramenta de Diagnostico Operacional.docx";
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(OUTPUT, buffer);
  console.log("Proposta criada com sucesso:", OUTPUT);
});
