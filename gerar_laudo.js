#!/usr/bin/env node
// PeríciaLab — Gerador de Laudo DOCX
// Uso: node gerar_laudo.js dados.json saida.docx
'use strict';

const fs   = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, HeadingLevel, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak,
  TabStopType, TabStopPosition, LineRuleType, NumberFormat,
  convertInchesToTwip, LevelFormat,
} = require('docx');

// ── CONSTANTES DE LAYOUT ─────────────────────────────────────────
// A4: 11906 × 16838 DXA  |  1 inch = 1440 DXA  |  1 cm = 567 DXA
const A4_W      = 11906;
const A4_H      = 16838;
const MARGIN    = 1134;           // 2 cm
const CONTENT_W = A4_W - MARGIN * 2; // 9638 DXA ≈ 16,96 cm

// ── FONTES ───────────────────────────────────────────────────────
const F_NORMAL  = 'Arial';
const F_CIT     = 'Bahnschrift SemiLight';
const F_CIT_B   = 'Bahnschrift SemiBold';
const SZ_NORMAL = 24;   // 12pt  (docx: half-points)
const SZ_CIT    = 22;   // 11pt
const SZ_TITULO = 24;   // 12pt bold
const SZ_SMALL  = 20;   // 10pt

// ── ESPAÇAMENTOS ─────────────────────────────────────────────────
// Texto normal: 1,5 linha; antes/depois 0,10cm = 57 DXA
const SP_NORMAL = { line: 360, lineRule: LineRuleType.AUTO, before: 57, after: 57 };
// Citação: 1,15 linha; mesmo espaçamento
const SP_CIT    = { line: 276, lineRule: LineRuleType.AUTO, before: 57, after: 57 };
// Título de seção: espaço maior acima
const SP_TITULO = { line: 360, lineRule: LineRuleType.AUTO, before: 200, after: 100 };

// ── BORDAS TABELA ─────────────────────────────────────────────────
const borderThin  = { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' };
const borderNone  = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const cellBorders = { top: borderThin, bottom: borderThin, left: borderThin, right: borderThin };
const noBorders   = { top: borderNone, bottom: borderNone, left: borderNone, right: borderNone };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

// ════════════════════════════════════════════════════════════════
// PARSERS HTML → paragraphs docx
// ════════════════════════════════════════════════════════════════
function parseHTMLtoParagraphs(html, isCitacao = false) {
  if (!html) return [];
  const results = [];

  // Divide em blocos: <blockquote>, <p>, <table>
  // Usa uma regex simples para separar blocos principais
  const blocks = splitBlocks(html);

  for (const block of blocks) {
    if (block.type === 'blockquote') {
      const paras = htmlBlockToParagraphs(block.content, true);
      results.push(...paras);
    } else if (block.type === 'table') {
      const tbl = htmlTableToDocx(block.content);
      if (tbl) results.push(tbl);
    } else {
      const paras = htmlBlockToParagraphs(block.content, isCitacao);
      results.push(...paras);
    }
  }
  return results.length ? results : [emptyPara()];
}

function splitBlocks(html) {
  const blocks = [];
  let remaining = html.replace(/\r\n/g, '\n').trim();

  const BQ_RE    = /<blockquote[^>]*>([\s\S]*?)<\/blockquote>/gi;
  const TABLE_RE = /<table[^>]*>([\s\S]*?)<\/table>/gi;

  // Merge both patterns, find all matches in order
  const allMatches = [];
  let m;
  const bqRe = /<blockquote[^>]*>([\s\S]*?)<\/blockquote>/gi;
  while ((m = bqRe.exec(html)) !== null)
    allMatches.push({ start: m.index, end: m.index + m[0].length, type: 'blockquote', content: m[1] });
  const tblRe = /<table[^>]*>([\s\S]*?)<\/table>/gi;
  while ((m = tblRe.exec(html)) !== null)
    allMatches.push({ start: m.index, end: m.index + m[0].length, type: 'table', content: m[0] });

  allMatches.sort((a, b) => a.start - b.start);

  let cursor = 0;
  for (const match of allMatches) {
    if (match.start > cursor) {
      const text = html.slice(cursor, match.start).trim();
      if (text) blocks.push({ type: 'text', content: text });
    }
    blocks.push({ type: match.type, content: match.content });
    cursor = match.end;
  }
  if (cursor < html.length) {
    const text = html.slice(cursor).trim();
    if (text) blocks.push({ type: 'text', content: text });
  }

  return blocks.length ? blocks : [{ type: 'text', content: html }];
}

function htmlBlockToParagraphs(html, isCitacao) {
  // Extrai <p> ou usa texto direto
  const paras = [];
  const pRe = /<p[^>]*>([\s\S]*?)<\/p>/gi;
  let m, lastIdx = 0, found = false;

  while ((m = pRe.exec(html)) !== null) {
    found = true;
    const inner = m[1];
    paras.push(makeParagraph(inner, isCitacao));
    lastIdx = m.index + m[0].length;
  }

  if (!found) {
    // Sem tags <p> — usa o texto direto
    const text = stripTags(html).trim();
    if (text) paras.push(makeParagraph(html, isCitacao));
  }

  return paras.length ? paras : [];
}

function makeParagraph(html, isCitacao) {
  const runs = parseRuns(html, isCitacao);
  if (isCitacao) {
    return new Paragraph({
      children: runs,
      indent: { left: Math.round(4 * 567) }, // 4cm
      spacing: SP_CIT,
      style: 'CitacaoABNT',
    });
  }
  return new Paragraph({
    children: runs,
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: Math.round(2 * 567) }, // 2cm primeira linha
    spacing: SP_NORMAL,
  });
}

function parseRuns(html, isCitacao) {
  if (!html) return [new TextRun('')];
  const runs = [];

  // Substitui variáveis <var>...</var> por [[ ]]
  html = html.replace(/<var>(.*?)<\/var>/gi, (_, v) => `[[${v}]]`);

  // Tokeniza em spans de formatação
  const tokenRe = /<(b|strong|em|i|u|br)[^>]*>([\s\S]*?)<\/\1>|<br\s*\/?>|\[\[([^\]]+)\]\]|([^<\[]+)/gi;
  let t;

  function makeRun(text, bold, italic, underline, isVar) {
    if (!text) return null;
    text = decodeEntities(text);
    const baseFont = isCitacao ? (bold ? F_CIT_B : F_CIT) : F_NORMAL;
    return new TextRun({
      text,
      font: isVar ? 'Courier New' : baseFont,
      size: isCitacao ? SZ_CIT : (isVar ? SZ_SMALL : SZ_NORMAL),
      bold: bold || false,
      italics: italic || false,
      underline: underline ? { type: 'single' } : undefined,
      color: isVar ? '92400e' : undefined,
    });
  }

  // Simple state parser
  function parse(str, bold=false, italic=false, underline=false) {
    const re = /(<(b|strong)>([\s\S]*?)<\/\2>)|(<(em|i)>([\s\S]*?)<\/\5>)|(<u>([\s\S]*?)<\/u>)|(<br\s*\/?>)|\[\[([^\]]+)\]\]|([^<\[]+)/gi;
    let m2;
    while ((m2 = re.exec(str)) !== null) {
      if (m2[1]) { parse(m2[3], true, italic, underline); }
      else if (m2[4]) { parse(m2[6], bold, true, underline); }
      else if (m2[7]) { parse(m2[8], bold, italic, true); }
      else if (m2[9]) { runs.push(new TextRun({ break: 1 })); }
      else if (m2[10]) { const r = makeRun(m2[10], bold, italic, underline, true); if(r) runs.push(r); }
      else if (m2[11]) {
        const text = m2[11].replace(/<[^>]+>/g,''); // strip any residual tags
        const r = makeRun(text, bold, italic, underline, false);
        if(r) runs.push(r);
      }
    }
  }

  parse(html);
  return runs.length ? runs : [new TextRun({ text: stripTags(decodeEntities(html)), font: isCitacao ? F_CIT : F_NORMAL, size: isCitacao ? SZ_CIT : SZ_NORMAL })];
}

function htmlTableToDocx(tableHtml) {
  try {
    const rowsHtml = [...tableHtml.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)].map(m => m[1]);
    if (!rowsHtml.length) return null;

    const rows = rowsHtml.map((rowHtml, ri) => {
      const cells = [...rowHtml.matchAll(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi)].map(m => m[1]);
      const isHeader = rowHtml.includes('<th');
      return new TableRow({
        tableHeader: isHeader,
        children: cells.map(cellHtml => new TableCell({
          borders: cellBorders,
          margins: cellMargins,
          shading: isHeader ? { fill: '1a3a6b', type: ShadingType.CLEAR } : undefined,
          children: [new Paragraph({
            children: [new TextRun({
              text: decodeEntities(stripTags(cellHtml)),
              font: F_NORMAL,
              size: SZ_SMALL,
              bold: isHeader,
              color: isHeader ? 'FFFFFF' : undefined,
            })],
            spacing: { before: 40, after: 40 },
          })],
        })),
      });
    });

    const colCount = rows[0]?.root?.[0]?.root?.length || 3;
    const colW = Math.floor(CONTENT_W / Math.max(colCount, 1));
    return new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: Array(colCount).fill(colW),
      rows,
    });
  } catch(e) {
    return null;
  }
}

// ════════════════════════════════════════════════════════════════
// BUILDERS DE PARÁGRAFOS PADRÃO
// ════════════════════════════════════════════════════════════════
function emptyPara() {
  return new Paragraph({ children: [new TextRun('')], spacing: { before: 0, after: 0 } });
}

function tituloSecao(num, texto) {
  return new Paragraph({
    children: [new TextRun({
      text: `${num} .   ${texto.toUpperCase()}`,
      font: F_NORMAL, size: SZ_TITULO, bold: true,
    })],
    spacing: SP_TITULO,
    alignment: AlignmentType.LEFT,
  });
}

function tituloSubsecao(num, texto) {
  return new Paragraph({
    children: [new TextRun({
      text: `${num} .   ${texto.toUpperCase()}`,
      font: F_NORMAL, size: SZ_TITULO, bold: true,
    })],
    spacing: { ...SP_NORMAL, before: 140 },
    alignment: AlignmentType.LEFT,
  });
}

function paraNormal(texto, bold=false, indent=true) {
  return new Paragraph({
    children: [new TextRun({ text: decodeEntities(texto), font: F_NORMAL, size: SZ_NORMAL, bold })],
    alignment: AlignmentType.JUSTIFIED,
    indent: indent ? { firstLine: Math.round(2 * 567) } : undefined,
    spacing: SP_NORMAL,
  });
}

function paraAviso(texto) {
  return new Paragraph({
    children: [new TextRun({ text: `✏  ${texto}`, font: F_NORMAL, size: SZ_SMALL, color: '856404' })],
    shading: { fill: 'FFF3CD', type: ShadingType.CLEAR },
    spacing: { before: 80, after: 80 },
    indent: { left: 200 },
    border: {
      top: { style: BorderStyle.SINGLE, size: 4, color: 'FFC107' },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: 'FFC107' },
      left: { style: BorderStyle.SINGLE, size: 16, color: 'FFC107' },
      right: { style: BorderStyle.SINGLE, size: 4, color: 'FFC107' },
    },
  });
}

function tabelaPartes(cabecalho, linhas) {
  const colCount = cabecalho.length;
  const colW = Math.floor(CONTENT_W / colCount);
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: Array(colCount).fill(colW),
    rows: [
      new TableRow({
        tableHeader: true,
        children: cabecalho.map(h => new TableCell({
          borders: cellBorders, margins: cellMargins,
          shading: { fill: '1a3a6b', type: ShadingType.CLEAR },
          children: [new Paragraph({
            children: [new TextRun({ text: h, font: F_NORMAL, size: SZ_SMALL, bold: true, color: 'FFFFFF' })],
            alignment: AlignmentType.CENTER, spacing: { before: 40, after: 40 },
          })],
        })),
      }),
      ...linhas.map(linha => new TableRow({
        children: linha.map(cel => new TableCell({
          borders: cellBorders, margins: cellMargins,
          children: [new Paragraph({
            children: [new TextRun({ text: String(cel||'—'), font: F_NORMAL, size: SZ_SMALL })],
            spacing: { before: 40, after: 40 },
          })],
        })),
      })),
    ],
  });
}

// ════════════════════════════════════════════════════════════════
// SUBSTITUIÇÃO DE VARIÁVEIS
// ════════════════════════════════════════════════════════════════
function subs(html, vars) {
  if (!html) return '';
  let r = html;
  for (const [k, v] of Object.entries(vars)) {
    r = r.replace(new RegExp(`{{${k}}}`, 'g'), v || `{{${k}}}`);
  }
  return r;
}

// ════════════════════════════════════════════════════════════════
// MONTAGEM DO DOCUMENTO
// ════════════════════════════════════════════════════════════════
function montarDocx(dados) {
  const p   = dados.processo;
  const d   = dados.diligencia || {};
  const T   = dados.textos     || {};
  const obj = p.objeto || 'insalubridade';

  const agentes = (d.agentes || []);

  const dtPericia = p.data_pericia
    ? new Date(p.data_pericia + 'T12:00:00').toLocaleDateString('pt-BR', { day:'2-digit', month:'long', year:'numeric' })
    : '{{DATA_DILIGENCIA}}';
  const dtLaudo = new Date().toLocaleDateString('pt-BR', { day:'2-digit', month:'long', year:'numeric' });

  const VARS = {
    RECLAMANTE: p.reclamante || '',
    RECLAMADA:  p.reclamada  || '',
    VARA:       p.vara       || '',
    PROCESSO:   p.numero     || '',
    ENDERECO:   p.endereco   || '',
    CIDADE:     p.cidade     || '',
    DATA_DILIGENCIA: dtPericia,
    HORARIO:    p.horario    || '',
    DATA_LAUDO: dtLaudo,
    TOTAL_PAGINAS: '{{TOTAL_PAGINAS}}',
  };

  const children = [];

  // ── CABEÇALHO IDENTIFICAÇÃO ──────────────────────────────────
  children.push(new Paragraph({
    children: [new TextRun({ text: p.perito_nome || 'CARLOS EDUARDO SILVA LAZARINI', font: F_NORMAL, size: 32, bold: true, color: '1a3a6b' })],
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
  }));
  children.push(new Paragraph({
    children: [new TextRun({ text: 'Eng. de Produção, Segurança do Trabalho e Ergonomista', font: F_NORMAL, size: SZ_SMALL })],
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 20 },
  }));
  children.push(new Paragraph({
    children: [new TextRun({ text: p.perito_crea || 'CREA-SP 506.938.283.5', font: F_NORMAL, size: SZ_SMALL, bold: true })],
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: '1a3a6b' } },
  }));

  // Destinatário e identificação
  children.push(new Paragraph({
    children: [new TextRun({ text: `EXMO(A). SENHOR(A) DOUTOR(A) JUIZ(A) DA ${(p.vara||'').toUpperCase()}.`, font: F_NORMAL, size: SZ_NORMAL })],
    spacing: { before: 200, after: 80 },
  }));
  children.push(paraNormal(`PROCESSO Nº: ${p.numero||''}`));
  children.push(paraNormal(`RECLAMANTE: ${(p.reclamante||'').toUpperCase()}`));
  children.push(paraNormal(`RECLAMADA: ${(p.reclamada||'').toUpperCase()}`));
  children.push(emptyPara());

  // Texto de apresentação
  const perito = p.perito_nome || 'CARLOS EDUARDO SILVA LAZARINI';
  const crea   = p.perito_crea || 'CREA-SP 506.938.283.5';
  children.push(paraNormal(`${perito.toUpperCase()}, Perito do Juízo, nos autos da Reclamação Trabalhista em referência, infra-assinado, em cumprimento ao r. despacho de fls., vem, mui respeitosamente, submeter à douta apreciação de V. Exa. o resultado de seu trabalho, consubstanciado no LAUDO PERICIAL, requerendo, assim, sua juntada aos autos.`));
  children.push(emptyPara());

  // ── TÓPICO 1 ─────────────────────────────────────────────────
  children.push(tituloSecao(1, 'OBJETO DA PERÍCIA'));
  const t1html = subs(T[`topico1_${obj}`] || '', VARS);
  children.push(...parseHTMLtoParagraphs(t1html || `<p>{{TEXTO_TOPICO_1_${obj.toUpperCase()}}}</p>`));

  // ── TÓPICO 2 ─────────────────────────────────────────────────
  children.push(tituloSecao(2, 'DILIGÊNCIA E FONTES DE INFORMAÇÃO'));
  const t2html = subs(T['topico2'] || '', VARS);
  children.push(...parseHTMLtoParagraphs(t2html || '<p>{{TEXTO_TOPICO_2}}</p>'));

  // Tabela reclamante
  const pRec = d.partes_reclamante || [];
  const pEmp = d.partes_reclamada  || [];
  if (pRec.length) {
    children.push(new Paragraph({ children:[new TextRun({text:'P  E  L  A     R  E  C  L  A  M  A  N  T  E', font:F_NORMAL, size:SZ_SMALL, bold:true})], alignment:AlignmentType.CENTER, spacing:{before:120,after:60}}));
    children.push(tabelaPartes(['Nome','Qualidade'], pRec.map(pt => [pt.nome, pt.mister])));
  }
  if (pEmp.length) {
    children.push(new Paragraph({ children:[new TextRun({text:'P  E  L  A     R  E  C  L  A  M  A  D  A', font:F_NORMAL, size:SZ_SMALL, bold:true})], alignment:AlignmentType.CENTER, spacing:{before:120,after:60}}));
    children.push(tabelaPartes(['Nome','Qualidade','Data de admissão'], pEmp.map(pt => [pt.nome, pt.mister, pt.admissao||'—'])));
  }

  // 2.1 Ocorrências
  children.push(tituloSubsecao('2.1', 'OCORRÊNCIAS DURANTE A DILIGÊNCIA PERICIAL'));
  if (d.ocorrencia && d.ocorrencia_texto) {
    children.push(paraNormal(d.ocorrencia_texto));
  } else {
    const t21 = subs(T['topico2_1'] || '', VARS);
    children.push(...parseHTMLtoParagraphs(t21 || '<p>Não houve ocorrências durante a diligência pericial.</p>'));
  }

  // ── TÓPICO 3 ─────────────────────────────────────────────────
  children.push(tituloSecao(3, 'LEGISLAÇÃO APLICADA'));
  if (obj === 'ambos') {
    children.push(tituloSubsecao('3.1', 'INSALUBRIDADE'));
    children.push(...parseHTMLtoParagraphs(subs(T['topico3_ins'] || '<p>{{TEXTO_TOPICO_3_INSALUBRIDADE}}</p>', VARS)));
    children.push(tituloSubsecao('3.2', 'PERICULOSIDADE'));
    children.push(...parseHTMLtoParagraphs(subs(T['topico3_peri'] || '<p>{{TEXTO_TOPICO_3_PERICULOSIDADE}}</p>', VARS)));
  } else {
    const t3key = obj === 'periculosidade' ? 'topico3_peri' : 'topico3_ins';
    children.push(...parseHTMLtoParagraphs(subs(T[t3key] || `<p>{{TEXTO_TOPICO_3_${obj.toUpperCase()}}}</p>`, VARS)));
  }

  // ── TÓPICO 4 ─────────────────────────────────────────────────
  children.push(tituloSecao(4, 'DADOS FUNCIONAIS DO RECLAMANTE'));
  children.push(tituloSubsecao('4.1', 'PELA RECLAMANTE'));
  children.push(tabelaPartes(
    ['Função', 'Data de admissão', 'Data de demissão'],
    [[d.func_reclamante||p.funcao||'—', fmtData(d.admissao_reclamante||p.admissao), fmtData(d.demissao_reclamante||p.demissao)]]
  ));
  if (d.ativ_autor) children.push(paraNormal(d.ativ_autor));

  children.push(tituloSubsecao('4.2', 'PELA RECLAMADA'));
  children.push(tabelaPartes(
    ['Função', 'Data de admissão', 'Data de demissão'],
    [[d.func_reclamada||'—', fmtData(d.admissao_reclamada), fmtData(d.demissao_reclamada)]]
  ));
  if (d.ativ_empresa) children.push(paraNormal(d.ativ_empresa));

  // Aviso de prescrição
  if (p.autuacao) {
    const dtAut = new Date(p.autuacao + 'T12:00:00');
    const dtPresc = new Date(dtAut);
    dtPresc.setFullYear(dtPresc.getFullYear() - 5);
    children.push(paraAviso(`Prescrição quinquenal: períodos anteriores a ${dtPresc.toLocaleDateString('pt-BR',{month:'long',year:'numeric'})} podem estar prescritos. (autuação: ${fmtData(p.autuacao)})`));
  }

  // ── TÓPICO 5 ─────────────────────────────────────────────────
  children.push(tituloSecao(5, 'EQUIPAMENTO DE PROTEÇÃO INDIVIDUAL (EPI)'));
  children.push(...parseHTMLtoParagraphs(subs(T['topico5'] || '<p>{{TEXTO_TOPICO_5}}</p>', VARS)));
  if (d.epis) children.push(paraNormal(`EPIs declarados verbalmente: ${d.epis}`));
  if (d.treinamentos) children.push(paraNormal(`Treinamentos realizados: ${d.treinamentos}`));

  // ── TÓPICO 6 ─────────────────────────────────────────────────
  children.push(tituloSecao(6, 'LOCAL DE TRABALHO'));
  const t6 = subs(T['topico6'] || `<p>As atividades laborais do Reclamante eram desenvolvidas nas dependências da Reclamada, situada à {{ENDERECO}}, município de {{CIDADE}}.</p>`, VARS);
  children.push(...parseHTMLtoParagraphs(t6));

  // ── TÓPICO 7 ─────────────────────────────────────────────────
  children.push(tituloSecao(7, 'IDENTIFICAÇÃO DA PRESENÇA DE AGENTES INSALUBRES E/OU PERIGOSOS'));
  const t7key = `topico7_intro_${obj}`;
  const t7html = subs(T[t7key] || T['topico7_intro_ambos'] || '<p>{{TEXTO_TOPICO_7_INTRO}}</p>', VARS);
  children.push(...parseHTMLtoParagraphs(t7html));

  // Imagens dos agentes
  if (agentes.length) {
    for (const ag of agentes) {
      const sid = agToStorageId(ag.id);
      const imgEntry = sid ? T[`img_${sid}`] : null;
      if (imgEntry?.data) {
        try {
          const b64 = imgEntry.data.split(',')[1];
          const buf = Buffer.from(b64, 'base64');
          const ext = (imgEntry.tipo||'image/png').includes('png') ? 'png' : 'jpeg';
          children.push(new Paragraph({
            children: [new ImageRun({
              data: buf, type: ext,
              transformation: { width: 180, height: 100 },
            })],
            spacing: { before: 80, after: 40 },
          }));
          children.push(paraNormal(ag.nome || ag.id, false, false));
        } catch(e) { /* imagem inválida, pula */ }
      } else {
        children.push(paraNormal(`[Imagem: ${ag.nome || ag.id}]`, false, false));
      }
    }
  }

  const t7fech = subs(T['topico7_fechamento'] || '<p>Não serão avaliados os demais agentes potencialmente insalubres e/ou perigosos devido à inexistência de exposição/contato, à exposição/contato ser eventual ou à sua concentração/intensidade ser considerada desprezível.</p>', VARS);
  children.push(...parseHTMLtoParagraphs(t7fech));

  // ── TÓPICO 8 — METODOLOGIA ───────────────────────────────────
  children.push(tituloSecao(8, 'METODOLOGIA'));
  agentes.forEach((ag, i) => {
    const sid = agToStorageId(ag.id);
    children.push(tituloSubsecao(`8.${i+1}`, `AGENTE ${(ag.nome||ag.id).toUpperCase()}`));
    const metHtml = subs(sid && T[`${sid}_met`] ? T[`${sid}_met`] : `<p>{{METODOLOGIA_${ag.id.toUpperCase()}}}</p>`, VARS);
    children.push(...parseHTMLtoParagraphs(metHtml));
  });

  // ── TÓPICO 9 — RESULTADO ─────────────────────────────────────
  children.push(tituloSecao(9, 'RESULTADO DAS AVALIAÇÕES'));
  agentes.forEach((ag, i) => {
    const sid = agToStorageId(ag.id);
    children.push(tituloSubsecao(`9.${i+1}`, `AGENTE ${(ag.nome||ag.id).toUpperCase()}`));
    const resHtml = subs(sid && T[`${sid}_res`] ? T[`${sid}_res`] : `<p>{{RESULTADO_${ag.id.toUpperCase()}}}</p>`, VARS);
    children.push(...parseHTMLtoParagraphs(resHtml));
    children.push(paraAviso('Insira aqui os dados, tabelas e conclusões da avaliação.'));
  });

  // ── TÓPICO 10 — CONCLUSÃO ─────────────────────────────────────
  children.push(tituloSecao(10, 'CONCLUSÃO PERICIAL'));
  children.push(paraAviso('Elaborar a conclusão pericial com base nos resultados das avaliações acima.'));

  // ── TÓPICO 11 — HONORÁRIOS ────────────────────────────────────
  children.push(tituloSecao(11, 'HONORÁRIOS PERICIAIS'));
  children.push(...parseHTMLtoParagraphs(subs(T['topico11_honorarios'] || '<p>{{TEXTO_HONORARIOS}}</p>', VARS)));

  // ── ENCERRAMENTO ─────────────────────────────────────────────
  children.push(tituloSecao(12, 'ENCERRAMENTO'));
  const tEnc = (T['encerramento'] || '<p>Em nada mais havendo, é dado por encerrado o presente laudo pericial.</p><p>{{CIDADE}}, {{DATA_LAUDO}}.</p>')
    .replace(/{{DATA_LAUDO}}/g, dtLaudo);
  children.push(...parseHTMLtoParagraphs(subs(tEnc, VARS)));

  // Assinatura
  children.push(emptyPara());
  children.push(new Paragraph({
    children: [new TextRun({ text: perito, font: F_NORMAL, size: SZ_NORMAL, bold: true })],
    alignment: AlignmentType.CENTER, spacing: { before: 400, after: 20 },
  }));
  children.push(new Paragraph({
    children: [new TextRun({ text: crea, font: F_NORMAL, size: SZ_SMALL })],
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 20 },
  }));
  children.push(new Paragraph({
    children: [new TextRun({ text: p.perito_email || 'eng.celazarini@gmail.com', font: F_NORMAL, size: SZ_SMALL })],
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
  }));

  // ── CABEÇALHO E RODAPÉ ───────────────────────────────────────
  const header = new Header({
    children: [new Paragraph({
      children: [
        new TextRun({ text: perito, font: F_NORMAL, size: SZ_SMALL, bold: true, color: '1a3a6b' }),
        new TextRun({ text: `  |  ${crea}`, font: F_NORMAL, size: SZ_SMALL, color: '888888' }),
      ],
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '1a3a6b' } },
      spacing: { before: 0, after: 80 },
    })],
  });

  const footer = new Footer({
    children: [new Paragraph({
      children: [
        new TextRun({ text: `Processo nº ${p.numero||''} · `, font: F_NORMAL, size: SZ_SMALL, color: '888888' }),
        new TextRun({ children: [PageNumber.CURRENT], font: F_NORMAL, size: SZ_SMALL, color: '888888' }),
        new TextRun({ text: ' / ', font: F_NORMAL, size: SZ_SMALL, color: '888888' }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: F_NORMAL, size: SZ_SMALL, color: '888888' }),
      ],
      alignment: AlignmentType.RIGHT,
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' } },
      spacing: { before: 80, after: 0 },
    })],
  });

  // ── ESTILOS ──────────────────────────────────────────────────
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: F_NORMAL, size: SZ_NORMAL } },
      },
      paragraphStyles: [
        {
          id: 'CitacaoABNT',
          name: 'Citacao ABNT',
          basedOn: 'Normal',
          run: { font: F_CIT, size: SZ_CIT },
          paragraph: {
            indent: { left: Math.round(4 * 567) },
            spacing: SP_CIT,
          },
        },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: A4_W, height: A4_H },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });

  return doc;
}

// ════════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════════
function fmtData(s) {
  if (!s) return '—';
  const d = new Date(s + 'T12:00:00');
  return isNaN(d) ? s : d.toLocaleDateString('pt-BR');
}

function stripTags(html) {
  return (html||'').replace(/<[^>]+>/g, '');
}

function decodeEntities(str) {
  return (str||'')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&nbsp;/g, ' ')
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h,16)))
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(parseInt(n,10)));
}

const AGENT_ID_MAP = {
  'nr15-a1':'ag_nr15_a1','nr15-a2':'ag_nr15_a2','nr15-a3':'ag_nr15_a3',
  'nr15-a5':'ag_nr15_a5','nr15-a6':'ag_nr15_a6','nr15-a7':'ag_nr15_a7',
  'nr15-a8':'ag_nr15_a8','nr15-a9':'ag_nr15_a9','nr15-a10':'ag_nr15_a10',
  'nr15-a11':'ag_nr15_a11','nr15-a12':'ag_nr15_a12','nr15-a13':'ag_nr15_a13',
  'nr15-a13a':'ag_nr15_a13a','nr15-a14':'ag_nr15_a14',
  'nr16-a1':'ag_nr16_a1','nr16-a2':'ag_nr16_a2','nr16-a3a':'ag_nr16_a3a',
  'nr16-a3b':'ag_nr16_a3b','nr16-a3c':'ag_nr16_a3c','nr16-a5':'ag_nr16_a5',
};
function agToStorageId(id) { return AGENT_ID_MAP[id] || null; }

// ════════════════════════════════════════════════════════════════
// PONTO DE ENTRADA
// ════════════════════════════════════════════════════════════════
async function main() {
  const args = process.argv.slice(2);
  if (args.length < 2) {
    console.error('Uso: node gerar_laudo.js <dados.json> <saida.docx>');
    process.exit(1);
  }
  const [inputFile, outputFile] = args;
  const dados = JSON.parse(fs.readFileSync(inputFile, 'utf8'));
  const doc   = montarDocx(dados);
  const buf   = await Packer.toBuffer(doc);
  fs.writeFileSync(outputFile, buf);
  console.log(`✓ DOCX gerado: ${outputFile} (${Math.round(buf.length/1024)} KB)`);
}

main().catch(e => { console.error('Erro:', e.message); process.exit(1); });
