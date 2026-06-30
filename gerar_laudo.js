#!/usr/bin/env node
// PeríciaLab — Gerador de Laudo DOCX (Windows-compatible)
// Usa template.docx como base via adm-zip (sem unzip/zip do sistema)
// Uso: node gerar_laudo.js dados.json saida.docx
'use strict';

const fs      = require('fs');
const path    = require('path');
const AdmZip  = require('adm-zip');

const SCRIPT_DIR = __dirname;
const TEMPLATE   = path.join(SCRIPT_DIR, 'template.docx');

// ── CONSTANTES ────────────────────────────────────────────────
const INDENT_FIRST = 709;
const INDENT_CIT   = 2268;
const SP_BEFORE    = 60;
const SP_AFTER     = 60;
const LINE_15      = 360;
const LINE_115     = 276;
const COR_VINHO    = '6B0F1A';
const CW           = 9071; // largura do conteúdo em DXA

// ════════════════════════════════════════════════════════════════
// HELPERS XML
// ════════════════════════════════════════════════════════════════
function esc(s) {
  return String(s || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&apos;');
}
function decode(s) {
  return (s||'').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>')
    .replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/&nbsp;/g,' ')
    .replace(/&#x([0-9a-fA-F]+);/g,(_,h)=>String.fromCharCode(parseInt(h,16)))
    .replace(/&#(\d+);/g,(_,n)=>String.fromCharCode(parseInt(n,10)));
}
function strip(h) { return (h||'').replace(/<[^>]+>/g,' ').replace(/\s+/g,' ').trim(); }
function fmtData(s) {
  if (!s) return '—';
  const d = new Date(s+'T12:00:00');
  return isNaN(d) ? s : d.toLocaleDateString('pt-BR');
}

// ── Run XML ───────────────────────────────────────────────────
function run(txt, o={}) {
  const b = o.bold      ? '<w:b/><w:bCs/>'        : '';
  const i = o.italic    ? '<w:i/><w:iCs/>'         : '';
  const u = o.underline ? '<w:u w:val="single"/>'  : '';
  const c = o.color     ? `<w:color w:val="${o.color}"/>` : '';
  const sz = o.sz || 24;
  const f  = o.font || 'Arial';
  return `<w:r><w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>${b}${i}${u}${c}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr><w:t xml:space="preserve">${esc(txt)}</w:t></w:r>`;
}

// ── HTML → runs XML ───────────────────────────────────────────
function htmlRuns(html, isCit=false) {
  const sz = isCit ? 22 : 24;
  const parts = [];
  function parse(str, bold=false, italic=false, ul=false) {
    const re = /(<(strong|b)>([\s\S]*?)<\/\2>)|(<(em|i)>([\s\S]*?)<\/\5>)|(<u>([\s\S]*?)<\/u>)|(<br\s*\/?>)|([^<]+)/gi;
    let m;
    while ((m=re.exec(str))!==null) {
      if (m[1])      parse(m[3],true,italic,ul);
      else if (m[4]) parse(m[6],bold,true,ul);
      else if (m[7]) parse(m[8],bold,italic,true);
      else if (m[9]) parts.push('<w:r><w:br/></w:r>');
      else if (m[10]) {
        const t = decode(m[10]);
        if (!t.trim()) continue;
        const b2=bold?'<w:b/><w:bCs/>':'', i2=italic?'<w:i/><w:iCs/>':'', u2=ul?'<w:u w:val="single"/>':'';
        parts.push(`<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>${b2}${i2}${u2}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr><w:t xml:space="preserve">${esc(t)}</w:t></w:r>`);
      }
    }
  }
  parse(html);
  return parts.length ? parts.join('') : run(strip(html),{sz});
}

// ── Parágrafo padrão ──────────────────────────────────────────
function paraHtml(html, o={}) {
  const isCit = o.citacao||false;
  const ind = isCit
    ? `<w:ind w:start="${INDENT_CIT}"/>`
    : (o.noIndent?'': `<w:ind w:firstLine="${INDENT_FIRST}"/>`);
  const line = isCit ? LINE_115 : LINE_15;
  const jc   = o.center ? 'center' : 'both';
  return `<w:p><w:pPr><w:pStyle w:val="Default"/><w:spacing w:lineRule="auto" w:line="${line}" w:before="${SP_BEFORE}" w:after="${SP_AFTER}"/><w:jc w:val="${jc}"/>${ind}</w:pPr>${htmlRuns(html,isCit)}</w:p>`;
}
function emptyPara() { return `<w:p><w:pPr><w:pStyle w:val="Default"/><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>`; }
function pageBreak() { return `<w:p><w:pPr><w:pStyle w:val="Default"/><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>`; }

// ── HTML em bloco (com blockquote) ────────────────────────────
function htmlBloco(html) {
  if (!html) return emptyPara();
  const r=[];
  const parts=[];
  let last=0;
  const bqRe=/<blockquote[^>]*>([\s\S]*?)<\/blockquote>/gi;
  let mx;
  while ((mx=bqRe.exec(html))!==null) {
    if (mx.index>last) parts.push({t:'text',c:html.slice(last,mx.index)});
    parts.push({t:'bq',c:mx[1]});
    last=mx.index+mx[0].length;
  }
  if (last<html.length) parts.push({t:'text',c:html.slice(last)});
  if (!parts.length) parts.push({t:'text',c:html});

  for (const p of parts) {
    if (p.t==='bq') {
      const pRe=/<p[^>]*>([\s\S]*?)<\/p>/gi; let pm,f=false;
      while ((pm=pRe.exec(p.c))!==null) { r.push(paraHtml(pm[1],{citacao:true})); f=true; }
      if (!f) r.push(paraHtml(p.c,{citacao:true}));
    } else {
      const pRe=/<p[^>]*>([\s\S]*?)<\/p>/gi; let pm,f=false;
      while ((pm=pRe.exec(p.c))!==null) { r.push(paraHtml(pm[1])); f=true; }
      if (!f && p.c.trim()) r.push(paraHtml(p.c));
    }
  }
  return r.length ? r.join('') : emptyPara();
}

// ── Título de seção ───────────────────────────────────────────
function titulo(num,txt) {
  return `<w:p><w:pPr><w:pStyle w:val="Default"/><w:spacing w:lineRule="auto" w:line="${LINE_15}" w:before="200" w:after="120"/><w:jc w:val="start"/></w:pPr>${run(`${num} .   ${txt.toUpperCase()}`,{bold:true})}</w:p>`;
}
function subTitulo(num,txt) {
  return `<w:p><w:pPr><w:pStyle w:val="Default"/><w:spacing w:lineRule="auto" w:line="${LINE_15}" w:before="160" w:after="100"/><w:jc w:val="start"/></w:pPr>${run(`${num} .   ${txt.toUpperCase()}`,{bold:true})}</w:p>`;
}

// ── Rótulo de partes (P E L A   R E C L A M A N T E) ─────────
function labelPartes(txt) {
  return `<w:p><w:pPr><w:pStyle w:val="Default"/><w:spacing w:lineRule="auto" w:line="${LINE_15}" w:before="160" w:after="80"/><w:jc w:val="center"/></w:pPr>${run(txt.split('').join('  ').toUpperCase(),{bold:true,sz:20})}</w:p>`;
}

// ── Tabela com header vinho ────────────────────────────────────
function tabelaVinho(cabecalho, linhas, colWidths) {
  const total = colWidths.reduce((a,b)=>a+b,0);
  const bv = (c) => `<w:top w:val="single" w:sz="4" w:color="${c}"/><w:bottom w:val="single" w:sz="4" w:color="${c}"/><w:left w:val="single" w:sz="4" w:color="${c}"/><w:right w:val="single" w:sz="4" w:color="${c}"/>`;
  const tcMar = `<w:tcMar><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>`;

  const headerCells = cabecalho.map((h,i)=>`<w:tc><w:tcPr><w:tcW w:w="${colWidths[i]}" w:type="dxa"/><w:shd w:val="clear" w:fill="${COR_VINHO}"/><w:tcBorders>${bv(COR_VINHO)}</w:tcBorders>${tcMar}</w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>${run(h,{bold:true,sz:16,color:'FFFFFF'})}</w:p></w:tc>`).join('');

  const dataRows = linhas.map(row=>{
    const cells = row.map((cel,i)=>`<w:tc><w:tcPr><w:tcW w:w="${colWidths[i]}" w:type="dxa"/><w:tcBorders>${bv('CCCCCC')}</w:tcBorders>${tcMar}</w:tcPr><w:p><w:pPr><w:jc w:val="start"/><w:spacing w:before="0" w:after="0"/></w:pPr>${run(String(cel||'—'),{sz:18})}</w:p></w:tc>`).join('');
    return `<w:tr>${cells}</w:tr>`;
  }).join('');

  return `<w:tbl><w:tblPr><w:tblW w:w="${total}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>${bv('CCCCCC')}</w:tblBorders></w:tblPr><w:tblGrid>${colWidths.map(w=>`<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr>${headerCells}</w:tr>${dataRows}</w:tbl>`;
}

// ── Tabela identificação (borda preta) ────────────────────────
function tabelaId(linhas, colWidths) {
  const total = colWidths.reduce((a,b)=>a+b,0);
  const bp = `<w:top w:val="single" w:sz="6" w:color="000000"/><w:bottom w:val="single" w:sz="6" w:color="000000"/><w:left w:val="single" w:sz="6" w:color="000000"/><w:right w:val="single" w:sz="6" w:color="000000"/>`;
  const tcMar = `<w:tcMar><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>`;
  const rows = linhas.map(([label,valor])=>`<w:tr>
    <w:tc><w:tcPr><w:tcW w:w="${colWidths[0]}" w:type="dxa"/><w:tcBorders>${bp}</w:tcBorders>${tcMar}</w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>${run(label,{bold:true})}</w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="${colWidths[1]}" w:type="dxa"/><w:tcBorders>${bp}</w:tcBorders>${tcMar}</w:tcPr><w:p><w:pPr><w:jc w:val="end"/><w:spacing w:before="0" w:after="0"/></w:pPr>${run(valor,{bold:true})}</w:p></w:tc>
  </w:tr>`).join('');
  return `<w:tbl><w:tblPr><w:tblW w:w="${total}" w:type="dxa"/></w:tblPr><w:tblGrid>${colWidths.map(w=>`<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`;
}

function paraAviso(txt) {
  return `<w:p><w:pPr><w:pStyle w:val="Default"/><w:spacing w:before="100" w:after="100"/><w:ind w:start="200"/><w:pBdr><w:left w:val="single" w:sz="24" w:color="FFC107"/></w:pBdr><w:shd w:val="clear" w:fill="FFF3CD"/></w:pPr>${run('⚠  '+txt,{sz:20,color:'7B4F00'})}</w:p>`;
}

function subs(html, vars) {
  if (!html) return '';
  let r=html;
  for (const [k,v] of Object.entries(vars)) r=r.replace(new RegExp(`{{${k}}}`, 'g'), v||'');
  return r;
}

// Converte objeto completo (insalubridade/periculosidade/ambos/ergonomia/
// insalubridade_periculosidade) para a chave curta usada no cadastro de
// textos padrão (ins/peri/ambos/erg) — deve bater com web-textos-padrao.html
function objVariante(obj) {
  const MAP = {
    'insalubridade': 'ins',
    'periculosidade': 'peri',
    'insalubridade_periculosidade': 'ambos',
    'ambos': 'ambos',
    'ergonomia': 'erg',
  };
  return MAP[obj] || 'ins';
}

// ════════════════════════════════════════════════════════════════
// MONTA O BODY XML
// ════════════════════════════════════════════════════════════════
function buildBody(dados) {
  const p  = dados.processo;
  const d  = dados.diligencia || {};
  const T  = dados.textos     || {};
  const obj = p.objeto || 'insalubridade';
  const objVar = objVariante(obj);
  const agentes = d.agentes || [];

  const AMAP = {
    'nr15-a1':'ag_nr15_a1','nr15-a2':'ag_nr15_a2','nr15-a3':'ag_nr15_a3',
    'nr15-a5':'ag_nr15_a5','nr15-a6':'ag_nr15_a6','nr15-a7':'ag_nr15_a7',
    'nr15-a8':'ag_nr15_a8','nr15-a9':'ag_nr15_a9','nr15-a10':'ag_nr15_a10',
    'nr15-a11':'ag_nr15_a11','nr15-a12':'ag_nr15_a12','nr15-a13':'ag_nr15_a13',
    'nr15-a14':'ag_nr15_a14','nr16-a1':'ag_nr16_a1','nr16-a2':'ag_nr16_a2',
    'nr16-a3a':'ag_nr16_a3a','nr16-a3b':'ag_nr16_a3b','nr16-a3c':'ag_nr16_a3c',
    'nr16-a5':'ag_nr16_a5',
  };

  const dtPericia = p.data_pericia
    ? new Date(p.data_pericia+'T12:00:00').toLocaleDateString('pt-BR',{day:'2-digit',month:'long',year:'numeric'})
    : '{{DATA_DILIGENCIA}}';
  const dtLaudo = new Date().toLocaleDateString('pt-BR',{day:'2-digit',month:'long',year:'numeric'});
  const VARS = { RECLAMANTE:p.reclamante||'', RECLAMADA:p.reclamada||'', VARA:p.vara||'',
    PROCESSO:p.numero||'', ENDERECO:p.endereco||'', CIDADE:p.cidade||'',
    DATA_DILIGENCIA:dtPericia, HORARIO:p.horario||'', DATA_LAUDO:dtLaudo };

  const CW1=Math.round(CW*0.35), CW2=CW-CW1;
  const perito = p.perito_nome||'CARLOS EDUARDO SILVA LAZARINI';
  const r=[];

  // Destinatário + tabela id
  r.push(paraHtml(`<strong>EXMO(A). SENHOR(A) DOUTOR(A) JUIZ(A) DA ${(p.vara||'').toUpperCase()}.</strong>`,{noIndent:true}));
  r.push(emptyPara());
  r.push(tabelaId([['PROCESSO Nº:',p.numero||''],['RECLAMANTE:',(p.reclamante||'').toUpperCase()],['RECLAMADA:',(p.reclamada||'').toUpperCase()]],[CW1,CW2]));
  r.push(emptyPara());
  r.push(paraHtml(`<strong>${perito.toUpperCase()}</strong>, Perito do Juízo, nos autos da Reclamação Trabalhista em referência, infra-assinado, em cumprimento ao r. despacho de fls., vem, mui respeitosamente, submeter à douta apreciação de V. Exa. o resultado de seu trabalho, consubstanciado no <strong>LAUDO PERICIAL</strong>, requerendo, assim, sua juntada aos autos.`));
  r.push(emptyPara());

  // Tópico 1
  r.push(pageBreak()); r.push(titulo(1,'OBJETO DA PERÍCIA'));
  r.push(htmlBloco(subs(T[`topico1_${objVar}`]||'',VARS)));

  // Tópico 2
  r.push(pageBreak()); r.push(titulo(2,'DILIGÊNCIA E FONTES DE INFORMAÇÃO'));
  r.push(htmlBloco(subs(T['topico2']||'',VARS)));
  if ((d.partes_reclamante||[]).length) {
    r.push(labelPartes('PELA RECLAMANTE'));
    r.push(tabelaVinho(['Nome','Mister','Documento'],(d.partes_reclamante||[]).map(pt=>[pt.nome||'',pt.mister||'',pt.documento||'—']),[Math.round(CW*.42),Math.round(CW*.27),Math.round(CW*.31)]));
    r.push(emptyPara());
  }
  if ((d.partes_reclamada||[]).length) {
    r.push(labelPartes('PELA RECLAMADA'));
    r.push(tabelaVinho(['Nome','Mister','Documento','Admissão'],(d.partes_reclamada||[]).map(pt=>[pt.nome||'',pt.mister||'',pt.documento||'—',pt.admissao||'—']),[Math.round(CW*.32),Math.round(CW*.21),Math.round(CW*.27),Math.round(CW*.20)]));
    r.push(emptyPara());
  }
  r.push(subTitulo('2.1','OCORRÊNCIAS DURANTE A DILIGÊNCIA PERICIAL'));
  r.push(htmlBloco(subs(d.ocorrencia&&d.ocorrencia_texto?`<p>${d.ocorrencia_texto}</p>`:T['topico2_1']||'<p>Não houve ocorrências durante a diligência pericial.</p>',VARS)));

  // Tópico 3
  r.push(pageBreak()); r.push(titulo(3,'LEGISLAÇÃO APLICADA'));
  if (obj==='ambos' || obj==='insalubridade_periculosidade') {
    r.push(subTitulo('3.1','INSALUBRIDADE')); r.push(htmlBloco(subs(T['topico3_ins']||'',VARS)));
    r.push(subTitulo('3.2','PERICULOSIDADE')); r.push(htmlBloco(subs(T['topico3_peri']||'',VARS)));
  } else {
    r.push(htmlBloco(subs(T[`topico3_${objVar}`]||'',VARS)));
  }

  // Tópico 4
  r.push(pageBreak()); r.push(titulo(4,'DADOS FUNCIONAIS DO RECLAMANTE'));
  r.push(subTitulo('4.1','PELA RECLAMANTE'));
  const funcR=d.func_reclamante||p.funcao||'—', admR=fmtData(d.admissao_reclamante||p.admissao), demR=fmtData(d.demissao_reclamante||p.demissao);
  r.push(paraHtml(`Conforme a petição inicial, o Reclamante exercia a função de <strong>${funcR}</strong>, admitido em ${admR}${demR!=='—'?`, com demissão em ${demR}`:''}.`));
  if (d.ativ_autor) r.push(paraHtml(d.ativ_autor));
  r.push(subTitulo('4.2','PELA RECLAMADA'));
  r.push(paraHtml(d.ativ_empresa||'Não há divergências em relação às atividades descritas pelo autor.'));
  if (p.autuacao) {
    const dtP=new Date(p.autuacao+'T12:00:00'); dtP.setFullYear(dtP.getFullYear()-5);
    r.push(paraAviso(`Prescrição quinquenal: períodos anteriores a ${dtP.toLocaleDateString('pt-BR',{month:'long',year:'numeric'})} podem estar prescritos. (autuação: ${fmtData(p.autuacao)})`));
  }

  // Tópico 5
  r.push(pageBreak()); r.push(titulo(5,'EQUIPAMENTO DE PROTEÇÃO INDIVIDUAL (EPI)'));
  r.push(htmlBloco(subs(T['topico5']||'',VARS)));
  if (d.epis)         r.push(paraHtml(`<strong>EPIs declarados verbalmente:</strong> ${d.epis}`));
  if (d.treinamentos) r.push(paraHtml(`<strong>Treinamentos realizados:</strong> ${d.treinamentos}`));

  // Tópico 6
  r.push(pageBreak()); r.push(titulo(6,'LOCAL DE TRABALHO'));
  r.push(htmlBloco(subs(T['topico6']||`<p>A Reclamante desempenhou suas atividades na sede da Reclamada, situada à {{ENDERECO}}, no município de {{CIDADE}}, Estado do Paraná.</p>`,VARS)));

  // Tópico 7
  r.push(pageBreak()); r.push(titulo(7,'IDENTIFICAÇÃO DA PRESENÇA DE AGENTES INSALUBRES E/OU PERIGOSOS'));
  r.push(htmlBloco(subs(T[`topico7_intro_${objVar}`]||T['topico7_intro_ambos']||'',VARS)));
  r.push(htmlBloco(subs(T['topico7_fechamento']||'<p>Não serão avaliados os demais agentes potencialmente insalubres e/ou perigosos devido à inexistência de exposição/contato.</p>',VARS)));

  // Tópico 8
  r.push(pageBreak()); r.push(titulo(8,'METODOLOGIA'));
  agentes.forEach((ag,i)=>{
    const sid=AMAP[ag.id];
    r.push(subTitulo(`8.${i+1}`,`AGENTE ${(ag.nome||ag.id).toUpperCase()}`));
    r.push(htmlBloco(subs(sid&&T[`${sid}_met`]?T[`${sid}_met`]:`<p>Metodologia do agente ${ag.nome||ag.id}.</p>`,VARS)));
  });

  // Tópico 9
  r.push(pageBreak()); r.push(titulo(9,'RESULTADO DAS AVALIAÇÕES'));
  agentes.forEach((ag,i)=>{
    const sid=AMAP[ag.id];
    r.push(subTitulo(`9.${i+1}`,`AGENTE ${(ag.nome||ag.id).toUpperCase()}`));
    r.push(htmlBloco(subs(sid&&T[`${sid}_res`]?T[`${sid}_res`]:`<p>Resultado da avaliação do agente ${ag.nome||ag.id}.</p>`,VARS)));
  });

  // Tópico 10
  r.push(pageBreak()); r.push(titulo(10,'CONCLUSÃO PERICIAL'));
  r.push(paraHtml('Com base nas informações expostas ao longo do presente laudo pericial, nos resultados das avaliações realizadas e nos riscos potenciais à saúde analisados sob a ótica da Higiene e Segurança do Trabalho, conclui-se que:'));
  r.push(emptyPara());

  // Tópico 11
  r.push(pageBreak()); r.push(titulo(11,'HONORÁRIOS PERICIAIS'));
  r.push(htmlBloco(subs(T['topico11_honorarios']||'',VARS)));

  // Tópico 12
  r.push(pageBreak()); r.push(titulo(12,'ENCERRAMENTO'));
  r.push(htmlBloco(subs(T['encerramento']||`<p>Em nada mais havendo, é dado por encerrado o presente laudo pericial.</p><p>{{CIDADE}} / PR, {{DATA_LAUDO}}.</p>`,VARS)));
  r.push(emptyPara()); r.push(emptyPara());
  r.push(paraHtml(`<strong>${perito}</strong>`,{noIndent:true,center:true}));
  r.push(paraHtml(p.perito_crea||'CREA-SP 506.938.283.5',{noIndent:true,center:true}));
  r.push(paraHtml(p.perito_email||'eng.celazarini@gmail.com',{noIndent:true,center:true}));

  return r.join('\n');
}

// ════════════════════════════════════════════════════════════════
// GERA O DOCX — usa adm-zip puro (funciona no Windows sem unzip)
// ════════════════════════════════════════════════════════════════
async function gerarDocx(dados, outputPath) {
  if (!fs.existsSync(TEMPLATE)) {
    throw new Error(`Template não encontrado: ${TEMPLATE}\nColoque o arquivo template.docx na mesma pasta que gerar_laudo.js`);
  }

  // Abre o template com adm-zip
  const zip = new AdmZip(TEMPLATE);

  // Lê o document.xml do template
  const docEntry = zip.getEntry('word/document.xml');
  if (!docEntry) throw new Error('Arquivo word/document.xml não encontrado no template');
  let docXml = docEntry.getData().toString('utf8');

  // Monta o novo body mantendo sectPr do template (configs de página, header, footer)
  const bodyContent = buildBody(dados);
  const sectPrMatch = docXml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
  const sectPr = sectPrMatch ? sectPrMatch[0] : '';
  const newBody = `<w:body>\n${bodyContent}\n${sectPr}\n</w:body>`;

  // Substitui o body no XML
  docXml = docXml.replace(/<w:body>[\s\S]*<\/w:body>/, newBody);

  // Atualiza o ZIP em memória
  zip.updateFile('word/document.xml', Buffer.from(docXml, 'utf8'));

  // Salva o arquivo de saída
  // Usa toBuffer() para garantir compatibilidade máxima
  const buf = zip.toBuffer();
  fs.writeFileSync(outputPath, buf);
}

// ════════════════════════════════════════════════════════════════
// PONTO DE ENTRADA
// ════════════════════════════════════════════════════════════════
async function main() {
  const [inputFile, outputFile] = process.argv.slice(2);
  if (!inputFile || !outputFile) {
    console.error('Uso: node gerar_laudo.js <dados.json> <saida.docx>');
    process.exit(1);
  }
  const dados = JSON.parse(fs.readFileSync(inputFile, 'utf8'));
  await gerarDocx(dados, outputFile);
  const sz = Math.round(fs.statSync(outputFile).size / 1024);
  console.log(`✓ DOCX gerado: ${outputFile} (${sz} KB)`);
}

main().catch(e => { console.error('Erro:', e.message); process.exit(1); });
