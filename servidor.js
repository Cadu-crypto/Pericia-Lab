#!/usr/bin/env node
// PeríciaLab — Servidor Local de Integração
// Roda na sua máquina: node servidor.js
// Acesse: http://localhost:3000

'use strict';

const http = require('http');
const fs   = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const PORT    = 3000;
const DIR     = __dirname;
const TMP_DIR = path.join(DIR, 'tmp');

if (!fs.existsSync(TMP_DIR)) fs.mkdirSync(TMP_DIR);

// ── MIME TYPES ────────────────────────────────────────────────
const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js':   'application/javascript; charset=utf-8',
  '.css':  'text/css; charset=utf-8',
  '.json': 'application/json',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.png':  'image/png',
  '.jpg':  'image/jpeg',
  '.ico':  'image/x-icon',
};

// ── CORS HEADERS ─────────────────────────────────────────────
function setCORS(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// ── BODY PARSER ──────────────────────────────────────────────
function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', c => chunks.push(c));
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

// ── RESPONSE HELPERS ─────────────────────────────────────────
function json(res, status, data) {
  const body = JSON.stringify(data);
  res.writeHead(status, { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) });
  res.end(body);
}

function sendFile(res, filePath) {
  const ext  = path.extname(filePath).toLowerCase();
  const mime = MIME[ext] || 'application/octet-stream';
  const buf  = fs.readFileSync(filePath);
  res.writeHead(200, { 'Content-Type': mime, 'Content-Length': buf.length });
  res.end(buf);
}

// ── ROTAS ─────────────────────────────────────────────────────
async function handleRequest(req, res) {
  setCORS(res);

  if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return; }

  const url = req.url.split('?')[0];

  // ── API: Gerar DOCX ───────────────────────────────────────
  if (req.method === 'POST' && url === '/api/gerar-docx') {
    try {
      const body  = await readBody(req);
      const dados = JSON.parse(body);

      // Valida campos mínimos
      if (!dados.processo?.numero) {
        return json(res, 400, { erro: 'Dados incompletos: processo.numero obrigatório' });
      }

      // Salva JSON temporário
      const ts       = Date.now();
      const jsonPath = path.join(TMP_DIR, `laudo_${ts}.json`);
      const docxPath = path.join(TMP_DIR, `laudo_${ts}.docx`);
      fs.writeFileSync(jsonPath, JSON.stringify(dados, null, 2), 'utf8');

      // Chama o gerador
      const gerador = path.join(DIR, 'gerar_laudo.js');
      if (!fs.existsSync(gerador)) {
        return json(res, 500, { erro: 'gerar_laudo.js não encontrado na pasta do servidor.' });
      }

      execSync(`node "${gerador}" "${jsonPath}" "${docxPath}"`, {
        cwd: DIR, timeout: 30000,
        stdio: ['ignore', 'pipe', 'pipe'],
      });

      if (!fs.existsSync(docxPath)) {
        return json(res, 500, { erro: 'Falha na geração do DOCX.' });
      }

      // Nome do arquivo para download
      const reclamante = (dados.processo.reclamante || 'processo')
        .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
        .replace(/[^a-zA-Z0-9\s]/g,'').replace(/\s+/g,'_').substring(0, 40);
      const filename = `Laudo_${reclamante}.docx`;

      // Envia o arquivo
      const buf = fs.readFileSync(docxPath);
      res.writeHead(200, {
        'Content-Type': MIME['.docx'],
        'Content-Disposition': `attachment; filename="${filename}"`,
        'Content-Length': buf.length,
      });
      res.end(buf);

      // Limpa temporários após 60s
      setTimeout(() => {
        try { fs.unlinkSync(jsonPath); } catch {}
        try { fs.unlinkSync(docxPath); } catch {}
      }, 60000);

    } catch (e) {
      console.error('[DOCX]', e.message);
      json(res, 500, { erro: e.message });
    }
    return;
  }

  // ── API: Status ────────────────────────────────────────────
  if (req.method === 'GET' && url === '/api/status') {
    return json(res, 200, {
      ok: true,
      versao: '1.0.0',
      servidor: 'PeríciaLab Local',
      timestamp: new Date().toISOString(),
    });
  }

  // ── Servir arquivos estáticos ──────────────────────────────
  let filePath = url === '/' ? '/index.html' : url;
  filePath = path.join(DIR, filePath.replace(/\.\./g, '')); // Evita path traversal

  if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
    return sendFile(res, filePath);
  }

  // 404
  res.writeHead(404, { 'Content-Type': 'text/plain' });
  res.end('Não encontrado: ' + url);
}

// ── INICIA SERVIDOR ───────────────────────────────────────────
const server = http.createServer(handleRequest);

server.listen(PORT, '127.0.0.1', () => {
  console.log('');
  console.log('  ╔════════════════════════════════════════╗');
  console.log('  ║     PeríciaLab — Servidor Local        ║');
  console.log('  ╠════════════════════════════════════════╣');
  console.log(`  ║  Endereço:  http://localhost:${PORT}       ║`);
  console.log('  ║  Para parar: Ctrl + C                  ║');
  console.log('  ╚════════════════════════════════════════╝');
  console.log('');
  console.log('  Aguardando conexões...');
});

server.on('error', e => {
  if (e.code === 'EADDRINUSE') {
    console.error(`\n  Porta ${PORT} em uso. Feche o outro processo e tente novamente.\n`);
  } else {
    console.error('\n  Erro:', e.message);
  }
  process.exit(1);
});

process.on('SIGINT', () => {
  console.log('\n\n  Servidor encerrado.\n');
  process.exit(0);
});
