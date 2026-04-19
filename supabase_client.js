// PeríciaLab — Cliente Supabase
// Inclua este script ANTES dos outros scripts em cada página:
// <script src="supabase_client.js"></script>

// ════════════════════════════════════════════════════════════════
// CONFIGURAÇÃO — preencha com suas credenciais do Supabase
// Painel Supabase > Settings > API
// ════════════════════════════════════════════════════════════════
const SUPABASE_CONFIG = {
  url:    'https://vzxcmxxqbqcjyherbztz.supabase.co',
  anon:   'sb_publishable_qKeZW-KS6GLo6H3Z2kZgNg_92iqy4Pn',
};

// ════════════════════════════════════════════════════════════════
// CLIENTE SUPABASE LEVE (sem dependência de npm)
// Usa a API REST do Supabase diretamente via fetch
// ════════════════════════════════════════════════════════════════
class SupabaseClient {
  constructor({ url, anon }) {
    this.url  = url.replace(/\/$/, '');
    this.anon = anon;
    this.token = null; // JWT após login
    this._loadToken();
  }

  // ── AUTH ────────────────────────────────────────────────────
  _loadToken() {
    try { this.token = localStorage.getItem('sb_token') || null; }
    catch { this.token = null; }
  }
  _saveToken(t) {
    this.token = t;
    try { if (t) localStorage.setItem('sb_token', t); else localStorage.removeItem('sb_token'); }
    catch {}
  }

  get _headers() {
    const h = {
      'Content-Type':  'application/json',
      'apikey':         this.anon,
      'Authorization': `Bearer ${this.token || this.anon}`,
    };
    return h;
  }

  async signIn(email, password) {
    const r = await fetch(`${this.url}/auth/v1/token?grant_type=password`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'apikey': this.anon },
      body: JSON.stringify({ email, password }),
    });
    const data = await r.json();
    if (!r.ok) throw new Error(data.error_description || data.message || 'Erro ao fazer login');
    this._saveToken(data.access_token);
    return data;
  }

  async signOut() {
    await fetch(`${this.url}/auth/v1/logout`, {
      method: 'POST', headers: this._headers,
    }).catch(() => {});
    this._saveToken(null);
  }

  isLoggedIn() { return !!this.token; }

  // ── REST API ────────────────────────────────────────────────
  async _req(method, path, body, params = {}) {
    const qs = new URLSearchParams(params).toString();
    const url = `${this.url}/rest/v1/${path}${qs ? '?' + qs : ''}`;
    const opts = { method, headers: { ...this._headers, 'Prefer': 'return=representation' } };
    if (body) opts.body = JSON.stringify(body);
    const r = await fetch(url, opts);
    if (r.status === 204) return [];
    const data = await r.json();
    if (!r.ok) throw new Error(data.message || data.error || `HTTP ${r.status}`);
    return data;
  }

  // SELECT
  async select(table, filters = {}, options = {}) {
    const params = { select: options.select || '*' };
    for (const [k, v] of Object.entries(filters)) {
      if (v !== undefined && v !== null) params[k] = `eq.${v}`;
    }
    if (options.order)  params.order  = options.order;
    if (options.limit)  params.limit  = options.limit;
    if (options.offset) params.offset = options.offset;
    return this._req('GET', table, null, params);
  }

  // INSERT
  async insert(table, row) {
    return this._req('POST', table, row);
  }

  // UPSERT (insert or update by primary key)
  async upsert(table, row, onConflict = 'id') {
    const r = await fetch(`${this.url}/rest/v1/${table}?on_conflict=${onConflict}`, {
      method: 'POST',
      headers: { ...this._headers, 'Prefer': 'return=representation,resolution=merge-duplicates' },
      body: JSON.stringify(row),
    });
    if (r.status === 204) return [];
    const data = await r.json();
    if (!r.ok) throw new Error(data.message || data.error || `HTTP ${r.status}`);
    return data;
  }

  // UPDATE
  async update(table, id, row) {
    return this._req('PATCH', table + `?id=eq.${id}`, row);
  }

  // DELETE
  async delete(table, id) {
    return this._req('DELETE', table + `?id=eq.${id}`);
  }

  // SELECT ONE
  async selectOne(table, filters = {}) {
    const rows = await this.select(table, filters, { limit: 1 });
    return rows[0] || null;
  }

  // SELECT por chave de texto
  async selectByKey(table, keyCol, keyVal) {
    const params = { select: '*', [keyCol]: `eq.${encodeURIComponent(keyVal)}` };
    const rows = await this._req('GET', table, null, params);
    return rows[0] || null;
  }

  // RPC (funções PostgreSQL)
  async rpc(fn, params = {}) {
    const r = await fetch(`${this.url}/rest/v1/rpc/${fn}`, {
      method: 'POST', headers: this._headers, body: JSON.stringify(params),
    });
    const data = await r.json();
    if (!r.ok) throw new Error(data.message || `RPC error: ${fn}`);
    return data;
  }
}

// Instância global
const sb = new SupabaseClient(SUPABASE_CONFIG);

// ════════════════════════════════════════════════════════════════
// SYNC MANAGER — sincroniza dados locais com o Supabase
// ════════════════════════════════════════════════════════════════
const SyncManager = {

  // ── Textos padrão: web → Supabase ──────────────────────────
  async syncTextos() {
    if (!sb.isLoggedIn()) return { ok: false, msg: 'Não autenticado' };
    try {
      const local = JSON.parse(localStorage.getItem('pericia_textos_v1') || '{}');
      let count = 0;
      for (const [chave, conteudo_html] of Object.entries(local)) {
        if (chave.startsWith('img_')) continue; // imagens tratadas separado
        await sb.upsert('textos_padrao', { chave, conteudo_html }, 'chave');
        count++;
      }
      return { ok: true, count };
    } catch(e) {
      return { ok: false, msg: e.message };
    }
  },

  // ── Textos padrão: Supabase → local ────────────────────────
  async fetchTextos() {
    if (!sb.isLoggedIn()) return { ok: false, msg: 'Não autenticado' };
    try {
      const rows = await sb.select('textos_padrao');
      const local = JSON.parse(localStorage.getItem('pericia_textos_v1') || '{}');
      for (const row of rows) {
        local[row.chave] = row.conteudo_html;
      }
      // Busca imagens também
      const imgs = await sb.select('imagens_agentes');
      for (const img of imgs) {
        local[`img_${img.agente_id}`] = {
          data: img.dados_base64,
          nome: img.nome_arquivo,
          tipo: img.mime_type,
        };
      }
      localStorage.setItem('pericia_textos_v1', JSON.stringify(local));
      return { ok: true, count: rows.length };
    } catch(e) {
      return { ok: false, msg: e.message };
    }
  },

  // ── Imagens: web → Supabase ─────────────────────────────────
  async syncImagens() {
    if (!sb.isLoggedIn()) return { ok: false, msg: 'Não autenticado' };
    try {
      const local = JSON.parse(localStorage.getItem('pericia_textos_v1') || '{}');
      let count = 0;
      for (const [chave, val] of Object.entries(local)) {
        if (!chave.startsWith('img_')) continue;
        const agente_id = chave.replace('img_', '');
        await sb.upsert('imagens_agentes', {
          agente_id,
          nome_arquivo:  val.nome  || agente_id,
          mime_type:     val.tipo  || 'image/png',
          dados_base64:  val.data  || '',
        }, 'agente_id');
        count++;
      }
      return { ok: true, count };
    } catch(e) {
      return { ok: false, msg: e.message };
    }
  },

  // ── Processos: local → Supabase ─────────────────────────────
  async syncProcessos() {
    if (!sb.isLoggedIn()) return { ok: false, msg: 'Não autenticado' };
    try {
      const local = JSON.parse(localStorage.getItem('pericia_processos_v1') || '[]');
      let count = 0;
      for (const proc of local) {
        const { diligencia, ...processo } = proc;
        await sb.upsert('processos', processo, 'id');
        if (diligencia) {
          await sb.upsert('diligencias', {
            ...diligencia,
            id: `dil_${proc.id}`,
            processo_id: proc.id,
          }, 'processo_id');
        }
        count++;
      }
      return { ok: true, count };
    } catch(e) {
      return { ok: false, msg: e.message };
    }
  },

  // ── Processos: Supabase → local ─────────────────────────────
  async fetchProcessos() {
    if (!sb.isLoggedIn()) return { ok: false, msg: 'Não autenticado' };
    try {
      const procs = await sb.select('processos', {}, { order: 'data_pericia.desc' });
      const dils  = await sb.select('diligencias');
      const dilMap = {};
      for (const d of dils) dilMap[d.processo_id] = d;

      const merged = procs.map(p => ({
        ...p,
        diligencia: dilMap[p.id] || null,
        diligencia_salva: !!dilMap[p.id],
      }));

      localStorage.setItem('pericia_processos_v1', JSON.stringify(merged));
      return { ok: true, count: procs.length };
    } catch(e) {
      return { ok: false, msg: e.message };
    }
  },

  // ── Sincronização completa ───────────────────────────────────
  async syncTudo() {
    const resultados = {};
    resultados.textos    = await this.syncTextos();
    resultados.imagens   = await this.syncImagens();
    resultados.processos = await this.syncProcessos();
    return resultados;
  },

  async fetchTudo() {
    const resultados = {};
    resultados.textos    = await this.fetchTextos();
    resultados.processos = await this.fetchProcessos();
    return resultados;
  },
};

// ════════════════════════════════════════════════════════════════
// SINCRONIZAÇÃO AUTOMÁTICA quando online
// ════════════════════════════════════════════════════════════════
window.addEventListener('online', async () => {
  if (sb.isLoggedIn()) {
    console.log('[Sync] Online — sincronizando...');
    await SyncManager.syncTudo();
    console.log('[Sync] Concluído.');
  }
});

console.log('[PeríciaLab] Cliente Supabase carregado.');
