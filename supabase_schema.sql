-- ═══════════════════════════════════════════════════════════════
-- PeríciaLab — Esquema do Banco de Dados (Supabase / PostgreSQL)
-- Execute este script no SQL Editor do Supabase
-- ═══════════════════════════════════════════════════════════════

-- ── EXTENSÕES ────────────────────────────────────────────────
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- ── TABELA: processos ─────────────────────────────────────────
-- Importados da planilha Excel via app iPad ou plataforma web
CREATE TABLE IF NOT EXISTS processos (
  id              UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  numero          TEXT NOT NULL,
  reclamante      TEXT,
  reclamada       TEXT,
  vara            TEXT,
  cidade          TEXT,
  endereco        TEXT,
  data_pericia    DATE,
  horario         TEXT,
  admissao        DATE,
  demissao        DATE,
  autuacao        DATE,
  funcao          TEXT,
  objeto          TEXT DEFAULT 'insalubridade', -- insalubridade | periculosidade | ambos
  nr15            TEXT,
  nr16            TEXT,
  perito_nome     TEXT,
  perito_crea     TEXT,
  perito_email    TEXT,
  status          TEXT DEFAULT 'aguardando',    -- aguardando | diligencia_realizada | laudo_gerado
  importado_em    TIMESTAMPTZ DEFAULT NOW(),
  atualizado_em   TIMESTAMPTZ DEFAULT NOW()
);

-- ── TABELA: diligencias ───────────────────────────────────────
-- Dados coletados no iPad durante a diligência in loco
CREATE TABLE IF NOT EXISTS diligencias (
  id                   UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  processo_id          UUID NOT NULL REFERENCES processos(id) ON DELETE CASCADE,
  partes_reclamante    JSONB DEFAULT '[]',  -- [{nome, mister}]
  partes_reclamada     JSONB DEFAULT '[]',  -- [{nome, mister, admissao, documento}]
  func_reclamante      TEXT,
  admissao_reclamante  DATE,
  demissao_reclamante  DATE,
  func_reclamada       TEXT,
  admissao_reclamada   DATE,
  demissao_reclamada   DATE,
  ativ_autor           TEXT,
  ativ_empresa         TEXT,
  epis                 TEXT,
  treinamentos         TEXT,
  agentes              JSONB DEFAULT '[]',  -- [{id, nome, nr, anexo, tipo}]
  ocorrencia           BOOLEAN DEFAULT FALSE,
  ocorrencia_texto     TEXT,
  salvo_em             TIMESTAMPTZ DEFAULT NOW(),
  sincronizado_em      TIMESTAMPTZ DEFAULT NOW(),
  UNIQUE (processo_id)
);

-- ── TABELA: textos_padrao ─────────────────────────────────────
-- Textos padrão cadastrados na plataforma web
-- chave: identificador único da seção (ex: topico1_insalubridade, ag_nr15_a1_met)
CREATE TABLE IF NOT EXISTS textos_padrao (
  id            UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  chave         TEXT NOT NULL UNIQUE,  -- ex: topico1_insalubridade
  conteudo_html TEXT,                  -- HTML formatado
  atualizado_em TIMESTAMPTZ DEFAULT NOW()
);

-- ── TABELA: imagens_agentes ───────────────────────────────────
-- Imagens de cada agente para o tópico 7
-- Armazena como base64 ou URL do Storage do Supabase
CREATE TABLE IF NOT EXISTS imagens_agentes (
  id            UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  agente_id     TEXT NOT NULL UNIQUE,  -- ex: ag_nr15_a1
  nome_arquivo  TEXT,
  mime_type     TEXT DEFAULT 'image/png',
  dados_base64  TEXT,                  -- base64 da imagem
  storage_url   TEXT,                  -- URL do Supabase Storage (alternativa ao base64)
  atualizado_em TIMESTAMPTZ DEFAULT NOW()
);

-- ── TRIGGERS: atualizado_em automático ───────────────────────
CREATE OR REPLACE FUNCTION atualizar_timestamp()
RETURNS TRIGGER AS $$
BEGIN
  NEW.atualizado_em = NOW();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE OR REPLACE TRIGGER trg_processos_updated
  BEFORE UPDATE ON processos
  FOR EACH ROW EXECUTE FUNCTION atualizar_timestamp();

CREATE OR REPLACE TRIGGER trg_textos_updated
  BEFORE UPDATE ON textos_padrao
  FOR EACH ROW EXECUTE FUNCTION atualizar_timestamp();

CREATE OR REPLACE TRIGGER trg_imagens_updated
  BEFORE UPDATE ON imagens_agentes
  FOR EACH ROW EXECUTE FUNCTION atualizar_timestamp();

-- ── ROW LEVEL SECURITY (RLS) ──────────────────────────────────
-- Habilita RLS em todas as tabelas (acesso somente autenticado)
ALTER TABLE processos       ENABLE ROW LEVEL SECURITY;
ALTER TABLE diligencias     ENABLE ROW LEVEL SECURITY;
ALTER TABLE textos_padrao   ENABLE ROW LEVEL SECURITY;
ALTER TABLE imagens_agentes ENABLE ROW LEVEL SECURITY;

-- Políticas: usuário autenticado tem acesso total às suas tabelas
-- (como o sistema é de uso individual, permite tudo para authenticated)
CREATE POLICY "acesso_autenticado" ON processos
  FOR ALL TO authenticated USING (true) WITH CHECK (true);

CREATE POLICY "acesso_autenticado" ON diligencias
  FOR ALL TO authenticated USING (true) WITH CHECK (true);

CREATE POLICY "acesso_autenticado" ON textos_padrao
  FOR ALL TO authenticated USING (true) WITH CHECK (true);

CREATE POLICY "acesso_autenticado" ON imagens_agentes
  FOR ALL TO authenticated USING (true) WITH CHECK (true);

-- ── ÍNDICES ───────────────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_processos_numero     ON processos(numero);
CREATE INDEX IF NOT EXISTS idx_processos_data       ON processos(data_pericia);
CREATE INDEX IF NOT EXISTS idx_processos_status     ON processos(status);
CREATE INDEX IF NOT EXISTS idx_diligencias_processo ON diligencias(processo_id);
CREATE INDEX IF NOT EXISTS idx_textos_chave         ON textos_padrao(chave);
CREATE INDEX IF NOT EXISTS idx_imagens_agente       ON imagens_agentes(agente_id);

-- ── COMENTÁRIOS ───────────────────────────────────────────────
COMMENT ON TABLE processos       IS 'Processos importados da planilha Excel do sistema de gestão';
COMMENT ON TABLE diligencias     IS 'Dados coletados no iPad durante a diligência pericial';
COMMENT ON TABLE textos_padrao   IS 'Textos padrão de cada seção do laudo, cadastrados na plataforma web';
COMMENT ON TABLE imagens_agentes IS 'Imagens dos agentes NR-15/NR-16 para o tópico 7 do laudo';

-- ── VERIFICAÇÃO ───────────────────────────────────────────────
SELECT table_name, pg_size_pretty(pg_total_relation_size(quote_ident(table_name))) AS tamanho
FROM information_schema.tables
WHERE table_schema = 'public'
ORDER BY table_name;
