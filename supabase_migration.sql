-- Tabela de premissas (projeções e transferências entre empresas)
-- Cada linha representa uma chave única (ex: "forecast_LAG_2026", "transferencias_2026")
CREATE TABLE IF NOT EXISTS premissas (
  chave TEXT PRIMARY KEY,
  dados JSONB
);

-- Acesso público sem autenticação (todos podem ler e gravar)
ALTER TABLE premissas ENABLE ROW LEVEL SECURITY;

CREATE POLICY "public_read"   ON premissas FOR SELECT USING (true);
CREATE POLICY "public_insert" ON premissas FOR INSERT WITH CHECK (true);
CREATE POLICY "public_update" ON premissas FOR UPDATE USING (true) WITH CHECK (true);
CREATE POLICY "public_delete" ON premissas FOR DELETE USING (true);
