# Pendências — Aproveitamento do Handoff ADMS

Data: 25/jun/2026
Base: `docs/aproveitamento-handoff-adms-v1.md`

> Lista apenas do que **falta fazer**. Cada item vira spec em
> `docs/superpowers/specs/` antes do código (contrato DOX) e exige TDD.

---

## ✅ Já feito
- **SP-fmt-1 — Validador de pacote TDT** (`src/tdt/validador_pacote.py` + CLI avisa
  quando o pacote é não-nativo). 5 testes verdes.

---

## 🔴 Crítico

### SP-fmt-2 — Exportador nativo (MS Excel COM) + fallback
- Portar `excel_native.py` do projeto antigo → `src/tdt/exportador_nativo.py`.
- Windows + Excel: re-save via COM (`DispatchEx`, `FileFormat=51`, thread com
  `CoInitialize` + lock) → pacote canônico aceito pelo ADMS.
- Fora do Windows / sem Excel: fallback `wb.save()` **com aviso** (já temos o
  validador para sinalizar).
- Expor flag `excel_nativo: bool` na UI (tela inicial/config).
- Testes: fallback + integração com o validador (o caminho COM roda só no
  Windows do usuário, não em CI).

---

## 🟠 Alta

### SP-cmd-1 — Control codes / Output Data Type por sigla
- `lista_padrao.py`: ler a coluna **`CONTROL CODE`** (já existe na planilha,
  hoje ignorada); adicionar `control_code` ao `SinalPadrao`.
- `engine_tdt.py`: derivar `Output Data Type` / `Output Coordinates` / control
  codes a partir da sigla, com presets (trip_close, single_close, latch, …) —
  hoje está cravado em `"SingleBit"`.

### SP-id-1 — Limpar Signal Custom ID
- Garantir que o campo Signal/Remote Point Custom ID saia **vazio (`None`)** —
  o ADMS gera o GUID; herdar o do template causa conflito.

### SP-fuzzy-1 — Guard contra sigla de 1 caractere no boost literal
- `matchers/fuzzy_match.py:31`: `boost = _BOOST_SIGLA if sigla.upper() in
  tokens else 0.0` não tem guard de tamanho. O projeto antigo confirmou em
  produção falso positivo de siglas `A`/`B` casando por essa via.
- **Risco real e não hipotético**: `Pontos Padrao ADMS_v1.xlsx` tem siglas de
  1 caractere de verdade (`N`, `S`, `P`, `Q`, `V` — neutro/potências), então o
  boost pode disparar de um token solto na descrição.
- Fix: `boost = _BOOST_SIGLA if len(sigla) > 1 and sigla.upper() in tokens
  else 0.0`. Teste dedicado com descrição contendo token de 1 letra (ex.
  "Fase A") + sigla `N`/`S` no corpus, garantindo que não ganha boost.

---

## 🟡 Revisar

### SP-db-1 — Double-bit por domínio + fallback por endereço
- Hoje: só heurística "endereços consecutivos + mesma sigla"
  (`normalizador_estrutural.py`).
- Adicionar lista curada de siglas double-bit (posição de disjuntor/chave: 52,
  seccionadoras) em `config.py`; cruzar com a heurística; divergência → revisão.

### SP-data-2 — Enriquecer a base padrão
- Ler colunas hoje ignoradas da aba `DMS Signal Explanation`: `Point type`,
  `Point category`, `Invert In/Out bits` — move decisões de heurística para dado
  de origem.

---

## 🟢 Defensivo / opcional

### SP-data-1 — Auditoria de cobertura de coleta
- Contagem `sinais extraídos vs linhas por aba` na `auditoria.py` (pega
  regressão de cobertura).
- `_SKIP_SHEETS` configurável em `config.py` (Capa, Calculados, Slot, Saca,
  Sumário, Config, …) — a detecção por conteúdo já descarta a maioria.

### SP-fmt-3 — Destaque de incertos no TDT (opcional)
- Pintar linhas não-ALTA de amarelo claro (`FFF2CC`) preservando estilo, como no
  projeto antigo.

---

## ⛔ Não fazer (becos sem saída do handoff)
- COM cross-workbook copy via `ActiveSheet` (corrompe template).
- IA paga (Groq/Gemini) — manter offline-first; SP2 segue em espera.
- Forçar openpyxl a gerar OOXML "compatível" na mão — só o Excel COM resolve.
