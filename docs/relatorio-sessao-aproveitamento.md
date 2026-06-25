# Relatório Técnico Completo — Sessão de Aproveitamento do Handoff ADMS

Data: 25/jun/2026
Branch: `claude/busy-dirac-t0wu6e`
Repositório: `eduardo-proton-rs/tdt-generator`

Este documento consolida, com profundidade técnica: (A) **todas as mudanças**
feitas nesta sessão, do início ao fim; (B) o que **já foi implementado** e o que
**falta**; (C) as **diferenças arquiteturais** entre o projeto antigo
(`ADMS-TDT-GENERATOR`, FastAPI+React, doravante **v1**) e o projeto atual
(`ProjetoTDT-V2` / pacote `tdt`, pipeline determinístico + PySide6, **v2**).

---

# PARTE A — Linha do tempo das mudanças (início → fim)

A sessão começou com o repositório `TDT-GENERATOR` **vazio** (só `.git`, sem
commits). Sequência cronológica:

## A.0 — Tentativa de clone bloqueada
- Pedido inicial: clonar `github.com/viniciusstoffel/ProjetoTDT-V2.git`.
- **Bloqueado (403)**: o escopo de GitHub desta sessão é restrito a
  `eduardo-proton-rs/tdt-generator`; o proxy recusa qualquer outro repositório,
  inclusive clone HTTPS público. Solução adotada: o usuário enviou o código como
  zip.

## A.1 — `b5f24bb` Import ProjetoTDT-V2 codebase
- Importação do zip `ProjetoTDTV2main.zip` → **149 arquivos, 20.017 inserções**.
- Conteúdo: `src/tdt/` (pipeline + UI), `tests/` (41 arquivos de teste),
  `docs/` (specs/plans superpowers, templates `.xlsx`, inputs reais), `bench/`,
  `scripts/`, `pyproject.toml`, `AGENTS.md` (hierarquia DOX).
- Commit raiz do projeto no branch de trabalho.

## A.2 — `7c43051` docs: revisão de aproveitamento do handoff ADMS v1
- Após ler `HANDOFF_COMPLETO.md` (estratégico) + código-fonte do **v1**.
- Criado `docs/aproveitamento-handoff-adms-v1.md` (208 linhas): tabela-veredito,
  análise dos 3 eixos pedidos (formatação, coleta de dados, single/double-bit),
  becos sem saída a não repetir, plano priorizado por SP, e pontos fortes do v2
  a manter.
- **Correção factual no meio do caminho**: a seção de "coleta de dados" foi
  reescrita após ler `identificador.py` — o v2 **já** percorre todas as sheets
  por conteúdo e lê linhas sem cap, portanto **não** sofre do bug §8 do v1. O
  item foi rebaixado de 🟠 Alta para 🟢 Já resolvido.

## A.3 — `81644b5` feat: validador de pacote TDT  ← **única mudança de código**
TDD (RED→GREEN). Arquivos:
- `src/tdt/validador_pacote.py` (+77): `validar_pacote_tdt(path|bytes) ->
  RelatorioPacote`. Abre o `.xlsx` como ZIP e checa marcadores que o parser TDI
  do ADMS exige.
- `tests/test_validador_pacote.py` (+76): 5 testes; rodam em CI **sem Excel**
  (fixtures sintéticos via `zipfile` + um gerado por openpyxl + o template real).
- `src/tdt/cli.py` (+9): após `wb.save()`, valida e imprime **AVISO** quando o
  pacote é não-nativo (risco "Invalid TDI file format"), orientando reabrir no
  MS Excel.
- `src/tdt/AGENTS.md` (DOX pass): nota sobre o risco do `wb.save()` e o validador.

**Detalhe técnico do discriminador** (verificado empiricamente nesta sessão):
| Marcador | MS Excel nativo | openpyxl |
|----------|-----------------|----------|
| `xl/sharedStrings.xml` | presente | **ausente** (strings inline) |
| comentários | `xl/comments1.xml` | `xl/comments/comment1.xml` (subpasta) |

Regra: `nativo = tem_sharedStrings and não-há-comentários-em-xl/comments/`.

**Achado confirmado por teste**: o próprio `docs/dnp3_template.xlsx` embarcado
**já é não-nativo** (sem sharedStrings; comentários em `xl/comments/`). Logo,
todo TDT gerado a partir dele tende a ser recusado pelo ADMS — o que torna o
exportador nativo (pendência SP-fmt-2) crítico.

**Resultado da suíte**: `240 passed`. Os 7 erros de coleta são exclusivamente
`sentence-transformers` (encoder) e PySide6 (UI) ausentes neste ambiente
Linux/CI — não relacionados à mudança.

## A.4 — `29f601a` docs: lista de pendências
- `docs/pendencias-aproveitamento-handoff.md`: só o que falta, agrupado por
  prioridade (🔴/🟠/🟡/🟢) + seção "não fazer".

## A.5 — `21bf5fc` docs: adiciona pendência SP-fuzzy-1
- Após ler `TECNICO_COMO_FOI_FEITO.md` (detalhe de implementação do v1).
- Achado **novo, verificado no código e nos dados**: `matchers/fuzzy_match.py:31`
  dá boost quando a sigla aparece literal nos tokens, **sem guard de tamanho**.
  A `Pontos Padrao ADMS_v1.xlsx` tem siglas reais de 1 caractere (`N`, `S`, `P`,
  `Q`, `V`). É o mesmo padrão que causou falso positivo confirmado no v1 (`_A`,
  `_B`). Item adicionado às pendências (🟠).

---

# PARTE B — Estado: implementado vs. falta

## ✅ Implementado nesta sessão
1. **Importação completa do código v2** para o repositório.
2. **Revisão técnica documentada** (`aproveitamento-handoff-adms-v1.md`).
3. **SP-fmt-1 — Validador de pacote TDT** (código + testes + CLI + DOX).
4. **Backlog priorizado** (`pendencias-aproveitamento-handoff.md`).

## ⏳ Falta (ordenado por valor/risco)

### 🔴 SP-fmt-2 — Exportador nativo (MS Excel COM) + fallback
- Portar `excel_native.py` do v1 → `src/tdt/exportador_nativo.py`.
- Windows+Excel: `DispatchEx("Excel.Application")`, `FileFormat=51`, thread com
  `CoInitialize`/`CoUninitialize` + `threading.Lock` global; `tbl.Resize()` dos
  ListObjects para cobrir todas as linhas escritas.
- Fora do Windows/sem Excel: fallback `wb.save()` **com aviso** (validador já
  sinaliza). Expor `excel_nativo: bool` na UI.
- Sem isso, **o ADMS recusa 100%** dos arquivos (confirmado em campo no v1). O
  caminho COM não roda em CI Linux; testa-se fallback + integração com validador.

### 🟠 SP-cmd-1 — Control codes / Output Data Type por sigla
- `lista_padrao.py`: ler a coluna **`CONTROL CODE`** (existe na planilha, hoje
  ignorada); adicionar `control_code` ao `SinalPadrao`.
- `engine_tdt.py`: derivar `Output Data Type`/`Output Coordinates`/control codes
  por presets (trip_close, close_close, single_close, single_trip, latch) — hoje
  cravado em `"SingleBit"`.

### 🟠 SP-id-1 — Limpar Signal Custom ID
- Garantir Signal/Remote Point Custom ID **vazio (`None`)** na saída — o ADMS
  gera o GUID; herdar o do template causa conflito.

### 🟠 SP-fuzzy-1 — Guard de sigla 1-char no boost literal
- `fuzzy_match.py:31`: `boost = _BOOST_SIGLA if len(sigla) > 1 and sigla.upper()
  in tokens else 0.0`. Teste: descrição com token de 1 letra ("Fase A") + sigla
  `N`/`S` no corpus não pode ganhar boost.

### 🟡 SP-db-1 — Double-bit por domínio + fallback por endereço
- Hoje só heurística "endereços consecutivos + mesma sigla"
  (`normalizador_estrutural.py`). Adicionar lista curada de siglas double-bit
  (52, seccionadoras) em `config.py`; cruzar com a heurística; divergência →
  revisão.

### 🟡 SP-data-2 — Enriquecer a base padrão
- Ler colunas hoje ignoradas da aba `DMS Signal Explanation`: `Point type`,
  `Point category`, `Invert In/Out bits` — move decisões de heurística para dado
  de origem.

### 🟢 SP-data-1 — Auditoria de cobertura de coleta
- Contagem `sinais extraídos vs linhas por aba` na `auditoria.py`;
  `_SKIP_SHEETS` opcional em `config.py`.

### 🟢 SP-fmt-3 — Destaque de incertos no TDT (opcional)
- Pintar linhas não-ALTA de amarelo (`FFF2CC`) preservando estilo (como no v1).

### ⛔ Não fazer (becos sem saída comprovados no v1)
- COM cross-workbook copy via `ActiveSheet` (corrompe o template — apagava a
  DiscreteSignals).
- Corrigir OOXML do openpyxl na mão (paths mudam a cada versão; inviável).
- IA paga (Groq/Gemini): cota estoura em listas grandes; manter offline-first.
- Ollama em CPU sem GPU (90s/8 sinais, acurácia ruim).

---

# PARTE C — Diferenças entre os projetos (v1 × v2)

## C.1 Visão de 30 mil pés

| Dimensão | v1 (ADMS-TDT-GENERATOR) | v2 (ProjetoTDT-V2 / `tdt`) |
|----------|-------------------------|-----------------------------|
| Arquitetura | Web app: FastAPI (backend 8077) + React/Vite (front 5180) | App desktop: pipeline Python puro + UI PySide6 |
| Distribuição | `ABRIR.BAT` sobe back+front, abre Chrome `--app` | `python -m tdt.cli` / `python -m tdt.ui_main` |
| Paradigma de geração | **Clonador**: tokeniza linhas de TDTs reais e faz `str.replace()` | **Composição**: classifica sinal → sigla, preenche template por display name |
| Classificação | Motor de **6 camadas determinísticas** + IA local opcional | **3 scorers paralelos** (TF-IDF + FAISS + fuzzy) → mescla → 7 regras → roteador |
| Base de conhecimento | `sigla_index.json` (2994+447 linhas pré-tokenizadas, 98MB de origem) | `Pontos Padrao ADMS_v*.xlsx` (≈692 disc + 62 ana) lido em runtime |
| Saída ADMS-aceita | **Sim** (re-save Excel COM, validado em campo) | **Não garantido** (openpyxl puro; SP-fmt-2 pendente) |
| Estado | Maduro/produção, monolito | Modular/SRP, TDD (234+ testes), em evolução por SPs |
| Plataforma | Windows-only (depende de Excel COM) | Cross-platform (COM seria opcional) |

## C.2 Geração da TDT — clonagem vs. composição (a diferença central)

**v1 (clonagem):** capturou de TDTs reais cada SIGLA como uma **linha inteira
pré-tokenizada** (43/61/48 valores) com placeholders `<<PREFIX>>`, `<<ALIAS>>`,
`<<MODULE>>`, `<<DEVICE>>`, `<<N>>`. Gerar = `str.replace()` dos tokens. Vantagem:
fidelidade absoluta a um TDT real (round-trip ZERO diffs). Custo: depende de uma
base gigante (98MB) não recriável, e de capturar exemplos reais por sigla.

**v2 (composição):** carrega o template, localiza colunas pelo **display name
(row 4)**, e **calcula** cada campo identity-dependent a partir do `SignalRecord`
classificado (`_nome_hierarquico`, `_aor_group`, `_remote_unit`,
`_device_mapping`, etc.). Vantagem: não precisa da base de 98MB, escala para
qualquer sigla da lista padrão. Custo: cada campo derivado é uma regra que pode
divergir de um TDT real — e o pacote sai não-nativo (daí SP-fmt-2/SP-fmt-1).

## C.3 Classificação textual descrição → sigla

**v1**: pipeline de regras (handoff §6), da mais forte à mais fraca:
1. Token (sufixo `_50F1` no nome) — ALTA 100, com guard `len>1`.
2. Semântico (`Tensão.*AB→VAB`) — ALTA 92, ordem importa (REATIVA antes de ATIVA).
3. Base oficial / 4. Descrição exata — ALTA 94/96.
5. Proteção (palavra→ANSI) — MÉDIA 82. 6. Fuzzy ANSI (Jaccard+bônus) — ≤78.
7. LLM local opcional — teto 85 (nunca ALTA).
Normalização `_norm`: maiúsculas, sem acento, expansão de abreviações
word-boundary, extração de código ANSI.

**v2**: três métodos **independentes e paralelos**, calibrados e mesclados:
- `scoring/tfidf.py` (sklearn, cosseno), `scoring/vetorial.py` (FAISS sobre
  sentence-transformers, e5/MiniLM), `matchers/fuzzy_match.py` (rapidfuzz
  token_set_ratio + boost de sigla literal).
- `scoring/mescla.py` soma ponderada (pesos ~0.34/0.33/0.33 em `config.py`).
- `motor_regras.py` 7 regras de domínio ajustam scores (R1 ANSI, R2 opostos,
  R3 fase, R4 estágio, R5 comando/status, R_eq equipamento, R6 lado tensão).
- `roteador.py` cascata: fuzzy≥0.95 → e5≥0.95 → consenso+gap dinâmico → quadrante;
  dual-pass para categoria incerta.
- Normalização N0-N5 (`normalizador.py`): **N0 extrai contexto estrutural**
  (equipamento/barra/fase) do texto **bruto** antes do colapso, depois N1-N5
  canoniza. Diferencial sobre o `_norm` plano do v1.

**Conclusão**: o classificador do v2 é mais sofisticado (embeddings + calibração
+ separação Discrete/Analog por categoria). O do v1 é mais simples mas com
conhecimento de domínio explícito embutido (regras de proteção/semântica).

## C.4 Coleta de dados (leitura de listas não-padrão)

**v1**: `parse_raw_excel` detecta formato (UTR-ID via `_find_utr_column` ou
estruturado via `_extract_structured` por keywords Módulo/Tipo/Descrição). Sofreu
o **bug §8** (abas OR, dedup por chave fraca, scan truncado) → 207 sinais sumiam;
corrigido para ler todas as abas + dedup por tupla completa + limites altos.

**v2**: `identificador.py` percorre **todas** as sheets e classifica cada uma por
**conteúdo** (`_eh_sheet_dados`: precisa de coluna de inteiros + coluna de texto);
`analise_colunas.py` acha a coluna de descrição por **similaridade de embedding ×
diversidade × comprimento**, índice por sequência de inteiros, tipo por
vocabulário A/C/D. `ler_rows` lê a sheet inteira (cap de 60 é só amostragem de
estrutura). **Não tem o bug §8** — detecção por conteúdo > keyword.

## C.5 Single-bit / double-bit / comando

- **Status double-bit (entrada)**: v1 **não trata**; v2 tem
  `normalizador_estrutural.corrigir()` (endereços consecutivos + mesma sigla →
  `is_double_bit`, `indices=(100,101)`). v2 está à frente, mas a heurística é
  frágil (ver SP-db-1).
- **Comando (saída)**: v1 tem `COMMAND_PRESETS` ricos (controlCodes/commandTimes/
  outputDataType: trip_close, latch, etc.); v2 crava `"SingleBit"` (ver SP-cmd-1).
  **v1 está à frente aqui.**
- **Pareamento D+C**: ambos pareiam por (módulo, sigla) → ReadWrite (v2:
  `dc_pairer.py`).

## C.6 Formatação / aceitação no ADMS

- **v1**: `excel_native.py` re-salva via Excel COM → pacote canônico, **aceito**
  ("No errors or warnings while parsing"). `_fit_table` (openpyxl) + `tbl.Resize()`
  (COM) mantêm a faixa banded; destaque amarelo `FFF2CC` em incertos.
- **v2**: `engine_tdt.salvar()` = `wb.save()` puro → **não-nativo**, risco de
  recusa. Tem `_expandir_tabela`/`_expandir_cf`/`_expandir_dv` (openpyxl) para
  ref de tabela, conditional formatting e data validation, mas **sem re-save
  nativo**. Esta sessão adicionou o **validador** que detecta o problema; o
  exportador nativo (SP-fmt-2) ainda falta.

## C.7 IA

- **v1**: testou Groq/Gemini/Gemma/Ollama; decisão final **só local (Ollama)**,
  padrão "sem IA". Truques: índice invertido no prompt, retry backoff, teto 85.
- **v2**: SP2 (agentes LLM) **em espera** — sem spec/código. Filosofia
  offline-first idêntica.

## C.8 Tabela-resumo: quem está à frente em quê

| Tema | À frente | Por quê |
|------|----------|---------|
| Classificação textual | **v2** | embeddings + calibração + 3 métodos + N0-N5 |
| Coleta de dados | **v2** | detecção por conteúdo, sem bug §8 |
| Double-bit de status | **v2** | v1 não trata |
| Aceitação no ADMS (formatação) | **v1** | re-save Excel COM validado em campo |
| Comando (control codes/output type) | **v1** | COMMAND_PRESETS reais |
| Conhecimento de domínio explícito | **v1** | regras de proteção/semântica embutidas |
| Arquitetura/manutenção/testes | **v2** | SRP, TDD, DOX, modular |
| Portabilidade | **v2** | cross-platform (COM opcional) |

> **Síntese**: o valor a importar do v1 está **na ponta de saída** (aceitação no
> ADMS via Excel COM) e no **conhecimento de domínio de comando** (control codes),
> não no classificador — onde o v2 já é superior. Foi exatamente isso que o
> backlog desta sessão priorizou.
