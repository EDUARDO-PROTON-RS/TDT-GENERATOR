# Aproveitamento do Handoff (ADMS-TDT-GENERATOR v1) → Projeto TDT v2

Data: 25/jun/2026
Fonte: `HANDOFF_COMPLETO.md` + código do projeto antigo `adms-tdt-generator`
(FastAPI + React) comparado com o projeto atual `tdt` (pipeline determinístico +
PySide6).

> Objetivo desta revisão: identificar o que do projeto antigo é **útil e
> aproveitável** para melhorar o atual, com foco nos três pontos levantados —
> **formatação do TDT**, **coleta de dados** e **double-bit vs single-bit** —, e
> propor um plano de aplicação respeitando o contrato DOX (TDD obrigatório,
> specs em `docs/superpowers/specs/`).

---

## 0. Veredito rápido

| Tema | Estado no projeto atual | Achado do handoff | Prioridade |
|------|-------------------------|-------------------|------------|
| **Formatação / aceitação no ADMS** | `engine_tdt.salvar()` usa `wb.save()` puro do openpyxl | ⚠️ **ADMS REJEITA** OOXML do openpyxl: *"Invalid TDI file format"*. Exige re-save nativo (MS Excel COM) | 🔴 **CRÍTICA** |
| **Coleta de dados (ler tudo)** | `identificador` percorre TODAS as sheets por conteúdo; `ler_rows` sem cap → **já robusto contra o bug §8** | Bug clássico do antigo: pular abas / dedup agressivo / truncar linhas → sinais somem (FredW 1222→1429) | 🟢 Já resolvido (resíduo menor) |
| **Output de comando (single/multi/data type)** | `Output Data Type` hardcoded `"SingleBit"` | `COMMAND_PRESETS` (trip_close, single_close, latch…) + coluna `CONTROL CODE` da lista padrão | 🟠 Alta |
| **Double-bit de status** | Heurística por endereços consecutivos + mesma sigla | Handoff **não** trata double-bit (lacuna dos dois) | 🟡 Revisar |
| **Limpar Signal Custom ID** | Não há lógica explícita (vaza GUID do template?) | ADMS gera o GUID; deixar vazio (`None`) ou dá conflito | 🟠 Alta |
| **Faixa banded / ListObject ref** | `_expandir_tabela()` estende ref via openpyxl | openpyxl não redimensiona ListObject de verdade; só o Excel COM resize | 🟡 Mitigado |
| **Abas-lixo conhecidas** | — | `_SKIP_SHEETS`: Capa, Calculados, Slot, Saca, Sumário, Config… | 🟢 Fácil |
| **Destaque de incertos (cor amarela)** | Revisão é na UI; TDT não destaca | `FFF2CC` nas linhas não-ALTA preservando estilo | 🟢 Opcional |
| **IA** | SP2 em espera | Conclusão: só **local/offline**; LLM nunca marca ALTA (teto 85) | 🟢 Já alinhado |

---

## 1. 🔴 FORMATAÇÃO — o achado que muda tudo

### O problema
O parser TDI do ADMS é **estrito**. O openpyxl grava OOXML não-canônico:
- comentários em `xl/comments/comment1.xml` (Excel usa `xl/comments1.xml`);
- **não gera** `xl/sharedStrings.xml` (usa strings inline);
- worksheets sem `codeName`, `xr:uid`, `mc:Ignorable`, declaração XML.

Resultado em campo (projeto antigo): rejeição com
*"Invalid TDI file format. Please use official Telemetry Data Template document
and MS Excel or save file in MS Excel prior to import."*

### O estado atual do nosso projeto
`src/tdt/engine_tdt.py::salvar()` faz **exatamente** `wb.save(destino)`. Ou seja,
**os TDTs que geramos hoje têm grande chance de ser recusados pelo ADMS** — e o
projeto não tem nenhuma consciência disso (sem validação, sem aviso).

### A solução validada em campo (projeto antigo)
`backend/excel_native.py`: re-salvar via **MS Excel real** (pywin32, `DispatchEx`,
`FileFormat=51`), numa thread com `CoInitialize` + lock global. O Excel reescreve
o pacote no formato canônico e, de quebra, redimensiona os ListObjects.
Confirmado em campo: após o fix, ADMS reportou *"No errors or warnings while
parsing DNP3_DiscreteSignals/AnalogSignals"*.

### Como aplicar (respeitando que somos cross-platform / Linux+PySide6)
1. **Porte opcional `excel_native`**: novo módulo `src/tdt/exportador_nativo.py`.
   - No Windows com Excel → re-save COM (os engenheiros da RGE rodam Windows+ADMS).
   - Sem Excel/fora do Windows → fallback para `wb.save()` **+ aviso explícito**
     na UI/CLI de que o arquivo pode ser recusado e precisa ser reaberto/salvo no
     Excel antes de importar.
2. **Validador de TDT** (`validar_pacote_tdt(bytes)`): abre o `.xlsx` como ZIP e
   checa presença de `xl/sharedStrings.xml` e `xl/comments1.xml` (caminho nativo).
   Roda sempre, é barato e dá o sinal verde/vermelho sem depender de ADMS.
   Vira teste determinístico (não precisa de Excel para rodar o teste do
   validador — só precisamos de um fixture nativo e um fixture openpyxl).
3. **Health flag** `excel_nativo: bool` exposto na UI (tela inicial/config),
   como o `/api/health → excelNative` do projeto antigo.

> Nota honesta: o re-save COM **não roda neste ambiente Linux** nem em CI. O
> teste de CI cobre o **validador** e o **fallback com aviso**; o caminho COM em
> si é testado manualmente no Windows do usuário. Isso é melhor que o silêncio
> atual.

---

## 2. 🟢 COLETA DE DADOS — já robusto, com resíduos menores

### Achado do handoff (§8): 3 falhas que faziam sinais sumirem (no ANTIGO)
1. Processava abas "preferidas" **OU** "outras", nunca as duas → pulava abas.
2. Dedup por `utr_id` (não-único no formato estruturado) → dropava sinais.
3. `_scan_grid` truncava em 3000 linhas / 30 colunas.

### Estado real do projeto atual (verificado no código)
`identificador.classificar()` percorre **todas** as `workbook.sheetnames` e
classifica cada uma por **conteúdo** (`_eh_sheet_dados`: precisa de coluna de
inteiros + coluna de texto). `ler_rows(ws)` sem `max_rows` lê a sheet **inteira**
— o `_MAX_SCAN=60` é só amostragem de **estrutura** (achar header / decidir se é
sheet de dados), não corta os dados. Logo, **os três bugs §8 não existem aqui** —
a arquitetura é superior à do antigo (detecção por conteúdo > keyword).

### Resíduos que ainda valem (baixa prioridade)
- **Header além da linha 60:** se uma sheet legítima tiver >60 linhas de
  metadados antes do header, `_eh_sheet_dados`/`_header_por_densidade` podem
  errar. Improvável, mas adicionar a contagem `sinais extraídos vs linhas por
  aba` na `auditoria.py` é a métrica que pega qualquer regressão de cobertura.
- **Abas-lixo explícitas:** a detecção por conteúdo já descarta Capa/Sumário etc.
  Uma `_SKIP_SHEETS` em `config.py` é defensiva, não essencial.
- **Manter** a detecção de coluna por embedding (`analise_colunas`) — é a maior
  vantagem sobre o `_extract_structured` por keyword do antigo.

---

## 3. 🟠 SINGLE-BIT vs DOUBLE-BIT vs COMANDO

Este é o ponto que o usuário pediu para revisar com cuidado. Há **três
sub-questões distintas** que costumam ser confundidas:

### 3.1 Status double-bit (entrada)
- **Atual:** `normalizador_estrutural.corrigir()` agrupa por `(módulo, sigla)`;
  se houver exatamente 2 registros com endereços **consecutivos** (100, 101),
  vira um sinal `is_double_bit=True` com `indices=(100,101)`. Caso contrário, vai
  para revisão (`endereco_duplicado`).
- **Handoff:** não trata double-bit (lacuna). Logo, **não há regra de campo a
  importar** — a heurística atual é nossa melhor referência.
- **Risco identificado:** a fonte do double-bit hoje é só "endereços
  consecutivos". Isso é frágil:
  - 2 sinais distintos com a mesma sigla por engano e endereços vizinhos seriam
    fundidos por engano;
  - um par 52a/52b nem sempre vem com a mesma sigla.
- **Recomendação:** complementar a heurística com **conhecimento de domínio por
  sigla** — siglas de posição de disjuntor/chave (52, seccionadoras) são
  tipicamente double-bit no ADMS. Como a planilha `Pontos Padrao ADMS` **não tem
  coluna de Input Data Type**, isso exige enriquecer a base padrão (ver 3.4) ou
  uma lista curada de siglas double-bit em `config.py`. Manter o fallback por
  endereços consecutivos como segundo sinal, e mandar para revisão quando os dois
  divergirem.

### 3.2 Output Data Type / formato de comando (saída) — **lacuna real**
- **Atual:** `engine_tdt._valores()` cravado em
  `"Output Data Type": "SingleBit"` para todo comando, sem control codes nem
  command times.
- **Projeto antigo:** `COMMAND_PRESETS` com conhecimento real extraído de TDTs:

  | preset | controlCodes | commandTimes | outputDataType | coords |
  |--------|--------------|--------------|----------------|--------|
  | trip_close | `TripPulseOn;ClosePulseOn` | `0.25;0.25` | SingleCoord | 2 |
  | close_close | `ClosePulseOn;ClosePulseOn` | `0.25;0.25` | MultiCoord | 2 |
  | single_close | `ClosePulseOn` | `0.25` | SingleCoord | 1 |
  | single_trip | `TripPulseOn` | `0.25` | SingleCoord | 1 |
  | latch | `LatchOn;LatchOff` | — | MultiCoord | 2 |

  E a coluna **`CONTROL CODE`** existe na nossa `Pontos Padrao ADMS_v1.xlsx`
  (`DiscreteSignals`) — **mas o `lista_padrao.py` atual não a lê**.
- **Recomendação:**
  1. Adicionar `control_code: str | None` ao `SinalPadrao` (ler coluna
     `CONTROL CODE`).
  2. Derivar `Output Data Type` / `Output Coordinates` / control codes a partir
     do `control_code` da sigla, com presets como no antigo. Isso resolve
     comando de abre/fecha (52) corretamente, em vez do `SingleBit` fixo.

### 3.3 Pareamento D+C
- **Atual:** `dc_pairer.py` pareia Input+Output do mesmo `(módulo, sigla)` →
  ReadWrite. Consistente com o DOX (`pareamento D+C por sigla+módulo`).
  **Sem mudança necessária**, mas alimenta o 3.2 (o ReadWrite resultante precisa
  do output data type correto).

### 3.4 Enriquecer a base padrão (origem da verdade)
A planilha tem colunas que **não lemos** e que carregam exatamente a informação
de tipo que estamos adivinhando: `CONTROL CODE`, e na aba
`DMS Signal Explanation`: `Point type`, `Point category`, `Invert In/Out bits`.
Ler essas colunas move a decisão de "heurística" para "dado de origem".

---

## 4. O que NÃO importar (becos sem saída do handoff)

- `build_fix_da.py` (COM cross-workbook copy via ActiveSheet) — **corrompe** o
  template. O antigo reconstruiu via `build_template_from_ura.py`.
- IA paga (Groq/Gemini) — decisão final foi **só local**. Nosso SP2 já está em
  espera; manter offline-first.
- `gemini-2.0-flash` (cota zero), Gemma via API (instável 500).
- Tentar fazer o openpyxl gerar OOXML "compatível" na mão — fútil; é o Excel COM.

---

## 5. Plano de aplicação proposto (com TDD, por SP)

Ordenado por valor/risco. Cada item vira spec em `docs/superpowers/specs/` antes
do código (contrato DOX).

1. **SP-fmt-1 — Validador de pacote TDT** (🔴, baixo risco, testável em CI)
   - `validar_pacote_tdt(path|bytes) -> RelatorioPacote` (checa sharedStrings +
     comments nativos). Teste com 2 fixtures (nativo vs openpyxl).
2. **SP-fmt-2 — Exportador nativo opcional + fallback com aviso** (🔴)
   - Porta `excel_native` como `exportador_nativo.py`; CLI/UI avisam quando o
     arquivo é "não-nativo".
3. **SP-cmd-1 — Control codes / Output Data Type por sigla** (🟠)
   - Ler `CONTROL CODE`; presets de comando; corrige single/multi coord.
4. **SP-data-1 — Auditoria de cobertura de coleta** (🟢, defensivo)
   - Contagem `sinais/linhas` por aba na `auditoria`; `_SKIP_SHEETS` opcional.
5. **SP-db-1 — Double-bit por domínio + fallback por endereço** (🟡)
   - Lista curada de siglas double-bit; cruzar com heurística atual.
6. **SP-id-1 — Limpar Signal Custom ID** (🟠) — garantir `None` na saída.

---

## 6. Pontos fortes do projeto atual (manter)

- Detecção de coluna por **conteúdo/embedding** (mais robusta que keyword).
- Pipeline determinístico em camadas + calibração + roteador com gap dinâmico.
- 3 scorers paralelos (TF-IDF + FAISS + fuzzy) com mescla ponderada.
- TDD com 234 testes; arquitetura SRP via `contracts.py`.
- Normalização N0-N5 com extração de contexto estrutural antes do colapso.

Estes superam o motor de regras de 6 camadas do projeto antigo. O valor do antigo
está **na ponta de saída** (aceitação no ADMS) e **no conhecimento de domínio de
comando**, não no classificador.
