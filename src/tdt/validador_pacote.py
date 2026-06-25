"""Valida se um pacote .xlsx está no formato canônico que o ADMS aceita.

O parser TDI do EcoStruxure ADMS é estrito e recusa pacotes OOXML gravados pelo
openpyxl ("Invalid TDI file format. Please use official Telemetry Data Template
document and MS Excel or save file in MS Excel prior to import.").

Marcadores que distinguem um pacote salvo pelo MS Excel (nativo) de um do
openpyxl:
- **`xl/sharedStrings.xml`**: o Excel sempre gera; o openpyxl usa strings inline
  e não cria esse membro. É o discriminador mais confiável.
- **caminho dos comentários**: o Excel usa `xl/comments1.xml`; o openpyxl usa
  `xl/comments/comment1.xml` (subpasta).

Este validador abre o `.xlsx` como ZIP e checa esses marcadores — não precisa do
ADMS nem do MS Excel, então roda em CI. Ele NÃO conserta o arquivo (isso é papel
do exportador nativo, que exige Excel); só sinaliza o risco de rejeição.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

_SHARED_STRINGS = "xl/sharedStrings.xml"
# comentários no estilo openpyxl: xl/comments/comment1.xml (subpasta)
_COMENTARIO_OPENPYXL = re.compile(r"^xl/comments/")


@dataclass(frozen=True)
class RelatorioPacote:
    """Diagnóstico de um pacote .xlsx quanto à aceitação pelo ADMS."""

    nativo: bool
    tem_shared_strings: bool
    comentarios_nao_nativos: tuple[str, ...]
    problemas: tuple[str, ...]


def validar_pacote_tdt(origem: str | Path | bytes) -> RelatorioPacote:
    """Classifica o pacote como nativo (aceito) ou não-nativo (risco de recusa).

    ``origem`` pode ser caminho (str/Path) ou os bytes do arquivo.
    """
    if isinstance(origem, (str, Path)):
        with zipfile.ZipFile(origem) as z:
            nomes = z.namelist()
    else:
        import io

        with zipfile.ZipFile(io.BytesIO(origem)) as z:
            nomes = z.namelist()

    tem_shared = _SHARED_STRINGS in nomes
    comentarios_nao_nativos = tuple(
        n for n in nomes if _COMENTARIO_OPENPYXL.match(n)
    )

    problemas: list[str] = []
    if not tem_shared:
        problemas.append(
            "sem xl/sharedStrings.xml — gravado pelo openpyxl (strings inline); "
            "o ADMS recusa. Reabra e salve no MS Excel antes de importar."
        )
    if comentarios_nao_nativos:
        problemas.append(
            "comentários em xl/comments/ (estilo openpyxl); o Excel nativo usa "
            "xl/comments1.xml."
        )

    return RelatorioPacote(
        nativo=not problemas,
        tem_shared_strings=tem_shared,
        comentarios_nao_nativos=comentarios_nao_nativos,
        problemas=tuple(problemas),
    )
