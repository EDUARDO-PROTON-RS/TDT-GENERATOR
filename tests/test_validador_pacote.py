"""Testes do validador de pacote TDT (detecção de OOXML não-canônico).

O ADMS recusa pacotes gravados pelo openpyxl ("Invalid TDI file format"):
faltam `xl/sharedStrings.xml` e os comentários ficam em `xl/comments/...`
(o Excel nativo usa `xl/comments1.xml`). O validador classifica o pacote sem
precisar do ADMS nem do Excel — roda em CI.
"""

from __future__ import annotations

import io
import zipfile

import openpyxl

from tdt.validador_pacote import RelatorioPacote, validar_pacote_tdt


def _xlsx_openpyxl() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Signal Name"
    ws["A2"] = "FWB_AL13_50F1"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _xlsx_nativo_simulado() -> bytes:
    """Pacote mínimo com os marcadores de um arquivo salvo pelo MS Excel."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("xl/workbook.xml", "<workbook/>")
        z.writestr("xl/sharedStrings.xml", "<sst/>")
        z.writestr("xl/comments1.xml", "<comments/>")
    return buf.getvalue()


def test_openpyxl_eh_nao_nativo():
    rel = validar_pacote_tdt(_xlsx_openpyxl())
    assert isinstance(rel, RelatorioPacote)
    assert rel.nativo is False
    assert rel.tem_shared_strings is False
    assert rel.problemas  # tem ao menos um problema descrito


def test_nativo_simulado_eh_aceito():
    rel = validar_pacote_tdt(_xlsx_nativo_simulado())
    assert rel.nativo is True
    assert rel.tem_shared_strings is True
    assert rel.comentarios_nao_nativos == ()
    assert rel.problemas == ()


def test_comentarios_no_caminho_openpyxl_sao_detectados():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("xl/sharedStrings.xml", "<sst/>")
        z.writestr("xl/comments/comment1.xml", "<comments/>")
    rel = validar_pacote_tdt(buf.getvalue())
    assert rel.comentarios_nao_nativos == ("xl/comments/comment1.xml",)
    assert rel.nativo is False


def test_aceita_path_alem_de_bytes(tmp_path):
    p = tmp_path / "saida.xlsx"
    p.write_bytes(_xlsx_openpyxl())
    rel = validar_pacote_tdt(p)
    assert rel.nativo is False


def test_template_oficial_do_projeto_e_nao_nativo(template_dnp3_path):
    """O template embarcado já passou pelo openpyxl: comprova o risco real."""
    rel = validar_pacote_tdt(template_dnp3_path)
    assert rel.nativo is False
