"""Testes do parâmetro `extra_text` de `analyze_pdf` -- texto que o usuário
cola/digita na tela do webapp quando a legenda de GIB não foi extraída
automaticamente da própria planta (ver caixa "Informações adicionais" e o
botão "Reprocessar" no webapp)."""

from __future__ import annotations

from reportlab.pdfgen import canvas

from pdf_takeoff.pdf_to_input import analyze_pdf


def _generate_plan_without_wall_legend(path: str) -> None:
    """Planta minima, com o titulo/escala certos pra passar pelo filtro de
    'prancha de planta de piso', mas SEM a nota 'Wall Linings' -- deve
    disparar o aviso de que o programa nao achou a legenda, a nao ser que o
    texto seja suprido via `extra_text`."""
    c = canvas.Canvas(path, pagesize=(1684, 1191))
    c.drawString(100, 1100, "Ground Floor Plan")
    c.drawString(100, 1080, "1:100 @ A2")
    c.rect(100, 100, 400, 300, stroke=1, fill=0)
    c.drawCentredString(300, 250, "Sala Teste")
    c.showPage()
    c.save()


def test_missing_wall_legend_warns_without_hint(tmp_path):
    pdf_path = tmp_path / "plan.pdf"
    _generate_plan_without_wall_legend(str(pdf_path))

    result = analyze_pdf(str(pdf_path))

    assert any("Wall Linings" in w for w in result.warnings)


def test_pasted_hint_supplies_missing_wall_legend(tmp_path):
    """Mesma planta, mas com a legenda que faltava colada em `extra_text`
    (como o usuário faria na caixa de 'Informações adicionais' do webapp) --
    o aviso de legenda não encontrada deve sumir, porque o mesmo leitor
    determinístico (`extract_wall_default`) agora acha a nota dentro do
    texto colado."""
    pdf_path = tmp_path / "plan.pdf"
    _generate_plan_without_wall_legend(str(pdf_path))

    hint = "Wall Linings\n10mm Gibboard internal linings (Aqualine to wet areas)"
    result = analyze_pdf(str(pdf_path), extra_text=hint)

    assert not any("Wall Linings" in w for w in result.warnings)
