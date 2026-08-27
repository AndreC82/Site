"""Página local no navegador pra facilitar o upload do PDF: arrasta o arquivo,
clica em processar, baixa a planilha pré-preenchida e o PDF de conferência —
sem precisar digitar comando nenhum depois de abrir a página.

Também fecha o ciclo de ajuste: se alguma informação não for encontrada na
planta (ex.: altura de pé-direito), a página pergunta antes de gerar a
planilha; e depois de você ajustar a planilha de entrada no Excel, dá pra
enviar ela de volta pra gerar o orçamento final, sem precisar do terminal.

Rodar com: python -m pdf_takeoff.webapp
Depois abra http://localhost:5000 no navegador.
"""

from __future__ import annotations

import tempfile
import threading
import uuid
import webbrowser
from pathlib import Path

from flask import Flask, render_template_string, request, send_file

from .input_template import build_from_input
from .pdf_to_input import analyze_pdf, render_review_pdf, write_prefilled_input

app = Flask(__name__)

_JOBS_DIR = Path(tempfile.gettempdir()) / "pdf_takeoff_webapp"
_JOBS_DIR.mkdir(exist_ok=True)

_STYLE = """
  body { font-family: system-ui, sans-serif; max-width: 720px; margin: 40px auto; padding: 0 20px; color: #222; }
  h1 { font-size: 1.4rem; }
  h2 { font-size: 1.1rem; margin-top: 40px; border-top: 1px solid #ddd; padding-top: 24px; }
  .drop {
    border: 2px dashed #888; border-radius: 10px; padding: 40px 20px; text-align: center;
    margin: 20px 0; background: #fafafa; cursor: pointer;
  }
  .drop.dragover { background: #eef6ff; border-color: #2b6cb0; }
  .drop input { display: none; }
  button {
    background: #2b6cb0; color: white; border: none; padding: 10px 22px; border-radius: 6px;
    font-size: 1rem; cursor: pointer;
  }
  button:disabled { background: #aaa; cursor: default; }
  .warn { background: #fff8e1; border: 1px solid #f0c040; border-radius: 6px; padding: 12px; margin: 10px 0; }
  .ask { background: #eef6ff; border: 1px solid #2b6cb0; border-radius: 6px; padding: 16px; margin: 20px 0; }
  .ask label { display: block; margin: 10px 0 4px; font-weight: bold; }
  .ask input[type=text] { padding: 6px; width: 120px; }
  .result { background: #f0fff4; border: 1px solid #38a169; border-radius: 6px; padding: 16px; margin: 20px 0; }
  a.download { display: inline-block; margin: 6px 12px 6px 0; }
  #filename, #filename2 { font-weight: bold; }
  #status { margin-top: 10px; color: #555; }
"""

_UPLOAD_SCRIPT = """
function wireDrop(dropId, inputId, filenameId, btnId, statusId, waitMsg) {
  const drop = document.getElementById(dropId);
  const input = document.getElementById(inputId);
  const filenameEl = document.getElementById(filenameId);
  const btn = document.getElementById(btnId);
  const statusEl = document.getElementById(statusId);
  const form = drop.closest('form');

  drop.addEventListener('click', () => input.click());
  drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('dragover'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
  drop.addEventListener('drop', (e) => {
    e.preventDefault();
    drop.classList.remove('dragover');
    if (e.dataTransfer.files.length) { input.files = e.dataTransfer.files; update(); }
  });
  input.addEventListener('change', update);
  function update() {
    if (input.files.length) { filenameEl.textContent = input.files[0].name; btn.disabled = false; }
  }
  form.addEventListener('submit', () => {
    btn.disabled = true;
    btn.textContent = 'Processando...';
    if (statusEl) statusEl.textContent = waitMsg;
  });
}
"""

_PAGE = """
<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<title>Quantificador de Gesso/Pintura</title>
<style>{{ style }}</style>
</head>
<body>
<h1>Quantificador automático de Gesso/Pintura</h1>

{% if questions %}
<p>Enviei "{{ questions.pdf_name }}" e não achei a altura de pé-direito pra
{{ 'este nível' if questions.levels|length == 1 else 'estes níveis' }}. Informe pra continuar:</p>
<form method="post" action="/completar">
  <input type="hidden" name="job_id" value="{{ questions.job_id }}">
  <div class="ask">
    {% for level in questions.levels %}
    <label for="altura_{{ level }}">Altura de pé-direito — {{ level }} (m)</label>
    <input type="text" id="altura_{{ level }}" name="altura_{{ level }}" value="2.70" required>
    {% endfor %}
  </div>
  <button type="submit">Continuar</button>
</form>

{% else %}

<p>Envie o PDF da planta (formato de plantas com convenção GIB) e clique em Processar.</p>
<form id="form" method="post" enctype="multipart/form-data" action="/processar">
  <div class="drop" id="drop">
    <p>Arraste o PDF aqui, ou clique pra escolher</p>
    <p id="filename"></p>
    <input type="file" id="fileInput" name="pdf" accept="application/pdf" required>
  </div>
  <label for="hints" style="display:block; margin-top:16px; font-weight:bold;">
    Informações adicionais (opcional)
  </label>
  <p style="color:#555; font-size:0.9rem; margin:4px 0 8px;">
    Se o programa não achar a legenda de GIB sozinho, cole o texto dela aqui (ex.: copiado da
    página onde ela está), ou digite a definição direto — ex. "GBTLIC60 = 2 camadas de 10mm
    Noiseline cada lado" ou "a legenda de teto está na página 7".
  </label>
  <textarea id="hints" name="hints" rows="4" style="width:100%; padding:8px; box-sizing:border-box; font-family:inherit;"></textarea>
  <div style="margin-top:12px;">
    <button type="submit" id="submitBtn" disabled>Processar</button>
  </div>
  <div id="status"></div>
</form>

{% if result %}
<div class="result">
  <p><strong>Pronto!</strong> {{ result.n_wall }} {{ result.wall_label }}, {{ result.n_ceiling }} {{ result.ceiling_label }}.</p>
  {% if result.warnings %}
  <div class="warn">
    <strong>Avisos:</strong>
    <ul>{% for w in result.warnings %}<li>{{ w }}</li>{% endfor %}</ul>
  </div>
  {% endif %}
  <a class="download" href="/baixar/{{ result.job_id }}/entrada.xlsx">⬇ Baixar planilha de entrada (.xlsx)</a>
  <a class="download" href="/baixar/{{ result.job_id }}/conferencia.pdf">⬇ Baixar PDF de conferência</a>
  <p>Abra a planilha, compare com o PDF de conferência, e ajuste qualquer linha que estiver
  errada (metros lineares, tipo de chapa, taxas). Depois use o formulário abaixo pra gerar o orçamento final.</p>

  <details style="margin-top:16px;">
    <summary style="cursor:pointer; font-weight:bold;">Faltou alguma informação? Adicione e reprocesse</summary>
    <form method="post" action="/reprocessar" style="margin-top:10px;">
      <input type="hidden" name="job_id" value="{{ result.job_id }}">
      <textarea name="hints" rows="4" style="width:100%; padding:8px; box-sizing:border-box; font-family:inherit;"
        placeholder='Cole aqui a legenda de GIB que faltou, ou digite ex.: &quot;a legenda de teto está na página 7&quot;'></textarea>
      <div style="margin-top:10px;"><button type="submit">Reprocessar com essa informação</button></div>
    </form>
  </details>
</div>
{% endif %}

<h2>Já ajustei a planilha de entrada — gerar orçamento final</h2>
<form id="form2" method="post" enctype="multipart/form-data" action="/construir">
  <div class="drop" id="drop2">
    <p>Arraste a planilha de entrada (já ajustada) aqui, ou clique pra escolher</p>
    <p id="filename2"></p>
    <input type="file" id="fileInput2" name="entrada" accept=".xlsx" required>
  </div>
  <button type="submit" id="submitBtn2" disabled>Gerar orçamento</button>
  <div id="status2"></div>
</form>

{% if orcamento %}
<div class="result">
  <p><strong>Orçamento gerado!</strong></p>
  <a class="download" href="/baixar/{{ orcamento.job_id }}/orcamento.xlsx">⬇ Baixar planilha de orçamento final</a>
</div>
{% endif %}

{% endif %}

<script>
{{ upload_script | safe }}
wireDrop('drop', 'fileInput', 'filename', 'submitBtn', 'status', 'Isso pode levar alguns minutos numa planta grande, aguarde...');
wireDrop('drop2', 'fileInput2', 'filename2', 'submitBtn2', 'status2', 'Gerando...');
</script>
</body>
</html>
"""


def _render(**kwargs):
    return render_template_string(_PAGE, style=_STYLE, upload_script=_UPLOAD_SCRIPT, **kwargs)


@app.route("/", methods=["GET"])
def index():
    return _render(result=None, questions=None, orcamento=None)


def _read_hints(job_dir: Path) -> str:
    hints_path = job_dir / "hints.txt"
    return hints_path.read_text(encoding="utf-8") if hints_path.exists() else ""


def _append_hints(job_dir: Path, new_text: str) -> None:
    new_text = (new_text or "").strip()
    if not new_text:
        return
    existing = _read_hints(job_dir)
    combined = f"{existing}\n{new_text}".strip() if existing else new_text
    (job_dir / "hints.txt").write_text(combined, encoding="utf-8")


def _read_height_overrides(job_dir: Path) -> dict[str, float]:
    path = job_dir / "alturas.txt"
    if not path.exists():
        return {}
    overrides = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        if "\t" not in line:
            continue
        level, value = line.split("\t", 1)
        try:
            overrides[level] = float(value)
        except ValueError:
            pass
    return overrides


def _save_height_overrides(job_dir: Path, overrides: dict[str, float]) -> None:
    if not overrides:
        return
    combined = {**_read_height_overrides(job_dir), **overrides}
    lines = [f"{level}\t{height}" for level, height in combined.items()]
    (job_dir / "alturas.txt").write_text("\n".join(lines), encoding="utf-8")


@app.route("/processar", methods=["POST"])
def processar():
    uploaded = request.files.get("pdf")
    if not uploaded or not uploaded.filename:
        return _render(result=None, questions=None, orcamento=None), 400

    job_id = uuid.uuid4().hex[:12]
    job_dir = _JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    pdf_path = job_dir / "planta.pdf"
    uploaded.save(pdf_path)
    (job_dir / "nome.txt").write_text(uploaded.filename, encoding="utf-8")
    _append_hints(job_dir, request.form.get("hints", ""))

    return _finish_analysis(job_id, str(pdf_path), height_overrides=None)


@app.route("/completar", methods=["POST"])
def completar():
    job_id = request.form.get("job_id", "")
    job_dir = _JOBS_DIR / job_id
    pdf_path = job_dir / "planta.pdf"
    if not pdf_path.exists():
        return "Sessão expirou — reenvie o PDF.", 404

    overrides = {}
    for key, value in request.form.items():
        if key.startswith("altura_"):
            level = key[len("altura_"):]
            try:
                overrides[level] = float(value.replace(",", "."))
            except ValueError:
                pass
    _save_height_overrides(job_dir, overrides)

    return _finish_analysis(job_id, str(pdf_path), height_overrides=overrides)


@app.route("/reprocessar", methods=["POST"])
def reprocessar():
    job_id = request.form.get("job_id", "")
    job_dir = _JOBS_DIR / job_id
    pdf_path = job_dir / "planta.pdf"
    if not pdf_path.exists():
        return "Sessão expirou — reenvie o PDF.", 404

    _append_hints(job_dir, request.form.get("hints", ""))
    return _finish_analysis(job_id, str(pdf_path), height_overrides=None)


def _finish_analysis(job_id: str, pdf_path: str, height_overrides: dict[str, float] | None):
    job_dir = _JOBS_DIR / job_id
    extra_text = _read_hints(job_dir)
    combined_overrides = {**_read_height_overrides(job_dir), **(height_overrides or {})}
    result = analyze_pdf(pdf_path, height_overrides=combined_overrides, extra_text=extra_text)

    if result.levels_missing_height:
        pdf_name = (job_dir / "nome.txt").read_text(encoding="utf-8") if (job_dir / "nome.txt").exists() else "o PDF"
        return _render(
            result=None,
            orcamento=None,
            questions={
                "job_id": job_id,
                "levels": sorted(result.levels_missing_height),
                "pdf_name": pdf_name,
            },
        )

    entrada_path = job_dir / "entrada.xlsx"
    conferencia_path = job_dir / "conferencia.pdf"
    write_prefilled_input(result, str(entrada_path))
    render_review_pdf(pdf_path, result, str(conferencia_path))

    if result.method == "room-perimeter":
        n_wall = sum(1 for r in result.rooms if r.perimeter_m > 0)
        n_ceiling = sum(1 for r in result.rooms if r.ceiling_area_m2 > 0)
        wall_label = "ambiente(s) com parede detectada"
        ceiling_label = "com teto detectado"
    else:
        n_wall = len(result.wall_rows)
        n_ceiling = 0
        wall_label = "linha(s) de parede detectada(s) (método: " + result.method + ")"
        ceiling_label = "com teto detectado — teto não é detectado por este método, adicione manualmente"

    return _render(
        questions=None,
        orcamento=None,
        result={
            "job_id": job_id,
            "n_wall": n_wall,
            "n_ceiling": n_ceiling,
            "wall_label": wall_label,
            "ceiling_label": ceiling_label,
            "warnings": result.warnings,
        },
    )


@app.route("/construir", methods=["POST"])
def construir():
    uploaded = request.files.get("entrada")
    if not uploaded or not uploaded.filename:
        return _render(result=None, questions=None, orcamento=None), 400

    job_id = uuid.uuid4().hex[:12]
    job_dir = _JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    entrada_path = job_dir / "entrada_ajustada.xlsx"
    uploaded.save(entrada_path)

    orcamento_path = job_dir / "orcamento.xlsx"
    build_from_input(str(entrada_path), str(orcamento_path))

    return _render(result=None, questions=None, orcamento={"job_id": job_id})


@app.route("/baixar/<job_id>/<filename>", methods=["GET"])
def baixar(job_id: str, filename: str):
    if filename not in ("entrada.xlsx", "conferencia.pdf", "orcamento.xlsx"):
        return "Arquivo não encontrado.", 404
    path = _JOBS_DIR / job_id / filename
    if not path.exists():
        return "Arquivo não encontrado (ou a sessão expirou — reenvie o arquivo).", 404
    return send_file(path, as_attachment=True)


def main() -> int:
    url = "http://127.0.0.1:5000"
    print(f"Abra {url} no navegador para enviar a planta.")
    print("(O navegador deve abrir sozinho em alguns segundos. Deixe esta janela aberta.)")
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
