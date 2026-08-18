"""Página local no navegador pra facilitar o upload do PDF: arrasta o arquivo,
clica em processar, baixa a planilha pré-preenchida e o PDF de conferência —
sem precisar digitar comando nenhum depois de abrir a página.

Rodar com: python -m pdf_takeoff.webapp
Depois abra http://localhost:5000 no navegador.
"""

from __future__ import annotations

import tempfile
import uuid
from pathlib import Path

from flask import Flask, render_template_string, request, send_file

from .pdf_to_input import analyze_pdf, render_review_pdf, write_prefilled_input

app = Flask(__name__)

_JOBS_DIR = Path(tempfile.gettempdir()) / "pdf_takeoff_webapp"
_JOBS_DIR.mkdir(exist_ok=True)

_PAGE = """
<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<title>Quantificador de Gesso/Pintura - Upload de Planta</title>
<style>
  body { font-family: system-ui, sans-serif; max-width: 720px; margin: 40px auto; padding: 0 20px; color: #222; }
  h1 { font-size: 1.4rem; }
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
  .result { background: #f0fff4; border: 1px solid #38a169; border-radius: 6px; padding: 16px; margin: 20px 0; }
  a.download { display: inline-block; margin: 6px 12px 6px 0; }
  #filename { font-weight: bold; }
  #status { margin-top: 10px; color: #555; }
</style>
</head>
<body>
<h1>Quantificador automático de Gesso/Pintura</h1>
<p>Envie o PDF da planta (formato de plantas com convenção GIB) e clique em Processar.</p>

<form id="form" method="post" enctype="multipart/form-data" action="/processar">
  <div class="drop" id="drop">
    <p>Arraste o PDF aqui, ou clique pra escolher</p>
    <p id="filename"></p>
    <input type="file" id="fileInput" name="pdf" accept="application/pdf" required>
  </div>
  <button type="submit" id="submitBtn" disabled>Processar</button>
  <div id="status"></div>
</form>

{% if result %}
<div class="result">
  <p><strong>Pronto!</strong> {{ result.n_wall }} ambiente(s) com parede detectada, {{ result.n_ceiling }} com teto detectado.</p>
  {% if result.warnings %}
  <div class="warn">
    <strong>Avisos:</strong>
    <ul>{% for w in result.warnings %}<li>{{ w }}</li>{% endfor %}</ul>
  </div>
  {% endif %}
  <a class="download" href="/baixar/{{ result.job_id }}/entrada.xlsx">⬇ Baixar planilha de entrada (.xlsx)</a>
  <a class="download" href="/baixar/{{ result.job_id }}/conferencia.pdf">⬇ Baixar PDF de conferência</a>
</div>
{% endif %}

<script>
  const drop = document.getElementById('drop');
  const input = document.getElementById('fileInput');
  const filenameEl = document.getElementById('filename');
  const submitBtn = document.getElementById('submitBtn');
  const statusEl = document.getElementById('status');
  const form = document.getElementById('form');

  drop.addEventListener('click', () => input.click());
  drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('dragover'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
  drop.addEventListener('drop', (e) => {
    e.preventDefault();
    drop.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
      input.files = e.dataTransfer.files;
      updateFilename();
    }
  });
  input.addEventListener('change', updateFilename);

  function updateFilename() {
    if (input.files.length) {
      filenameEl.textContent = input.files[0].name;
      submitBtn.disabled = false;
    }
  }

  form.addEventListener('submit', () => {
    submitBtn.disabled = true;
    submitBtn.textContent = 'Processando...';
    statusEl.textContent = 'Isso pode levar alguns minutos numa planta grande, aguarde...';
  });
</script>
</body>
</html>
"""


@app.route("/", methods=["GET"])
def index():
    return render_template_string(_PAGE, result=None)


@app.route("/processar", methods=["POST"])
def processar():
    uploaded = request.files.get("pdf")
    if not uploaded or not uploaded.filename:
        return render_template_string(_PAGE, result=None), 400

    job_id = uuid.uuid4().hex[:12]
    job_dir = _JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    pdf_path = job_dir / "planta.pdf"
    uploaded.save(pdf_path)

    entrada_path = job_dir / "entrada.xlsx"
    conferencia_path = job_dir / "conferencia.pdf"

    result = analyze_pdf(str(pdf_path))
    write_prefilled_input(result, str(entrada_path))
    render_review_pdf(str(pdf_path), result, str(conferencia_path))

    n_wall = sum(1 for r in result.rooms if r.perimeter_m > 0)
    n_ceiling = sum(1 for r in result.rooms if r.ceiling_area_m2 > 0)

    return render_template_string(
        _PAGE,
        result={
            "job_id": job_id,
            "n_wall": n_wall,
            "n_ceiling": n_ceiling,
            "warnings": result.warnings,
        },
    )


@app.route("/baixar/<job_id>/<filename>", methods=["GET"])
def baixar(job_id: str, filename: str):
    if filename not in ("entrada.xlsx", "conferencia.pdf"):
        return "Arquivo não encontrado.", 404
    path = _JOBS_DIR / job_id / filename
    if not path.exists():
        return "Arquivo não encontrado (ou a sessão expirou — reenvie o PDF).", 404
    return send_file(path, as_attachment=True)


def main() -> int:
    print("Abra http://localhost:5000 no navegador para enviar a planta.")
    app.run(host="127.0.0.1", port=5000, debug=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
