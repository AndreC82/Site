import io

from reportlab.pdfgen import canvas

from pdf_takeoff.demo import generate_sample_plant
from pdf_takeoff.webapp import app


def _generate_gib_plant_without_stud_height(path: str) -> None:
    """Planta sintética no formato GIB (com Scale/Wall Linings/Ceiling Linings),
    mas sem a nota de altura de pé-direito — deve disparar a pergunta na página.
    """
    c = canvas.Canvas(path, pagesize=(1684, 1191))
    c.drawString(100, 1100, "Lower Floor Plan")
    c.drawString(100, 1080, "1:100 @ A2")
    c.drawString(100, 900, "Wall Linings")
    c.drawString(100, 885, "10mm Gibboard internal linings (Aqualine to wet areas)")
    c.rect(100, 100, 400, 300, stroke=1, fill=0)
    c.drawCentredString(300, 250, "Sala Teste")
    c.showPage()
    c.save()


def test_index_page_loads():
    client = app.test_client()
    resp = client.get("/")
    assert resp.status_code == 200
    assert b"Arraste o PDF aqui" in resp.data


def test_upload_and_process_pdf(tmp_path):
    pdf_path = tmp_path / "planta.pdf"
    generate_sample_plant(str(pdf_path))

    client = app.test_client()
    with open(pdf_path, "rb") as f:
        data = {"pdf": (io.BytesIO(f.read()), "planta.pdf")}
        resp = client.post("/processar", data=data, content_type="multipart/form-data")

    assert resp.status_code == 200
    assert b"ambiente(s) com parede detectada" in resp.data
    assert b"/baixar/" in resp.data


def test_processar_without_file_returns_400():
    client = app.test_client()
    resp = client.post("/processar", data={}, content_type="multipart/form-data")
    assert resp.status_code == 400


def test_baixar_unknown_job_returns_404():
    client = app.test_client()
    resp = client.get("/baixar/naoexiste123/entrada.xlsx")
    assert resp.status_code == 404


def test_baixar_rejects_arbitrary_filename():
    client = app.test_client()
    resp = client.get("/baixar/algumjob/../../etc/passwd")
    assert resp.status_code == 404


def test_missing_stud_height_triggers_question_and_completar_finishes(tmp_path):
    pdf_path = tmp_path / "planta.pdf"
    _generate_gib_plant_without_stud_height(str(pdf_path))

    client = app.test_client()
    with open(pdf_path, "rb") as f:
        data = {"pdf": (io.BytesIO(f.read()), "planta.pdf")}
        resp = client.post("/processar", data=data, content_type="multipart/form-data")

    assert resp.status_code == 200
    assert "altura de p\xe9-direito".encode() in resp.data.lower() or b"altura_" in resp.data
    assert b'name="job_id" value="' in resp.data

    job_id = resp.data.split(b'name="job_id" value="')[1].split(b'"')[0].decode()

    resp2 = client.post("/completar", data={"job_id": job_id, "altura_Lower": "2.70"})
    assert resp2.status_code == 200
    assert b"/baixar/" in resp2.data
    assert b"ambiente(s) com parede detectada" in resp2.data


def test_completar_with_unknown_job_returns_404():
    client = app.test_client()
    resp = client.post("/completar", data={"job_id": "naoexiste", "altura_Lower": "2.7"})
    assert resp.status_code == 404


def test_construir_end_to_end_from_downloaded_entrada(tmp_path):
    pdf_path = tmp_path / "planta.pdf"
    _generate_gib_plant_without_stud_height(str(pdf_path))

    client = app.test_client()
    with open(pdf_path, "rb") as f:
        resp = client.post(
            "/processar",
            data={"pdf": (io.BytesIO(f.read()), "planta.pdf")},
            content_type="multipart/form-data",
        )
    job_id = resp.data.split(b'name="job_id" value="')[1].split(b'"')[0].decode()

    resp2 = client.post("/completar", data={"job_id": job_id, "altura_Lower": "2.70"})
    entrada_job_id = resp2.data.split(b'/baixar/')[1].split(b'/')[0].decode()

    entrada_resp = client.get(f"/baixar/{entrada_job_id}/entrada.xlsx")
    assert entrada_resp.status_code == 200

    resp3 = client.post(
        "/construir",
        data={"entrada": (io.BytesIO(entrada_resp.data), "entrada.xlsx")},
        content_type="multipart/form-data",
    )
    assert resp3.status_code == 200
    assert b"Or\xc3\xa7amento gerado" in resp3.data
    assert b"orcamento.xlsx" in resp3.data
