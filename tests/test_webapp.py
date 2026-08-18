import io

from pdf_takeoff.demo import generate_sample_plant
from pdf_takeoff.webapp import app


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
