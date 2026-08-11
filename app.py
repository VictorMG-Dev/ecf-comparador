"""
app.py — Servidor Flask para a interface web do ECF Comparador.

Rotas:
  GET  /                  → Página principal (UI)
  POST /processar         → Recebe uploads + labels, inicia processamento
  GET  /stream/<job_id>   → SSE com log em tempo real
  GET  /download/<job_id> → Download do Excel gerado
  GET  /status/<job_id>   → JSON com status do job
"""

from __future__ import annotations

import io
import json
import logging
import os
import queue
import sys
import tempfile
import threading
import time
import uuid
from pathlib import Path
from typing import Any, Dict

from flask import (
    Flask,
    Response,
    jsonify,
    render_template,
    request,
    send_file,
)

# ---------------------------------------------------------------------------
# Bootstrap: garante que os módulos do projeto estejam no path
# ---------------------------------------------------------------------------
sys.path.insert(0, str(Path(__file__).parent))

from comparator import ResultadoBloco, comparar_bloco
from normalizer import normalizar_df
from reader import ler_arquivo_ecf
from writer import gerar_excel

# ---------------------------------------------------------------------------
# App Flask
# ---------------------------------------------------------------------------
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB por arquivo

# Diretório temporário para arquivos de upload e saída
TEMP_DIR = Path(tempfile.gettempdir()) / "ecf_comparador"
TEMP_DIR.mkdir(exist_ok=True)

# Registro de jobs em memória: { job_id: { status, result_path, queue, ... } }
JOBS: Dict[str, Dict[str, Any]] = {}
JOBS_LOCK = threading.Lock()

# ---------------------------------------------------------------------------
# Logging customizado que alimenta a fila SSE por job
# ---------------------------------------------------------------------------

class QueueHandler(logging.Handler):
    """Handler que envia registros de log para uma fila (SSE)."""

    def __init__(self, q: queue.Queue):
        super().__init__()
        self.q = q

    def emit(self, record: logging.LogRecord) -> None:
        try:
            nivel = record.levelname
            modulo = record.name
            msg = self.format(record)
            # Formata para o frontend
            self.q.put({
                "type": "log",
                "level": nivel.lower(),
                "module": modulo,
                "message": record.getMessage(),
                "full": msg,
            })
        except Exception:  # noqa: BLE001
            pass


# ---------------------------------------------------------------------------
# Rotas
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/processar", methods=["POST"])
def processar():
    """Recebe os dois arquivos e os labels, inicia processamento em thread."""
    arquivo_ant = request.files.get("arquivo_anterior")
    arquivo_atu = request.files.get("arquivo_atual")
    label_ant = request.form.get("label_anterior", "2024").strip() or "2024"
    label_atu = request.form.get("label_atual", "2025").strip() or "2025"
    nome_saida = request.form.get("nome_saida", "").strip() or "ecf_revisao"

    if not arquivo_ant or not arquivo_atu:
        return jsonify({"erro": "Envie os dois arquivos Excel."}), 400

    # Valida extensões
    for arq in (arquivo_ant, arquivo_atu):
        ext = Path(arq.filename or "").suffix.lower()
        if ext not in (".xlsx", ".xls", ".xlsm", ".xlsb", ".txt"):
            return jsonify({
                "erro": f"Arquivo '{arq.filename}' não é suportado (.xlsx/.xls/.xlsm/.xlsb/.txt)."
            }), 400

    # Salva uploads em disco temporário
    job_id = uuid.uuid4().hex
    job_dir = TEMP_DIR / job_id
    job_dir.mkdir()

    ext_ant = Path(arquivo_ant.filename).suffix.lower()
    ext_atu = Path(arquivo_atu.filename).suffix.lower()
    path_ant = job_dir / f"anterior{ext_ant}"
    path_atu = job_dir / f"atual{ext_atu}"
    path_out = job_dir / f"{nome_saida}.xlsx"

    arquivo_ant.save(str(path_ant))
    arquivo_atu.save(str(path_atu))

    # Registra job
    q: queue.Queue = queue.Queue()
    with JOBS_LOCK:
        JOBS[job_id] = {
            "status": "rodando",
            "progresso": 0,
            "q": q,
            "result_path": None,
            "nome_saida": f"{nome_saida}.xlsx",
            "resumo": None,
            "erro": None,
            "criado_em": time.time(),
        }

    # Inicia processamento em thread separada
    t = threading.Thread(
        target=_processar_job,
        args=(job_id, path_ant, path_atu, path_out, label_ant, label_atu, q),
        daemon=True,
    )
    t.start()

    return jsonify({"job_id": job_id})


@app.route("/stream/<job_id>")
def stream(job_id: str):
    """SSE: transmite logs e eventos de progresso em tempo real."""
    with JOBS_LOCK:
        job = JOBS.get(job_id)

    if not job:
        def erro_gen():
            yield f"data: {json.dumps({'type': 'erro', 'message': 'Job não encontrado.'})}\n\n"
        return Response(erro_gen(), mimetype="text/event-stream")

    def gerador():
        q = job["q"]
        while True:
            try:
                evento = q.get(timeout=30)
                yield f"data: {json.dumps(evento, ensure_ascii=False)}\n\n"
                if evento.get("type") in ("concluido", "erro"):
                    break
            except queue.Empty:
                # Heartbeat para manter conexão SSE viva
                yield f"data: {json.dumps({'type': 'heartbeat'})}\n\n"

    return Response(
        gerador(),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


@app.route("/status/<job_id>")
def status(job_id: str):
    """Retorna JSON com status atual do job."""
    with JOBS_LOCK:
        job = JOBS.get(job_id)

    if not job:
        return jsonify({"erro": "Job não encontrado."}), 404

    return jsonify({
        "status": job["status"],
        "progresso": job["progresso"],
        "resumo": job["resumo"],
        "erro": job["erro"],
    })


@app.route("/download/<job_id>")
def download(job_id: str):
    """Retorna o arquivo Excel gerado para download."""
    with JOBS_LOCK:
        job = JOBS.get(job_id)

    if not job or not job.get("result_path"):
        return jsonify({"erro": "Resultado não disponível."}), 404

    path = Path(job["result_path"])
    if not path.exists():
        return jsonify({"erro": "Arquivo não encontrado no servidor."}), 404

    return send_file(
        str(path),
        as_attachment=True,
        download_name=job["nome_saida"],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Lógica de processamento (roda em thread)
# ---------------------------------------------------------------------------

def _processar_job(
    job_id: str,
    path_ant: Path,
    path_atu: Path,
    path_out: Path,
    label_ant: str,
    label_atu: str,
    q: queue.Queue,
) -> None:
    """Executa toda a pipeline de comparação e publica eventos na fila."""

    def _pub(tipo: str, **kwargs):
        q.put({"type": tipo, **kwargs})

    def _log(nivel: str, modulo: str, msg: str):
        q.put({"type": "log", "level": nivel, "module": modulo, "message": msg})

    try:
        # Configura logger para este job
        root_logger = logging.getLogger()
        qh = QueueHandler(q)
        qh.setLevel(logging.DEBUG)
        qh.setFormatter(logging.Formatter("%(message)s"))
        root_logger.addHandler(qh)
        root_logger.setLevel(logging.DEBUG)

        _pub("progresso", valor=5, texto="Lendo arquivo anterior...")

        # --- Leitura ---
        abas_ant = ler_arquivo_ecf(path_ant)
        _pub("progresso", valor=20, texto=f"Arquivo anterior lido: {len(abas_ant)} abas/blocos")

        abas_atu = ler_arquivo_ecf(path_atu)
        _pub("progresso", valor=35, texto=f"Arquivo atual lido: {len(abas_atu)} abas/blocos")

        # --- Processamento por bloco ---
        todos_blocos = sorted(set(abas_ant.keys()) | set(abas_atu.keys()))
        total = len(todos_blocos)
        resultados = []

        for i, bloco in enumerate(todos_blocos):
            pct = int(35 + ((i + 1) / total) * 45)
            _pub("progresso", valor=pct, texto=f"Processando bloco {bloco} [{i+1}/{total}]")
            _pub("bloco", nome=bloco, idx=i+1, total=total)

            df_a = abas_ant.get(bloco)
            df_u = abas_atu.get(bloco)

            if df_a is not None and not df_a.empty:
                df_a = normalizar_df(df_a)
            if df_u is not None and not df_u.empty:
                df_u = normalizar_df(df_u)

            resultado = comparar_bloco(bloco, df_a, df_u)
            resultados.append(resultado)

        _pub("progresso", valor=85, texto="Gerando arquivo Excel de saída...")

        # --- Escrita ---
        gerar_excel(
            resultados=resultados,
            caminho_saida=str(path_out),
            label_anterior=label_ant,
            label_atual=label_atu,
        )

        _pub("progresso", valor=98, texto="Finalizando...")

        # --- Resumo para o frontend ---
        resumo = []
        totais = {"incluidos": 0, "excluidos": 0, "alterados": 0, "sem_alteracao": 0}
        for r in resultados:
            aviso = r.aviso
            tag = ""
            if "NOVO" in aviso.upper():
                tag = "novo"
            elif "REMOVIDO" in aviso.upper():
                tag = "removido"
            elif r.chave_automatica:
                tag = "auto"

            resumo.append({
                "bloco": r.nome_bloco,
                "incluidos": r.qtd_incluidos,
                "excluidos": r.qtd_excluidos,
                "alterados": r.qtd_alterados,
                "sem_alteracao": r.qtd_sem_alteracao,
                "tag": tag,
                "aviso": aviso,
                "erros": r.erros,
            })
            totais["incluidos"] += r.qtd_incluidos
            totais["excluidos"] += r.qtd_excluidos
            totais["alterados"] += r.qtd_alterados
            totais["sem_alteracao"] += r.qtd_sem_alteracao

        with JOBS_LOCK:
            JOBS[job_id]["status"] = "concluido"
            JOBS[job_id]["progresso"] = 100
            JOBS[job_id]["result_path"] = str(path_out)
            JOBS[job_id]["resumo"] = {"blocos": resumo, "totais": totais}

        _pub("progresso", valor=100, texto="Concluido!")
        _pub("concluido", resumo={"blocos": resumo, "totais": totais}, job_id=job_id)

    except Exception as exc:  # noqa: BLE001
        import traceback
        tb = traceback.format_exc()
        with JOBS_LOCK:
            JOBS[job_id]["status"] = "erro"
            JOBS[job_id]["erro"] = str(exc)
        _pub("erro", message=str(exc), traceback=tb)

    finally:
        # Remove o QueueHandler para não vazar entre jobs
        root_logger = logging.getLogger()
        root_logger.handlers = [
            h for h in root_logger.handlers if not isinstance(h, QueueHandler)
        ]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print("=" * 60)
    print("  ECF Comparador — Interface Web")
    print("  Acesse: http://localhost:5000")
    print("=" * 60)
    app.run(debug=False, host="0.0.0.0", port=5000, threaded=True)
