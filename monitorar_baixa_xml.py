from __future__ import annotations

import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path


INTERVALO_MINUTOS = 2
RODAR_IMEDIATAMENTE = True
CAMINHO_SCRIPT_ALVO = Path(__file__).with_name("baixar_xml_nfe.py")
CAMINHO_LOG = Path(__file__).with_name("monitorar_baixa_xml.log")


def agora_texto() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def registrar(mensagem: str) -> None:
    linha = f"[{agora_texto()}] {mensagem}"
    print(linha, flush=True)
    CAMINHO_LOG.parent.mkdir(parents=True, exist_ok=True)
    with CAMINHO_LOG.open("a", encoding="utf-8") as arquivo:
        arquivo.write(linha + "\n")


def executar_script_alvo() -> int:
    if not CAMINHO_SCRIPT_ALVO.exists():
        registrar(f"Arquivo nao encontrado: {CAMINHO_SCRIPT_ALVO}")
        return 1

    registrar(f"Executando {CAMINHO_SCRIPT_ALVO.name}...")
    processo = subprocess.run(
        [sys.executable, str(CAMINHO_SCRIPT_ALVO)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=str(CAMINHO_SCRIPT_ALVO.parent),
    )

    stdout = (processo.stdout or "").strip()
    stderr = (processo.stderr or "").strip()

    if stdout:
        for linha in stdout.splitlines():
            registrar(f"saida: {linha}")

    if stderr:
        for linha in stderr.splitlines():
            registrar(f"erro: {linha}")

    registrar(f"Retorno do processo: {processo.returncode}")
    return processo.returncode


def aguardar_proxima_execucao(segundos: int) -> None:
    restante = segundos
    while restante > 0:
        pausa = min(restante, 30)
        time.sleep(pausa)
        restante -= pausa


def main() -> int:
    intervalo_segundos = INTERVALO_MINUTOS * 60
    registrar(
        f"Monitor iniciado. Intervalo: {INTERVALO_MINUTOS} minuto(s). "
        "Pressione Ctrl+C para parar."
    )

    try:
        primeira_execucao = True
        while True:
            if RODAR_IMEDIATAMENTE or not primeira_execucao:
                executar_script_alvo()
            else:
                registrar("Primeira execucao pulada pela configuracao.")

            primeira_execucao = False
            proxima = datetime.now().timestamp() + intervalo_segundos
            registrar(
                "Proxima tentativa em "
                f"{INTERVALO_MINUTOS} minuto(s), por volta de "
                f"{datetime.fromtimestamp(proxima).strftime('%d/%m/%Y %H:%M:%S')}."
            )
            aguardar_proxima_execucao(intervalo_segundos)
    except KeyboardInterrupt:
        registrar("Monitor encerrado pelo usuario.")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
