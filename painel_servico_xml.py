from __future__ import annotations

import ast
import ctypes
import json
import os
import re
import signal
import subprocess
import sys
from pathlib import Path

import flet as ft


BASE_DIR = Path(__file__).resolve().parent
ARQUIVO_BAIXA = BASE_DIR / "baixar_xml_nfe.py"
ARQUIVO_MONITOR = BASE_DIR / "monitorar_baixa_xml.py"
ARQUIVO_PID = BASE_DIR / "monitorar_baixa_xml.pid"


def ler_configuracao_atual() -> dict[str, str]:
    conteudo = ARQUIVO_BAIXA.read_text(encoding="utf-8")
    chaves = {
        "CNPJ_AUTOR": "",
        "CAMINHO_CERTIFICADO_PFX": "",
        "SENHA_CERTIFICADO": "",
        "CAMINHO_SALVAR_XML": "",
    }

    for chave in chaves:
        padrao = rf"^{chave}\s*=\s*(.+)$"
        encontrado = re.search(padrao, conteudo, flags=re.MULTILINE)
        if encontrado:
            valor_bruto = encontrado.group(1).strip()
            try:
                chaves[chave] = str(ast.literal_eval(valor_bruto))
            except Exception:
                chaves[chave] = valor_bruto.strip('"').strip("'")

    return chaves


def substituir_linha(conteudo: str, variavel: str, valor: str) -> str:
    linha_nova = f"{variavel} = {json.dumps(valor, ensure_ascii=False)}"
    padrao = rf"^{variavel}\s*=.*$"
    return re.sub(padrao, linha_nova, conteudo, flags=re.MULTILINE)


def salvar_configuracao(
    cnpj: str,
    caminho_certificado: str,
    senha: str,
    caminho_saida: str,
) -> None:
    conteudo = ARQUIVO_BAIXA.read_text(encoding="utf-8")
    conteudo = substituir_linha(conteudo, "CNPJ_AUTOR", cnpj)
    conteudo = substituir_linha(conteudo, "CAMINHO_CERTIFICADO_PFX", caminho_certificado)
    conteudo = substituir_linha(conteudo, "SENHA_CERTIFICADO", senha)
    conteudo = substituir_linha(conteudo, "CAMINHO_SALVAR_XML", caminho_saida)
    ARQUIVO_BAIXA.write_text(conteudo, encoding="utf-8")


def ler_pid() -> int | None:
    if not ARQUIVO_PID.exists():
        return None

    try:
        return int(ARQUIVO_PID.read_text(encoding="utf-8").strip())
    except (ValueError, OSError):
        return None


def processo_ativo(pid: int | None) -> bool:
    if not pid:
        return False

    try:
        os.kill(pid, 0)
    except OSError:
        return False
    return True


def iniciar_monitor() -> str:
    pid_atual = ler_pid()
    if processo_ativo(pid_atual):
        return f"Servico ja esta rodando. PID: {pid_atual}"

    kwargs: dict[str, object] = {}
    if os.name == "nt":
        kwargs["creationflags"] = (
            subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS
        )

    processo = subprocess.Popen(
        [sys.executable, str(ARQUIVO_MONITOR)],
        cwd=str(BASE_DIR),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        stdin=subprocess.DEVNULL,
        **kwargs,
    )
    ARQUIVO_PID.write_text(str(processo.pid), encoding="utf-8")
    return f"Servico iniciado com sucesso. PID: {processo.pid}"


def parar_monitor() -> str:
    pid = ler_pid()
    if not processo_ativo(pid):
        try:
            ARQUIVO_PID.unlink(missing_ok=True)
        except OSError:
            pass
        return "Servico nao esta em execucao."

    if os.name == "nt":
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            capture_output=True,
            text=True,
            check=False,
        )
    else:
        os.kill(pid, signal.SIGTERM)

    try:
        ARQUIVO_PID.unlink(missing_ok=True)
    except OSError:
        pass
    return f"Servico parado. PID finalizado: {pid}"


def status_servico() -> str:
    pid = ler_pid()
    if processo_ativo(pid):
        return f"Status do servico: rodando (PID {pid})"
    return "Status do servico: parado"


def recarregar_aplicacao() -> None:
    subprocess.Popen(
        [sys.executable, str(Path(__file__).resolve())],
        cwd=str(BASE_DIR),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        stdin=subprocess.DEVNULL,
    )


def obter_area_util_monitor_principal() -> tuple[int, int, int, int] | None:
    if os.name != "nt":
        return None

    class RECT(ctypes.Structure):
        _fields_ = [
            ("left", ctypes.c_long),
            ("top", ctypes.c_long),
            ("right", ctypes.c_long),
            ("bottom", ctypes.c_long),
        ]

    area_util = RECT()
    spi_getworkarea = 48
    sucesso = ctypes.windll.user32.SystemParametersInfoW(
        spi_getworkarea,
        0,
        ctypes.byref(area_util),
        0,
    )
    if not sucesso:
        return None

    return area_util.left, area_util.top, area_util.right, area_util.bottom


def configurar_janela(page: ft.Page) -> int:
    area_util = obter_area_util_monitor_principal()
    if area_util is None:
        largura_janela = 470
        altura_janela = 760
    else:
        esquerda, topo, direita, baixo = area_util
        largura_util = direita - esquerda
        altura_util = baixo - topo
        largura_janela = min(max(int(largura_util * 0.34), 430), 560)
        altura_janela = min(max(int(altura_util * 0.78), 720), 900)
        page.window.left = esquerda + max(0, (largura_util - largura_janela) // 2)
        page.window.top = topo + max(0, (altura_util - altura_janela) // 2)

    page.title = "Servico de XML NFe"
    page.padding = 0
    page.spacing = 0
    page.scroll = ft.ScrollMode.AUTO
    page.theme_mode = ft.ThemeMode.LIGHT
    page.bgcolor = "#b08cff"
    page.window.width = largura_janela
    page.window.height = altura_janela
    page.window.resizable = True
    page.window.min_width = 410
    page.window.min_height = 700
    return largura_janela


def criar_input(
    label: str,
    valor: str,
    icone: ft.Icons,
    largura: int,
    password: bool = False,
    autofocus: bool = False,
) -> ft.TextField:
    return ft.TextField(
        label=label,
        value=valor,
        width=largura,
        autofocus=autofocus,
        password=password,
        can_reveal_password=password,
        prefix_icon=icone,
        filled=True,
        bgcolor="#f1e5ff",
        color="#27153f",
        border=ft.InputBorder.OUTLINE,
        border_radius=20,
        border_color="#caa8ff",
        focused_border_color="#7c4dff",
        focused_bgcolor="#ffffff",
        content_padding=18,
        label_style=ft.TextStyle(color="#6f4aa6", size=12, weight=ft.FontWeight.W_600),
    )


def criar_botao(
    texto: str,
    icone: ft.Icons,
    cor_fundo: str,
    cor_texto: str,
    largura: int,
    acao,
) -> ft.Button:
    return ft.Button(
        texto,
        icon=icone,
        on_click=acao,
        width=largura,
        height=54,
        bgcolor=cor_fundo,
        color=cor_texto,
        elevation=6,
        style=ft.ButtonStyle(
            shape=ft.RoundedRectangleBorder(radius=18),
            padding=ft.Padding.symmetric(horizontal=18, vertical=14),
        ),
    )


async def main(page: ft.Page) -> None:
    largura_janela = configurar_janela(page)
    largura_bloco = min(420, max(330, largura_janela - 86))
    config = ler_configuracao_atual()

    campo_cnpj = criar_input(
        "CNPJ CLIENTE",
        config["CNPJ_AUTOR"],
        ft.Icons.BADGE_OUTLINED,
        largura_bloco,
        autofocus=True,
    )
    campo_certificado = criar_input(
        "CAMINHO CERTIFICADO",
        config["CAMINHO_CERTIFICADO_PFX"],
        ft.Icons.FOLDER_OPEN_OUTLINED,
        largura_bloco,
    )
    campo_senha = criar_input(
        "SENHA CERTIFICADO",
        config["SENHA_CERTIFICADO"],
        ft.Icons.LOCK_OUTLINED,
        largura_bloco,
        password=True,
    )
    campo_saida = criar_input(
        "CAMINHO PARA SALVAR INFORMACOES",
        config["CAMINHO_SALVAR_XML"],
        ft.Icons.SAVE_AS_OUTLINED,
        largura_bloco,
    )

    status = ft.Text(
        value=status_servico(),
        color="#f4ecff",
        size=13,
        weight=ft.FontWeight.W_600,
    )
    mensagem = ft.Text(
        value="",
        color="#d9c7ff",
        size=12,
        selectable=True,
    )
    mensagem_box = ft.Container(
        width=largura_bloco,
        padding=ft.Padding.symmetric(horizontal=16, vertical=14),
        border_radius=18,
        bgcolor="#33204f",
        content=mensagem,
    )

    def atualizar_status(texto: str, cor: str, fundo: str) -> None:
        mensagem.value = texto
        mensagem.color = cor
        mensagem_box.bgcolor = fundo
        status.value = status_servico()
        page.update()

    def ao_salvar(_: ft.ControlEvent) -> None:
        try:
            salvar_configuracao(
                campo_cnpj.value or "",
                campo_certificado.value or "",
                campo_senha.value or "",
                campo_saida.value or "",
            )
            atualizar_status(
                "Configuracoes salvas no baixar_xml_nfe.py.",
                "#dcffe9",
                "#214635",
            )
        except Exception as exc:
            atualizar_status(
                f"Erro ao salvar configuracoes: {exc}",
                "#ffe1e8",
                "#5a2131",
            )

    def ao_rodar(_: ft.ControlEvent) -> None:
        try:
            texto = iniciar_monitor()
            atualizar_status(texto, "#e8fff1", "#214635")
        except Exception as exc:
            atualizar_status(
                f"Erro ao iniciar servico: {exc}",
                "#ffe1e8",
                "#5a2131",
            )

    def ao_parar(_: ft.ControlEvent) -> None:
        try:
            texto = parar_monitor()
            atualizar_status(texto, "#fff4dd", "#5f431a")
        except Exception as exc:
            atualizar_status(
                f"Erro ao parar servico: {exc}",
                "#ffe1e8",
                "#5a2131",
            )

    def ao_recarregar(_: ft.ControlEvent) -> None:
        try:
            recarregar_aplicacao()
            page.window.close()
        except Exception as exc:
            atualizar_status(
                f"Erro ao recarregar layout: {exc}",
                "#ffe1e8",
                "#5a2131",
            )

    topo = ft.Row(
        width=largura_bloco,
        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
        vertical_alignment=ft.CrossAxisAlignment.START,
        controls=[
            ft.Column(
                spacing=4,
                controls=[
                    ft.Text(
                        "Painel de XML",
                        size=26,
                        weight=ft.FontWeight.BOLD,
                        color="#ffffff",
                    ),
                    ft.Text(
                        "Controle do servico e configuracoes do download.",
                        size=12,
                        color="#e3d7ff",
                    ),
                ],
            ),
            ft.Container(
                padding=ft.Padding.symmetric(horizontal=14, vertical=8),
                border_radius=18,
                bgcolor="#f3ecff22",
                border=ft.Border.all(1, "#ffffff22"),
                content=ft.Icon(ft.Icons.CLOUD_SYNC_OUTLINED, color="#ffffff"),
            ),
        ],
    )

    status_chip = ft.Container(
        padding=ft.Padding.symmetric(horizontal=16, vertical=10),
        border_radius=18,
        bgcolor="#ffffff18",
        border=ft.Border.all(1, "#ffffff18"),
        content=ft.Row(
            spacing=10,
            controls=[
                ft.Icon(ft.Icons.RADAR_OUTLINED, color="#ffffff", size=18),
                status,
            ],
        ),
    )

    card_inputs = ft.Container(
        width=largura_bloco,
        padding=26,
        border_radius=34,
        gradient=ft.LinearGradient(
            begin=ft.Alignment(-1, -1),
            end=ft.Alignment(1, 1),
            colors=["#3d2466", "#2a1847", "#1d1238"],
        ),
        border=ft.Border.all(1, "#ffffff14"),
        shadow=[
            ft.BoxShadow(
                spread_radius=0,
                blur_radius=34,
                color="#12091f66",
                offset=ft.Offset(0, 18),
            ),
            ft.BoxShadow(
                spread_radius=0,
                blur_radius=10,
                color="#ffffff10",
                offset=ft.Offset(0, -2),
            ),
        ],
        content=ft.Column(
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=18,
            controls=[
                ft.Row(
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    controls=[
                        ft.Text(
                            "Configuracoes",
                            size=18,
                            weight=ft.FontWeight.W_700,
                            color="#ffffff",
                        ),
                        ft.Icon(ft.Icons.SETTINGS_ROUNDED, color="#dbc5ff"),
                    ],
                ),
                ft.Text(
                    "Preencha os dados abaixo para atualizar o servico automaticamente.",
                    size=12,
                    color="#d9c7ff",
                ),
                campo_cnpj,
                campo_certificado,
                campo_senha,
                campo_saida,
                mensagem_box,
            ],
        ),
    )

    botoes = ft.Column(
        width=largura_bloco,
        spacing=12,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        controls=[
            criar_botao(
                "Salvar configuracoes",
                ft.Icons.SAVE_ROUNDED,
                "#ffffff",
                "#301c50",
                largura_bloco,
                ao_salvar,
            ),
            criar_botao(
                "Recarregar layout",
                ft.Icons.REFRESH_ROUNDED,
                "#7a4dff",
                "#ffffff",
                largura_bloco,
                ao_recarregar,
            ),
            criar_botao(
                "Parar servico",
                ft.Icons.STOP_CIRCLE_OUTLINED,
                "#4f2d79",
                "#ffe8f0",
                largura_bloco,
                ao_parar,
            ),
            criar_botao(
                "Rodar servico",
                ft.Icons.PLAY_CIRCLE_OUTLINE_ROUNDED,
                "#2dc7b8",
                "#081a1f",
                largura_bloco,
                ao_rodar,
            ),
        ],
    )

    conteudo = ft.Container(
        expand=True,
        padding=ft.Padding.symmetric(horizontal=28, vertical=26),
        content=ft.Column(
            expand=True,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[
                topo,
                ft.Container(height=18),
                status_chip,
                ft.Container(height=18),
                card_inputs,
                ft.Container(expand=True),
                botoes,
                ft.Container(height=10),
            ],
        ),
    )

    fundo = ft.Stack(
        expand=True,
        controls=[
            ft.Container(
                expand=True,
                gradient=ft.LinearGradient(
                    begin=ft.Alignment(-1, -1),
                    end=ft.Alignment(1, 1),
                    colors=["#d4b7ff", "#9a6df5", "#6c4ce0"],
                ),
            ),
            ft.Container(
                width=280,
                height=280,
                left=-40,
                top=90,
                border_radius=140,
                blur=90,
                gradient=ft.LinearGradient(colors=["#ffffff88", "#ffffff00"]),
            ),
            ft.Container(
                width=240,
                height=240,
                right=-20,
                top=40,
                border_radius=120,
                blur=70,
                gradient=ft.LinearGradient(colors=["#6842cf", "#9f7eff11"]),
            ),
            ft.Container(
                width=210,
                height=210,
                right=10,
                bottom=120,
                border_radius=105,
                blur=80,
                gradient=ft.LinearGradient(colors=["#ffffff44", "#ffffff00"]),
            ),
            conteudo,
        ],
    )

    page.add(fundo)


if __name__ == "__main__":
    ft.run(main)
