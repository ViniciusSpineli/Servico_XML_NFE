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


# Arquivos principais controlados pelo painel:
# - baixar_xml_nfe.py: consulta a SEFAZ e salva os XMLs
# - monitorar_baixa_xml.py: roda em loop tentando executar a consulta
# - monitorar_baixa_xml.pid: guarda o PID do monitor para iniciar/parar o serviço
BASE_DIR = Path(__file__).resolve().parent
ARQUIVO_BAIXA = BASE_DIR / "baixar_xml_nfe.py"
ARQUIVO_MONITOR = BASE_DIR / "monitorar_baixa_xml.py"
ARQUIVO_SERVICO_BAT = BASE_DIR / "servico.bat"
ARQUIVO_PID = BASE_DIR / "monitorar_baixa_xml.pid"
ARQUIVO_LOGO = BASE_DIR / "fbs_2017.ico"


def ler_configuracao_atual() -> dict[str, str]:
    # Lê do baixar_xml_nfe.py os valores atuais que devem aparecer preenchidos na tela.
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
    # Substitui a linha inteira de uma variável no arquivo Python de configuração.
    if variavel in {
        "CAMINHO_CERTIFICADO_PFX",
        "CAMINHO_CERTIFICADO_PEM",
        "CAMINHO_CHAVE_PRIVADA_PEM",
        "CAMINHO_SALVAR_XML",
        "CAMINHO_ARQUIVO_ULT_NSU",
        "CAMINHO_ARQUIVO_PROXIMA_CONSULTA",
    }:
        valor_seguro = valor.replace("\\", "\\\\").replace('"', '\\"')
        linha_nova = f'{variavel} = r"{valor_seguro}"'
    else:
        linha_nova = f"{variavel} = {json.dumps(valor, ensure_ascii=False)}"
    padrao = rf"^{variavel}\s*=.*$"
    return re.sub(padrao, linha_nova, conteudo, flags=re.MULTILINE)


def salvar_configuracao(
    cnpj: str,
    caminho_certificado: str,
    senha: str,
    caminho_saida: str,
) -> None:
    # Persiste no baixar_xml_nfe.py o que foi preenchido nos inputs do painel.
    conteudo = ARQUIVO_BAIXA.read_text(encoding="utf-8")
    conteudo = substituir_linha(conteudo, "CNPJ_AUTOR", cnpj)
    conteudo = substituir_linha(conteudo, "CAMINHO_CERTIFICADO_PFX", caminho_certificado)
    conteudo = substituir_linha(conteudo, "SENHA_CERTIFICADO", senha)
    conteudo = substituir_linha(conteudo, "CAMINHO_SALVAR_XML", caminho_saida)
    ARQUIVO_BAIXA.write_text(conteudo, encoding="utf-8")


def ler_pid() -> int | None:
    # Se o monitor já estiver rodando, o PID salvo aqui permite pará-lo depois.
    if not ARQUIVO_PID.exists():
        return None

    try:
        return int(ARQUIVO_PID.read_text(encoding="utf-8").strip())
    except (ValueError, OSError):
        return None


def processo_ativo(pid: int | None) -> bool:
    # Checa se o PID salvo ainda corresponde a um processo ativo no sistema.
    if not pid:
        return False

    try:
        os.kill(pid, 0)
    except OSError:
        return False
    return True


def iniciar_monitor() -> str:
    # Inicia o monitor em segundo plano e salva o PID para controle futuro.
    pid_atual = ler_pid()
    if processo_ativo(pid_atual):
        return f"Servico ja esta rodando. PID: {pid_atual}"

    if not ARQUIVO_SERVICO_BAT.exists():
        return f"Arquivo nao encontrado: {ARQUIVO_SERVICO_BAT}"

    kwargs: dict[str, object] = {}
    if os.name == "nt":
        kwargs["creationflags"] = (
            subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS
        )

    processo = subprocess.Popen(
        ["cmd", "/c", str(ARQUIVO_SERVICO_BAT)],
        cwd=str(BASE_DIR),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        stdin=subprocess.DEVNULL,
        **kwargs,
    )
    ARQUIVO_PID.write_text(str(processo.pid), encoding="utf-8")
    return f"Servico iniciado com sucesso. PID: {processo.pid}"


def parar_monitor() -> str:
    # Para o monitor usando o PID salvo pelo painel.
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
    # Texto simples usado na interface para mostrar se o serviço está ligado.
    pid = ler_pid()
    if processo_ativo(pid):
        return f"Status do servico: rodando (PID {pid})"
    return "Status do servico: parado"


def recarregar_aplicacao() -> None:
    # Abre uma nova instância do painel para refletir alterações visuais/código.
    subprocess.Popen(
        [sys.executable, str(Path(__file__).resolve())],
        cwd=str(BASE_DIR),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        stdin=subprocess.DEVNULL,
    )


def obter_area_util_monitor_principal() -> tuple[int, int, int, int] | None:
    # No Windows, pega a área útil do monitor principal para dimensionar a janela.
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
    # Define tamanho, posição e limites mínimos da janela do Flet.
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
    page.bgcolor = "#e7eaef"
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
    # Fábrica visual dos campos de entrada para manter o estilo consistente.
    return ft.TextField(
        label=label,
        value=valor,
        width=largura,
        autofocus=autofocus,
        password=password,
        can_reveal_password=password,
        prefix_icon=icone,
        filled=True,
        bgcolor="#f5f7fa",
        color="#24313d",
        border=ft.InputBorder.OUTLINE,
        border_radius=20,
        border_color="#d9dee7",
        focused_border_color="#7b8da6",
        focused_bgcolor="#ffffff",
        content_padding=18,
        label_style=ft.TextStyle(color="#5d6d80", size=12, weight=ft.FontWeight.W_600),
    )


def criar_botao(
    texto: str,
    icone: ft.Icons,
    cor_fundo: str,
    cor_texto: str,
    largura: int,
    acao,
) -> ft.Button:
    # Fábrica visual dos botões principais do painel.
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
    # Função principal da interface. Monta a janela, carrega config e cria a tela.
    largura_janela = configurar_janela(page)
    largura_bloco = min(420, max(330, largura_janela - 86))
    config = ler_configuracao_atual()

    # Inputs principais que o usuário edita para alimentar o baixar_xml_nfe.py.
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

    # Bloco de mensagem/feedback usado para mostrar sucesso, erro e status do serviço.
    status = ft.Text(
        value=status_servico(),
        color="#334155",
        size=13,
        weight=ft.FontWeight.W_600,
    )
    mensagem = ft.Text(
        value="",
        color="#445468",
        size=12,
        selectable=True,
    )
    mensagem_box = ft.Container(
        width=largura_bloco,
        padding=ft.Padding.symmetric(horizontal=16, vertical=14),
        border_radius=18,
        bgcolor="#eef2f6",
        content=mensagem,
    )

    def atualizar_status(texto: str, cor: str, fundo: str) -> None:
        # Atualiza o feedback visual no painel sem reiniciar a aplicação.
        mensagem.value = texto
        mensagem.color = cor
        mensagem_box.bgcolor = fundo
        status.value = status_servico()
        page.update()

    def ao_salvar(_: ft.ControlEvent) -> None:
        # Botão "Salvar configurações": grava os dados do formulário no script de baixa.
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
        # Botão "Rodar serviço": inicia o monitor em segundo plano.
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
        # Botão "Parar serviço": encerra o monitor usando o PID salvo.
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
        # Botão "Recarregar layout": abre uma nova instância do painel e fecha a atual.
        try:
            recarregar_aplicacao()
            page.window.close()
        except Exception as exc:
            atualizar_status(
                f"Erro ao recarregar layout: {exc}",
                "#ffe1e8",
                "#5a2131",
            )

    # Cabeçalho superior com título do painel e ícone decorativo.
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
                        color="#263341",
                    ),
                    ft.Text(
                        "Controle do servico e configuracoes do download.",
                        size=12,
                        color="#66758a",
                    ),
                ],
            ),
            ft.Container(
                padding=ft.Padding.symmetric(horizontal=14, vertical=8),
                border_radius=18,
                bgcolor="#ffffffcc",
                border=ft.Border.all(1, "#d7dde6"),
                shadow=[
                    ft.BoxShadow(
                        blur_radius=16,
                        color="#a8b4c21f",
                        offset=ft.Offset(0, 8),
                    )
                ],
                content=ft.Icon(ft.Icons.CLOUD_SYNC_OUTLINED, color="#64748b"),
            ),
        ],
    )

    # Cartão pequeno que mostra rapidamente o estado do serviço.
    status_chip = ft.Container(
        padding=ft.Padding.symmetric(horizontal=16, vertical=10),
        border_radius=18,
        bgcolor="#ffffffd9",
        border=ft.Border.all(1, "#d8dee7"),
        shadow=[
            ft.BoxShadow(
                blur_radius=18,
                color="#a8b4c22b",
                offset=ft.Offset(0, 10),
            )
        ],
        content=ft.Row(
            spacing=10,
            controls=[
                ft.Icon(ft.Icons.RADAR_OUTLINED, color="#64748b", size=18),
                status,
            ],
        ),
    )

    # Card principal com relevo onde ficam os 4 inputs e a área de mensagens.
    card_inputs = ft.Container(
        width=largura_bloco,
        padding=26,
        border_radius=34,
        bgcolor="#ffffffee",
        border=ft.Border.all(1, "#dbe2ea"),
        shadow=[
            ft.BoxShadow(
                spread_radius=0,
                blur_radius=34,
                color="#93a1b333",
                offset=ft.Offset(0, 18),
            ),
            ft.BoxShadow(
                spread_radius=0,
                blur_radius=10,
                color="#ffffffcc",
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
                            color="#253241",
                        ),
                        ft.Icon(ft.Icons.SETTINGS_ROUNDED, color="#6f7f93"),
                    ],
                ),
                ft.Text(
                    "Preencha os dados abaixo para atualizar o servico automaticamente.",
                    size=12,
                    color="#68778a",
                ),
                campo_cnpj,
                campo_certificado,
                campo_senha,
                campo_saida,
                mensagem_box,
            ],
        ),
    )

    # Botões de ação, empilhados no rodapé para combinar com o visual do painel.
    botoes = ft.Column(
        width=largura_bloco,
        spacing=12,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        controls=[
            criar_botao(
                "Salvar configuracoes",
                ft.Icons.SAVE_ROUNDED,
                "#ffffff",
                "#24313d",
                largura_bloco,
                ao_salvar,
            ),
           
            criar_botao(
                "Parar servico",
                ft.Icons.STOP_CIRCLE_OUTLINED,
                "#ced6df",
                "#2d3a47",
                largura_bloco,
                ao_parar,
            ),
            criar_botao(
                "Rodar servico",
                ft.Icons.PLAY_CIRCLE_OUTLINE_ROUNDED,
                "#8ea4ba",
                "#ffffff",
                largura_bloco,
                ao_rodar,
            ),
        ],
    )

    # Conteúdo central da tela. Tudo que fica acima do background decorativo.
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

    # Fundo em camadas:
    # - gradiente cinza claro
    # - marca d'água com o logo
    # - elementos suaves de luz
    # - conteúdo principal por cima
    fundo = ft.Stack(
        expand=True,
        controls=[
            ft.Container(
                expand=True,
                gradient=ft.LinearGradient(
                    begin=ft.Alignment(-1, -1),
                    end=ft.Alignment(1, 1),
                    colors=["#f2f4f7", "#e4e8ed", "#d9dfe6"],
                ),
            ),
            ft.Container(
                content=ft.Image(
                    src=str(ARQUIVO_LOGO),
                    width=420,
                    height=420,
                    fit=ft.BoxFit.CONTAIN,
                    opacity=0.08,
                ),
                align=ft.Alignment(0.65, -0.1),
            ),
            
            conteudo,
        ],
    )

    page.add(fundo)


if __name__ == "__main__":
    # Ponto de entrada do aplicativo Flet.
    ft.run(main)
