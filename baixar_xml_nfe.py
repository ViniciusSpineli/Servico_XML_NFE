from __future__ import annotations

import base64
import csv
import gzip
import os
import re
import sys
import tempfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path


# =========================
# PREENCHA ESTES DADOS
# =========================
CNPJ_AUTOR = "26.826.794/0001-30"
UF_AUTOR = "35"  # Codigo IBGE da UF do autor. Ex.: 35=SP, 41=PR, 43=RS
AMBIENTE = "producao"  # "producao" ou "homologacao"

# Opcao 1: certificado A1 em PFX/P12
CAMINHO_CERTIFICADO_PFX = r"C:\Users\vinic\OneDrive\Área de Trabalho\Projetos_Python\DANFE\GV_BONELLO_SENHA_123456.pfx"
SENHA_CERTIFICADO = "123456"

# Opcao 2: se voce ja tiver PEM separado, preencha os dois abaixo e ignore o PFX
CAMINHO_CERTIFICADO_PEM = ""
CAMINHO_CHAVE_PRIVADA_PEM = ""
CAMINHO_ARQUIVO_ULT_NSU = r""
CAMINHO_ARQUIVO_PROXIMA_CONSULTA = r""

CAMINHO_SALVAR_XML = r"C:\Users\vinic\OneDrive\Área de Trabalho\Nova pasta"

MAX_CONSULTAS = 30
TIMEOUT_SEGUNDOS = 60


SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/"
WSDL_NS = "http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe"
NFE_NS = "http://www.portalfiscal.inf.br/nfe"
NS = {"soap": SOAP_NS, "wsdl": WSDL_NS, "nfe": NFE_NS}

URLS = {
    "producao": "https://www1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx",
    "homologacao": "https://hom1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx",
}


def somente_digitos(valor: str) -> str:
    return re.sub(r"\D", "", valor or "")


def obter_area_de_trabalho() -> Path:
    candidatos = []

    one_drive = os.environ.get("OneDrive")
    if one_drive:
        candidatos.append(Path(one_drive) / "Área de Trabalho")
        candidatos.append(Path(one_drive) / "Desktop")

    home = Path.home()
    candidatos.append(home / "Desktop")
    candidatos.append(home / "Área de Trabalho")

    for caminho in candidatos:
        if caminho.exists():
            return caminho

    return Path.cwd()


def obter_pasta_saida() -> Path:
    caminho_configurado = CAMINHO_SALVAR_XML.strip()
    if caminho_configurado:
        return Path(caminho_configurado).expanduser()
    return obter_area_de_trabalho()


def obter_arquivo_ult_nsu() -> Path:
    caminho_configurado = CAMINHO_ARQUIVO_ULT_NSU.strip()
    if caminho_configurado:
        return Path(caminho_configurado).expanduser()
    return Path(__file__).with_name("ult_nsu_sefaz.txt")


def obter_arquivo_proxima_consulta() -> Path:
    caminho_configurado = CAMINHO_ARQUIVO_PROXIMA_CONSULTA.strip()
    if caminho_configurado:
        return Path(caminho_configurado).expanduser()
    return Path(__file__).with_name("proxima_consulta_sefaz.txt")


PASTA_SAIDA = obter_pasta_saida()
ARQUIVO_ULT_NSU = obter_arquivo_ult_nsu()
ARQUIVO_PROXIMA_CONSULTA = obter_arquivo_proxima_consulta()


def validar_configuracao() -> tuple[str, str, str]:
    cnpj = somente_digitos(CNPJ_AUTOR)
    uf = somente_digitos(UF_AUTOR)
    ambiente = AMBIENTE.strip().lower()

    if len(cnpj) != 14:
        raise ValueError("O CNPJ_AUTOR precisa ter exatamente 14 digitos.")

    if len(uf) != 2:
        raise ValueError("A UF_AUTOR precisa ter 2 digitos do codigo IBGE da UF.")

    if ambiente not in URLS:
        raise ValueError('AMBIENTE deve ser "producao" ou "homologacao".')

    if not possui_certificado_configurado():
        raise ValueError(
            "Configure o certificado em PFX/P12 ou informe os caminhos PEM."
        )

    return cnpj, uf, ambiente


def possui_certificado_configurado() -> bool:
    if CAMINHO_CERTIFICADO_PEM and CAMINHO_CHAVE_PRIVADA_PEM:
        return True
    return bool(CAMINHO_CERTIFICADO_PFX and SENHA_CERTIFICADO)


def formatar_nsu(nsu: str) -> str:
    return somente_digitos(nsu).zfill(15)


def carregar_ult_nsu() -> str:
    if not ARQUIVO_ULT_NSU.exists():
        return "0"

    conteudo = ARQUIVO_ULT_NSU.read_text(encoding="utf-8").strip()
    if not conteudo:
        return "0"
    return formatar_nsu(conteudo)


def salvar_ult_nsu(nsu: str) -> None:
    ARQUIVO_ULT_NSU.parent.mkdir(parents=True, exist_ok=True)
    ARQUIVO_ULT_NSU.write_text(formatar_nsu(nsu), encoding="utf-8")


def agora_local() -> datetime:
    return datetime.now().astimezone()


def carregar_proxima_consulta() -> datetime | None:
    if not ARQUIVO_PROXIMA_CONSULTA.exists():
        return None

    conteudo = ARQUIVO_PROXIMA_CONSULTA.read_text(encoding="utf-8").strip()
    if not conteudo:
        return None

    try:
        return datetime.fromisoformat(conteudo)
    except ValueError:
        return None


def salvar_proxima_consulta(data_hora: datetime) -> None:
    ARQUIVO_PROXIMA_CONSULTA.parent.mkdir(parents=True, exist_ok=True)
    ARQUIVO_PROXIMA_CONSULTA.write_text(data_hora.isoformat(), encoding="utf-8")


def limpar_proxima_consulta() -> None:
    try:
        ARQUIVO_PROXIMA_CONSULTA.unlink(missing_ok=True)
    except OSError:
        pass


def formatar_data_hora_local(data_hora: datetime) -> str:
    return data_hora.astimezone().strftime("%d/%m/%Y %H:%M:%S")


def formatar_tempo_restante(data_hora: datetime) -> str:
    restante = data_hora - agora_local()
    total_segundos = max(0, int(restante.total_seconds()))
    horas, resto = divmod(total_segundos, 3600)
    minutos, segundos = divmod(resto, 60)

    partes = []
    if horas:
        partes.append(f"{horas}h")
    if minutos or horas:
        partes.append(f"{minutos}min")
    partes.append(f"{segundos}s")
    return " ".join(partes)


def interpretar_dhresp(valor: str) -> datetime | None:
    texto = (valor or "").strip()
    if not texto:
        return None

    if texto.endswith("Z"):
        texto = texto[:-1] + "+00:00"

    try:
        data_hora = datetime.fromisoformat(texto)
    except ValueError:
        return None

    if data_hora.tzinfo is None:
        return data_hora.astimezone()
    return data_hora


def montar_envelope(cnpj: str, uf: str, ambiente: str, ult_nsu: str) -> str:
    tp_amb = "1" if ambiente == "producao" else "2"
    ult_nsu = formatar_nsu(ult_nsu)

    return f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:soap="{SOAP_NS}">
  <soap:Body>
    <nfeDistDFeInteresse xmlns="{WSDL_NS}">
      <nfeDadosMsg xmlns="{WSDL_NS}">
        <distDFeInt xmlns="{NFE_NS}" versao="1.01">
          <tpAmb>{tp_amb}</tpAmb>
          <cUFAutor>{uf}</cUFAutor>
          <CNPJ>{cnpj}</CNPJ>
          <distNSU>
            <ultNSU>{ult_nsu}</ultNSU>
          </distNSU>
        </distDFeInt>
      </nfeDadosMsg>
    </nfeDistDFeInteresse>
  </soap:Body>
</soap:Envelope>"""


def importar_requests():
    try:
        import requests
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError(
            'Biblioteca ausente: requests. Instale com "pip install requests".'
        ) from exc
    return requests


def importar_pkcs12():
    try:
        from cryptography.hazmat.primitives import serialization
        from cryptography.hazmat.primitives.serialization import pkcs12
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError(
            'Biblioteca ausente: cryptography. Instale com "pip install cryptography".'
        ) from exc
    return serialization, pkcs12


def carregar_certificado() -> tuple[tuple[str, str], list[Path]]:
    if CAMINHO_CERTIFICADO_PEM and CAMINHO_CHAVE_PRIVADA_PEM:
        cert = Path(CAMINHO_CERTIFICADO_PEM)
        chave = Path(CAMINHO_CHAVE_PRIVADA_PEM)
        if not cert.exists():
            raise FileNotFoundError(f"Certificado PEM nao encontrado: {cert}")
        if not chave.exists():
            raise FileNotFoundError(f"Chave privada PEM nao encontrada: {chave}")
        return (str(cert), str(chave)), []

    pfx = Path(CAMINHO_CERTIFICADO_PFX)
    if not pfx.exists():
        raise FileNotFoundError(f"Certificado PFX/P12 nao encontrado: {pfx}")

    serialization, pkcs12 = importar_pkcs12()
    pfx_bytes = pfx.read_bytes()
    senha = SENHA_CERTIFICADO.encode("utf-8") if SENHA_CERTIFICADO else None
    chave_privada, certificado, adicionais = pkcs12.load_key_and_certificates(
        pfx_bytes,
        senha,
    )

    if chave_privada is None or certificado is None:
        raise ValueError("Nao foi possivel ler chave privada e certificado do PFX.")

    key_pem = chave_privada.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.TraditionalOpenSSL,
        encryption_algorithm=serialization.NoEncryption(),
    )
    cert_pem = certificado.public_bytes(serialization.Encoding.PEM)

    for item in adicionais or []:
        cert_pem += item.public_bytes(serialization.Encoding.PEM)

    arquivos_temporarios: list[Path] = []
    cert_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pem")
    key_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pem")

    try:
        cert_temp.write(cert_pem)
        key_temp.write(key_pem)
        cert_temp.close()
        key_temp.close()
        cert_path = Path(cert_temp.name)
        key_path = Path(key_temp.name)
        arquivos_temporarios.extend([cert_path, key_path])
        return (str(cert_path), str(key_path)), arquivos_temporarios
    except Exception:
        cert_temp.close()
        key_temp.close()
        raise


def consultar_distribuicao(envelope: str, ambiente: str, cert_files: tuple[str, str]) -> bytes:
    requests = importar_requests()
    headers = {
        "Content-Type": "text/xml; charset=utf-8",
        "SOAPAction": f'"{WSDL_NS}/nfeDistDFeInteresse"',
    }
    response = requests.post(
        URLS[ambiente],
        data=envelope.encode("utf-8"),
        headers=headers,
        cert=cert_files,
        timeout=TIMEOUT_SEGUNDOS,
    )
    if not response.content and response.status_code >= 400:
        response.raise_for_status()
    return response.content


def extrair_retorno(xml_bytes: bytes) -> ET.Element:
    root = ET.fromstring(xml_bytes)

    fault = root.find(".//soap:Fault", NS)
    if fault is not None:
        fault_text = " ".join(text.strip() for text in fault.itertext() if text.strip())
        raise RuntimeError(f"SOAP Fault retornado pela SEFAZ: {fault_text}")

    ret = root.find(".//nfe:retDistDFeInt", NS)
    if ret is not None:
        return ret

    inner = root.find(".//wsdl:nfeDistDFeInteresseResult", NS)
    if inner is not None and inner.text and inner.text.strip():
        return ET.fromstring(inner.text.strip())

    raise RuntimeError("Nao foi possivel localizar o retDistDFeInt na resposta da SEFAZ.")


def encontrar_texto(elemento: ET.Element, tag_local: str) -> str:
    encontrado = elemento.find(f".//{{{NFE_NS}}}{tag_local}")
    if encontrado is None or encontrado.text is None:
        return ""
    return encontrado.text.strip()


def interpretar_data_documento(valor: str) -> datetime | None:
    texto = (valor or "").strip()
    if not texto:
        return None

    if texto.endswith("Z"):
        texto = texto[:-1] + "+00:00"

    try:
        return datetime.fromisoformat(texto)
    except ValueError:
        return None


def obter_pasta_competencia(xml_bytes: bytes) -> str:
    tags_preferidas = (
        "dhEvento",
        "dhEmi",
        "dEmi",
        "dhRecbto",
        "dhRegEvento",
        "dhSaiEnt",
    )

    try:
        raiz = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return "sem-data"

    for tag_preferida in tags_preferidas:
        for elem in raiz.iter():
            local = elem.tag.rsplit("}", 1)[-1]
            if local != tag_preferida or not elem.text:
                continue

            data_hora = interpretar_data_documento(elem.text)
            if data_hora is not None:
                return data_hora.strftime("%m-%Y")

    return "sem-data"


def extrair_chave_acesso(xml_bytes: bytes) -> str:
    try:
        raiz = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return ""

    chave = ""
    for elem in raiz.iter():
        local = elem.tag.rsplit("}", 1)[-1]
        if local == "chNFe" and elem.text:
            chave = somente_digitos(elem.text)
            if len(chave) == 44:
                return chave
        if local == "infNFe":
            chave = somente_digitos(elem.attrib.get("Id", "").replace("NFe", ""))
            if len(chave) == 44:
                return chave
        if local == "infEvento":
            chave = somente_digitos(elem.attrib.get("Id", ""))
            if len(chave) >= 44:
                return chave[-44:]

    return ""


def obter_texto_local(elemento: ET.Element, tags_locais: tuple[str, ...]) -> str:
    for elem in elemento.iter():
        local = elem.tag.rsplit("}", 1)[-1]
        if local in tags_locais and elem.text:
            texto = elem.text.strip()
            if texto:
                return texto
    return ""


def obter_texto_em_grupo(
    elemento: ET.Element,
    grupo_local: str,
    tags_locais: tuple[str, ...],
) -> str:
    for grupo in elemento.iter():
        local = grupo.tag.rsplit("}", 1)[-1]
        if local != grupo_local:
            continue

        for filho in grupo.iter():
            local_filho = filho.tag.rsplit("}", 1)[-1]
            if local_filho in tags_locais and filho.text:
                texto = filho.text.strip()
                if texto:
                    return texto
    return ""


def extrair_resumo_documento(xml_bytes: bytes, nsu: str) -> dict[str, str]:
    resumo = {
        "NSU": formatar_nsu(nsu),
        "Valor": "",
        "DtEmissao": "",
        "Emitente": "",
        "CNPJEmitente": "",
        "Chave": extrair_chave_acesso(xml_bytes),
        "NumeroNota": "",
    }

    try:
        raiz = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return resumo

    resumo["Valor"] = obter_texto_local(raiz, ("vNF", "vEvento"))
    resumo["DtEmissao"] = obter_texto_local(
        raiz,
        ("dhEmi", "dEmi", "dhEvento", "dhRecbto", "dhRegEvento", "dhSaiEnt"),
    )
    resumo["Emitente"] = (
        obter_texto_em_grupo(raiz, "emit", ("xNome",))
        or obter_texto_local(raiz, ("xNome", "xNomeEmit"))
    )
    resumo["CNPJEmitente"] = (
        obter_texto_em_grupo(raiz, "emit", ("CNPJ", "CPF"))
        or somente_digitos(obter_texto_local(raiz, ("CNPJ", "CPF", "CNPJEmit")))
    )
    resumo["NumeroNota"] = obter_texto_local(raiz, ("nNF",))

    if resumo["CNPJEmitente"]:
        resumo["CNPJEmitente"] = somente_digitos(resumo["CNPJEmitente"])

    return resumo


def nomear_documento(xml_bytes: bytes, schema: str, nsu: str, indice: int) -> str:
    chave = extrair_chave_acesso(xml_bytes)
    if chave:
        return f"{chave}.xml"

    schema_limpo = Path(schema).stem.replace(".", "_")
    return f"nsu-{formatar_nsu(nsu)}-{schema_limpo}-{indice}.xml"


def extrair_documentos(retorno: ET.Element) -> list[dict[str, str | bytes]]:
    documentos: list[dict[str, str | bytes]] = []

    for indice, doc_zip in enumerate(retorno.findall(f".//{{{NFE_NS}}}docZip"), start=1):
        schema = doc_zip.attrib.get("schema", "documento.xsd")
        nsu = doc_zip.attrib.get("NSU", "")
        conteudo_b64 = "".join(doc_zip.itertext()).strip()
        if not conteudo_b64:
            continue

        xml_compactado = base64.b64decode(conteudo_b64)
        xml_descompactado = gzip.decompress(xml_compactado)
        documentos.append(
            {
                "schema": schema,
                "nsu": formatar_nsu(nsu),
                "xml_bytes": xml_descompactado,
                "indice_lote": str(indice),
                "resumo": extrair_resumo_documento(xml_descompactado, nsu),
            }
        )

    return documentos


def salvar_resumo_csv(documentos: list[dict[str, str | bytes]]) -> Path | None:
    if not documentos:
        return None

    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    caminho_csv = PASTA_SAIDA / f"resumo_xmls_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    campos = ["NSU", "Valor", "DtEmissao", "Emitente", "CNPJEmitente", "Chave", "NumeroNota"]

    with caminho_csv.open("w", encoding="utf-8-sig", newline="") as arquivo:
        writer = csv.DictWriter(arquivo, fieldnames=campos, delimiter=";")
        writer.writeheader()
        for documento in documentos:
            resumo = documento.get("resumo", {})
            writer.writerow({campo: str(resumo.get(campo, "")) for campo in campos})

    return caminho_csv


def salvar_documentos(documentos: list[dict[str, str | bytes]]) -> tuple[list[Path], Path | None]:
    PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
    arquivos: list[Path] = []

    for indice, documento in enumerate(documentos, start=1):
        xml_bytes = documento["xml_bytes"]
        schema = str(documento["schema"])
        nsu = str(documento["nsu"])
        nome_arquivo = nomear_documento(xml_bytes, schema, nsu, indice)
        pasta_competencia = PASTA_SAIDA / obter_pasta_competencia(xml_bytes)
        pasta_competencia.mkdir(parents=True, exist_ok=True)
        destino = pasta_competencia / nome_arquivo
        destino.write_bytes(xml_bytes)
        arquivos.append(destino)

    return arquivos, salvar_resumo_csv(documentos)


def consultar_ultimos_documentos(
    cnpj: str,
    uf: str,
    ambiente: str,
    cert_files: tuple[str, str],
) -> tuple[list[dict[str, str | bytes]], str, str]:
    ult_nsu = carregar_ult_nsu()
    documentos: list[dict[str, str | bytes]] = []
    ultimo_cstat = ""
    ultimo_xmotivo = ""

    for tentativa in range(1, MAX_CONSULTAS + 1):
        envelope = montar_envelope(cnpj, uf, ambiente, ult_nsu)
        resposta = consultar_distribuicao(envelope, ambiente, cert_files)
        retorno = extrair_retorno(resposta)

        cstat = encontrar_texto(retorno, "cStat")
        xmotivo = encontrar_texto(retorno, "xMotivo")
        dh_resp = encontrar_texto(retorno, "dhResp")
        ult_nsu_retorno = formatar_nsu(encontrar_texto(retorno, "ultNSU") or ult_nsu)
        max_nsu = formatar_nsu(encontrar_texto(retorno, "maxNSU") or ult_nsu_retorno)
        ultimo_cstat = cstat
        ultimo_xmotivo = xmotivo
        documentos.extend(extrair_documentos(retorno))
        salvar_ult_nsu(ult_nsu_retorno)

        if cstat == "656":
            base_bloqueio = interpretar_dhresp(dh_resp) or agora_local()
            salvar_proxima_consulta(base_bloqueio + timedelta(hours=1))
            break

        if cstat == "137":
            limpar_proxima_consulta()
            break

        if ult_nsu_retorno == max_nsu:
            limpar_proxima_consulta()
            break

        if ult_nsu_retorno == formatar_nsu(ult_nsu):
            limpar_proxima_consulta()
            break

        ult_nsu = ult_nsu_retorno

    documentos_ordenados = sorted(documentos, key=lambda item: str(item["nsu"]))
    return documentos_ordenados, ultimo_cstat, ultimo_xmotivo


def limpar_temporarios(arquivos: list[Path]) -> None:
    for arquivo in arquivos:
        try:
            arquivo.unlink(missing_ok=True)
        except OSError:
            pass


def main() -> int:
    temporarios: list[Path] = []

    try:
        proxima_consulta = carregar_proxima_consulta()
        if proxima_consulta and agora_local() < proxima_consulta:
            print(
                "Proxima consulta permitida pela SEFAZ: "
                f"{formatar_data_hora_local(proxima_consulta)} "
                f"(faltam {formatar_tempo_restante(proxima_consulta)})."
            )
            return 1

        cnpj, uf, ambiente = validar_configuracao()
        cert_files, temporarios = carregar_certificado()
        documentos, cstat_final, motivo_final = consultar_ultimos_documentos(
            cnpj,
            uf,
            ambiente,
            cert_files,
        )

        arquivos, arquivo_resumo = salvar_documentos(documentos)
        if arquivos:
            print(f"Arquivos salvos: {len(arquivos)}")
            for arquivo in arquivos:
                print(f"- {arquivo.resolve()}")
            if arquivo_resumo:
                print(f"\nResumo CSV: {arquivo_resumo.resolve()}")

            if any("resNFe" in arquivo.name for arquivo in arquivos):
                print(
                    "\nAviso: a SEFAZ pode retornar apenas o resumo da NF-e "
                    "em vez do XML completo, dependendo da sua permissao no documento."
                )
            return 0

        if cstat_final == "656":
            proxima_consulta = carregar_proxima_consulta()
            horario = (
                formatar_data_hora_local(proxima_consulta)
                if proxima_consulta
                else "daqui a 1 hora"
            )
            print(
                "A SEFAZ retornou Consumo Indevido (656). O ultNSU mais recente foi "
                f"salvo em {ARQUIVO_ULT_NSU.resolve()}. Proxima tentativa: {horario}."
            )
            return 1

        print(
            "\nNenhum documento foi retornado. Isso normalmente significa que nao ha "
            "XML novo disponivel para esse certificado ou a SEFAZ nao liberou documentos "
            f"para esse ator. Ultimo retorno SEFAZ: {cstat_final} - {motivo_final}"
        )
        return 1

    except Exception as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        return 1
    finally:
        limpar_temporarios(temporarios)


if __name__ == "__main__":
    raise SystemExit(main())
