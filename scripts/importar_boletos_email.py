"""
Importa boletos que chegam por e-mail e cria (ou completa) o lançamento no CAP.

De onde vêm as regras (nada fica escrito neste arquivo — o repositório é público):
  - Tela "Boletos por e-mail" do CAP -> TabelaImportacao da planilha financeira:
      [Id, Ativo, Remetente, Assunto contém, Texto no PDF, Despesa de, Descrição,
       Tipo, Senha do PDF, Atualizado em]
    Remetente pode ser o e-mail inteiro ou só o domínio (ex.: "claro.com.br").
    "Texto no PDF" separa contas do mesmo remetente (código do cliente, contrato,
    número da unidade...). Sem ele, a regra vale para todo PDF daquele remetente.
  - Tela "Emissão de Documentos" -> TabelaEmissao: cada linha da Neoenergia com
    Código vira uma regra (remetente "neoenergia", texto no PDF = código).

Fluxo (1x por dia via .github/workflows/importar-boletos-email.yml):
  1. Lê os e-mails dos últimos DIAS_BUSCA dias com anexo na caixa CAIXA_EMAIL.
  2. Para cada PDF de um remetente com regra: acha a regra pelo texto, lê o código
     de barras (conta de consumo 48 dígitos ou boleto bancário 47 dígitos — de
     onde saem valor e, no bancário, o vencimento), vencimento e valor escritos.
  3. Lançamento do mesmo mês (mesma pessoa, descrição parecida):
       pago ou com o mesmo código de barras -> não mexe;
       a vencer -> completa vencimento, valor, código de barras e boleto.
     Sem lançamento no mês -> cria "A Vencer" com o boleto anexado.
  4. Manda um resumo por e-mail quando importou algo ou não conseguiu ler algo.

Simular (não grava nada): SIMULAR=1.

Permissões de APLICATIVO no App Registration "Gestão Financeira Claude":
Mail.Read, Mail.Send e Files.ReadWrite.All.

Variáveis de ambiente: CAP_CLIENT_SECRET (obrigatória), NEO_PDF_SENHAS (opcional,
senhas extras separadas por vírgula), SIMULAR (opcional).
"""

import base64
import difflib
import io
import os
import re
import sys
import unicodedata
from datetime import date, datetime, timedelta

import msal
import requests
from pypdf import PdfReader

try:
    from zoneinfo import ZoneInfo
    HOJE = datetime.now(ZoneInfo("America/Sao_Paulo")).date()
except Exception:
    HOJE = date.today()

TENANT_ID = "0ea62e2f-152a-4380-bf4b-497083aa0326"
CLIENT_ID = "9291254e-8c79-4641-8d7d-c5771d82ccde"
CLIENT_SECRET = os.environ.get("CAP_CLIENT_SECRET")
if not CLIENT_SECRET:
    sys.exit("ERRO: defina a variável de ambiente CAP_CLIENT_SECRET antes de rodar este script.")

DRIVE_ID = "b!Xyf9Y4l0qUWv1kO5T5rKOn3FJmBOsbRFnpjvpPbvNvrn80E1xObESKagjmnZsZfm"
ITEM_ID = "01ZF6ARXTIQSWRHIKDQVHIHZRFMV5YVCAD"
ANEXOS_PATH = "Gestão Financeira Claude/Anexos"
CAIXA_EMAIL = "natanael.silva@sinerggia.com.br"
LINK_DASHBOARD = "https://sinerggia-dev.github.io/cap-contas-a-pagar/"
DIAS_BUSCA = 40
SIMULAR = os.environ.get("SIMULAR", "").strip().lower() in ("1", "sim", "true", "yes")

GRAPH = "https://graph.microsoft.com/v1.0"


# ---------------------------------------------------------------- Graph

def get_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}", client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Falha ao obter token: {result.get('error_description')}")
    return result["access_token"]


def ler_tabela(headers, nome_tabela, obrigatoria=True):
    url = f"{GRAPH}/drives/{DRIVE_ID}/items/{ITEM_ID}/workbook/tables/{nome_tabela}/rows"
    linhas = []
    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code == 404 and not obrigatoria:
            return []
        resp.raise_for_status()
        dados = resp.json()
        linhas.extend({"index": item["index"], "values": item["values"][0]} for item in dados.get("value", []))
        url = dados.get("@odata.nextLink")
    return linhas


# ---------------------------------------------------------------- texto

def sem_acento(t):
    t = unicodedata.normalize("NFD", str(t or ""))
    return "".join(c for c in t if unicodedata.category(c) != "Mn").lower()


def chave_desc(t):
    t = re.sub(r"\s+-\s*|\s*-\s+", " - ", str(t or ""))
    return re.sub(r"\s+", " ", sem_acento(t)).strip()


def parecida(a, b):
    """Mesma descrição com erro de digitação ("Casa Gravatá 21" x "Gravtá 21").
    Os números têm de ser iguais: "Apto 102" e "Apto 201" são unidades diferentes."""
    ka, kb = chave_desc(a), chave_desc(b)
    if re.findall(r"\d+", ka) != re.findall(r"\d+", kb):
        return False
    return difflib.SequenceMatcher(None, ka, kb).ratio() >= 0.75


def so_digitos(t):
    return re.sub(r"\D", "", str(t or ""))


def serial_para_data(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return date(1899, 12, 30) + timedelta(days=int(v))
    txt = str(v).strip()
    if txt.replace(".", "", 1).isdigit():
        return date(1899, 12, 30) + timedelta(days=int(float(txt)))
    try:
        return datetime.strptime(txt, "%d/%m/%Y").date()
    except ValueError:
        return None


def data_para_serial(d):
    return (d - date(1899, 12, 30)).days


def valor_br(txt):
    return float(txt.replace(".", "").replace(",", "."))


def brl(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def ativo(v):
    return sem_acento(v).strip() not in ("nao", "não", "false", "0", "inativo")


# ---------------------------------------------------------------- regras

def carregar_regras(headers):
    regras = []
    # Tela "Boletos por e-mail".
    for linha in ler_tabela(headers, "TabelaImportacao", obrigatoria=False):
        v = [str(x or "").strip() for x in linha["values"]] + [""] * 10
        if not v[2] or not ativo(v[1]) or not v[5] or not v[6]:
            continue
        regras.append({"remetente": v[2].lower(), "assunto": v[3], "texto": v[4], "nome": v[5],
                       "desc": v[6], "tipo": v[7] or "Outras", "senhas": [v[8]] if v[8] else []})
    # Neoenergia a partir da Emissão de Documentos.
    for linha in ler_tabela(headers, "TabelaEmissao", obrigatoria=False):
        v = [str(x or "").strip() for x in linha["values"]] + [""] * 11
        desc, nome, link, usuario, codigo = v[1], v[2], v[4], v[5], v[7]
        if "neoenergia" not in (link + " " + desc).lower() or not so_digitos(codigo):
            continue
        d = so_digitos(usuario)
        regras.append({"remetente": "neoenergia", "assunto": "", "texto": codigo, "nome": nome,
                       "desc": desc, "tipo": "Energia",
                       "senhas": [d[:5], d[:6], d[:4], d] if d else []})
    return regras


def texto_bate(alvo, texto, texto_digitos):
    """Texto da regra aparece no PDF. Se for numero (codigo), compara so os digitos,
    sem zeros a esquerda; senao, compara sem acento e sem maiusculas."""
    alvo = str(alvo or "").strip()
    if not alvo:
        return True
    dig = so_digitos(alvo)
    if dig and len(dig) >= max(4, len(re.sub(r"\s", "", alvo)) * 0.6):
        return dig.lstrip("0") in texto_digitos
    return re.sub(r"\s+", " ", sem_acento(alvo)) in re.sub(r"\s+", " ", sem_acento(texto))


# ---------------------------------------------------------------- PDF

def texto_pdf(conteudo, senhas):
    leitor = PdfReader(io.BytesIO(conteudo))
    if leitor.is_encrypted:
        for s in [""] + senhas:
            try:
                if leitor.decrypt(s):
                    break
            except Exception:
                continue
        else:
            raise ValueError("PDF protegido por senha (nenhuma senha conhecida abriu)")
    return "\n".join((p.extract_text() or "") for p in leitor.pages)


RE_CONSUMO = re.compile(  # conta de consumo / arrecadação: 4 blocos de 11 + DV, começa com 8
    r"(8\d{10})\s*-?\s*(\d)\s*(\d{11})\s*-?\s*(\d)\s*(\d{11})\s*-?\s*(\d)\s*(\d{11})\s*-?\s*(\d)"
)
RE_BANCARIO = re.compile(  # boleto bancário: 47 dígitos
    r"(\d{5})\.?(\d{5})\s+(\d{5})\.?(\d{6})\s+(\d{5})\.?(\d{6})\s+(\d)\s+(\d{14})"
)


def data_do_fator(fator):
    """Fator de vencimento do boleto bancário. A contagem recomeçou em 22/02/2025
    (fator 1000); fica a data das duas bases que cair mais perto de hoje."""
    if fator < 1000:
        return None
    candidatas = [date(1997, 10, 7) + timedelta(days=fator), date(2025, 2, 22) + timedelta(days=fator - 1000)]
    return min(candidatas, key=lambda d: abs((d - HOJE).days))


def ler_boleto(texto):
    achado = {}
    m = RE_CONSUMO.search(texto)
    if m:
        g = m.groups()
        achado["barra"] = "".join(g)
        codigo44 = g[0] + g[2] + g[4] + g[6]
        if codigo44[2] in "68":  # valor efetivo em centavos nas posições 5-15
            achado["valor_barra"] = int(codigo44[4:15]) / 100
    else:
        m = RE_BANCARIO.search(texto)
        if m:
            g = m.groups()
            achado["barra"] = "".join(g)
            fator, valor = int(g[7][:4]), int(g[7][4:]) / 100
            if valor > 0:
                achado["valor_barra"] = valor
            venc = data_do_fator(fator)
            if venc:
                achado["venc_barra"] = venc

    plano = re.sub(r"[ \t]+", " ", texto)
    for padrao in (
        r"TOTAL\s+A\s+PAGAR[^\d]{0,40}?([\d.]+,\d{2})",
        r"VALOR\s+A\s+PAGAR[^\d]{0,40}?([\d.]+,\d{2})",
        r"VALOR\s+DO\s+DOCUMENTO[^\d]{0,40}?([\d.]+,\d{2})",
        r"VALOR\s+COBRADO[^\d]{0,40}?([\d.]+,\d{2})",
        r"VALOR\s+TOTAL[^\d]{0,40}?([\d.]+,\d{2})",
        r"TOTAL\s+DA\s+FATURA[^\d]{0,40}?([\d.]+,\d{2})",
    ):
        m = re.search(padrao, plano, re.I)
        if m:
            achado["valor_texto"] = valor_br(m.group(1))
            break
    m = re.search(r"VENCIMENTO[^\d]{0,40}?(\d{2}/\d{2}/\d{4})", plano, re.I)
    if m:
        try:
            achado["venc_texto"] = datetime.strptime(m.group(1), "%d/%m/%Y").date()
        except ValueError:
            pass
    return achado


# ---------------------------------------------------------------- e-mail

def remetente_de(msg):
    return (((msg.get("from") or {}).get("emailAddress") or {}).get("address") or "").lower()


def emails_com_pdf(headers, regras):
    desde = (HOJE - timedelta(days=DIAS_BUSCA)).isoformat() + "T00:00:00Z"
    url = (f"{GRAPH}/users/{CAIXA_EMAIL}/messages"
           f"?$filter=receivedDateTime ge {desde} and hasAttachments eq true"
           f"&$select=id,subject,from,receivedDateTime&$top=50")
    mensagens = []
    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code == 403:
            raise RuntimeError("Sem permissão para ler o e-mail: falta a permissão de aplicativo Mail.Read.")
        resp.raise_for_status()
        dados = resp.json()
        mensagens.extend(dados.get("value", []))
        url = dados.get("@odata.nextLink")
    remetentes = {r["remetente"] for r in regras}
    for m in sorted(mensagens, key=lambda x: x["receivedDateTime"]):
        de = remetente_de(m)
        if not any(r in de for r in remetentes):
            continue
        resp = requests.get(f"{GRAPH}/users/{CAIXA_EMAIL}/messages/{m['id']}/attachments", headers=headers)
        resp.raise_for_status()
        for anexo in resp.json().get("value", []):
            nome = anexo.get("name") or ""
            if anexo.get("@odata.type") == "#microsoft.graph.fileAttachment" and nome.lower().endswith(".pdf"):
                yield m, nome, base64.b64decode(anexo["contentBytes"])


def enviar_resumo(headers, linhas):
    corpo = "Boletos importados do e-mail\n\n" + "\n".join(linhas) + f"\n\n🔗 {LINK_DASHBOARD}"
    body = {"message": {
        "subject": "CAP - Boletos importados do e-mail",
        "body": {"contentType": "Text", "content": corpo},
        "toRecipients": [{"emailAddress": {"address": CAIXA_EMAIL}}],
    }, "saveToSentItems": "false"}
    resp = requests.post(f"{GRAPH}/users/{CAIXA_EMAIL}/sendMail",
                         headers={**headers, "Content-Type": "application/json"}, json=body)
    resp.raise_for_status()


# ---------------------------------------------------------------- planilha

def subir_boleto(headers, conteudo, desc, venc):
    nome = f"boleto-email-{chave_desc(desc).replace(' ', '-')}-{venc.strftime('%Y-%m')}.pdf"
    nome = re.sub(r"[^a-z0-9.\-]", "", nome)
    caminho = requests.utils.quote(f"{ANEXOS_PATH}/{nome}")
    resp = requests.put(f"{GRAPH}/drives/{DRIVE_ID}/root:/{caminho}:/content",
                        headers={**headers, "Content-Type": "application/pdf"}, data=conteudo)
    resp.raise_for_status()
    return resp.json().get("webUrl", "")


def gravar_linha(headers, index, valores):
    base = f"{GRAPH}/drives/{DRIVE_ID}/items/{ITEM_ID}/workbook/tables/TabelaLancamentos/rows"
    h = {**headers, "Content-Type": "application/json"}
    if index is None:
        resp = requests.post(base, headers=h, json={"values": [valores]})
    else:
        resp = requests.patch(f"{base}/itemAt(index={index})", headers=h, json={"values": [valores]})
    resp.raise_for_status()


def nome_da_linha(v):
    empresa = str(v[3] or "")
    return empresa.split(" - ")[0] if " - " in empresa else empresa


def desc_da_linha(v):
    full = str(v[7] or "")
    m = re.match(r"^(.*) - Parcela [\d/]+$", full)
    return m.group(1) if m else full


# ---------------------------------------------------------------- principal

def main():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    regras = carregar_regras(headers)
    if not regras:
        print("Nenhuma regra cadastrada (Boletos por e-mail / Neoenergia na Emissão). Nada a fazer.")
        return
    print(f"{len(regras)} regra(s):")
    for r in regras:
        print(f"  - {r['desc']} ({r['nome']}) <- {r['remetente']}" + (f" | texto: {r['texto']}" if r["texto"] else ""))

    extras = [s.strip() for s in os.environ.get("NEO_PDF_SENHAS", "").split(",") if s.strip()]
    lancamentos = ler_tabela(headers, "TabelaLancamentos")
    resumo, problemas = [], []

    for msg, nome_pdf, conteudo in emails_com_pdf(headers, regras):
        recebido = msg["receivedDateTime"][:10]
        de, assunto = remetente_de(msg), sem_acento(msg.get("subject"))
        candidatas = [r for r in regras if r["remetente"] in de and sem_acento(r["assunto"]) in assunto]
        if not candidatas:
            continue
        senhas = list(dict.fromkeys([s for r in candidatas for s in r["senhas"] if s] + extras))
        try:
            texto = texto_pdf(conteudo, senhas)
        except Exception as exc:
            problemas.append(f"⚠ {nome_pdf} de {de} ({recebido}): {exc}")
            continue
        digitos = so_digitos(texto)
        achadas = [r for r in candidatas if texto_bate(r["texto"], texto, digitos)]
        # Com e sem texto na mesma caixa: a que achou o texto ganha.
        com_texto = [r for r in achadas if r["texto"]]
        achadas = com_texto or achadas
        if len(achadas) != 1:
            problemas.append(f"⚠ {nome_pdf} de {de} ({recebido}): " +
                             ("nenhuma regra bateu com o PDF" if not achadas
                              else "mais de uma regra bateu — preencha 'Texto no PDF' para separar"))
            continue
        regra = achadas[0]
        b = ler_boleto(texto)
        valor = b.get("valor_barra") or b.get("valor_texto")
        venc = b.get("venc_texto") or b.get("venc_barra")
        if not venc or not valor:
            problemas.append(f"⚠ {regra['desc']} ({recebido}): não achei " +
                             " e ".join(x for x, ok in (("vencimento", venc), ("valor", valor)) if not ok))
            continue
        if b.get("valor_barra") and b.get("valor_texto") and abs(b["valor_barra"] - b["valor_texto"]) > 0.01:
            problemas.append(f"⚠ {regra['desc']}: valor do código de barras ({brl(b['valor_barra'])}) diferente "
                             f"do escrito ({brl(b['valor_texto'])}) — usei o do código de barras.")
        barra = b.get("barra", "")

        da_conta = [l for l in lancamentos
                    if nome_da_linha(l["values"]).strip().lower() == regra["nome"].strip().lower()
                    and parecida(desc_da_linha(l["values"]), regra["desc"])]
        if barra and any(so_digitos(l["values"][10]) == barra for l in da_conta):
            print(f"= {regra['desc']} {venc:%m/%Y}: já importado (mesmo código de barras).")
            continue
        do_mes = [l for l in da_conta
                  if (serial_para_data(l["values"][5]) or date.min).strftime("%Y-%m") == venc.strftime("%Y-%m")]
        # A descrição segue a do histórico, para somar junto nos gráficos.
        historico = sorted(da_conta, key=lambda l: serial_para_data(l["values"][5]) or date.min)
        desc_final = desc_da_linha(historico[-1]["values"]) if historico else regra["desc"]
        nota = f"Importado do e-mail ({recebido[8:10]}/{recebido[5:7]}/{recebido[:4]})"

        if do_mes:
            linha = do_mes[0]
            v = list(linha["values"]) + [""] * (17 - len(linha["values"]))
            if serial_para_data(v[6]) or str(v[9] or "").strip().lower() == "pago":
                print(f"= {regra['desc']} {venc:%m/%Y}: lançamento do mês já está pago — não mexi.")
                continue
            acao = f"✔ Completado: {desc_final} · {regra['nome']} · vence {venc:%d/%m/%Y} · {brl(valor)}"
            if not SIMULAR:
                url_boleto = subir_boleto(headers, conteudo, desc_final, venc)
                v[5], v[10], v[11] = data_para_serial(venc), barra, valor
                v[13] = url_boleto or v[13]
                v[16] = v[16] or nota
                gravar_linha(headers, linha["index"], v)
                linha["values"] = v
        else:
            acao = f"✚ Criado: {desc_final} · {regra['nome']} · vence {venc:%d/%m/%Y} · {brl(valor)}"
            if not SIMULAR:
                url_boleto = subir_boleto(headers, conteudo, desc_final, venc)
                novo_id = max([int(float(l["values"][0])) for l in lancamentos
                               if str(l["values"][0]).replace(".", "", 1).isdigit()] or [0]) + 1
                v = [novo_id, "", "", regra["nome"], "Envio de Boleto", data_para_serial(venc), "",
                     desc_final, regra["tipo"], "A Vencer", barra, valor, "", url_boleto, "", "", nota]
                gravar_linha(headers, None, v)
                lancamentos.append({"index": len(lancamentos), "values": v})
        resumo.append(acao + (" (simulação)" if SIMULAR else ""))
        print(acao)

    for p in problemas:
        print(p)
    if not resumo and not problemas:
        print("Nenhum boleto novo nos e-mails.")
        return
    if not SIMULAR:
        try:
            enviar_resumo(headers, resumo + ([""] if resumo and problemas else []) + problemas)
        except Exception as exc:
            print(f"Aviso: resumo por e-mail não enviado: {exc}")


if __name__ == "__main__":
    main()
