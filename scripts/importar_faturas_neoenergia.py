"""
Importa as faturas da Neoenergia que chegam por e-mail (Fatura Digital) e cria
ou completa o lançamento correspondente no CAP.

Fluxo (1x por dia via GitHub Actions, .github/workflows/importar-faturas-neoenergia.yml):
  1. Lê na planilha financeira a TabelaEmissao (tela "Emissão de Documentos" do
     CAP): cada linha com link da Neoenergia e Código preenchido é uma unidade
     consumidora — descrição, "Despesa de" e código vêm de lá, nada fica escrito
     neste arquivo (o repositório é público).
  2. Procura na caixa do REMETENTE_EMAIL os e-mails da Neoenergia dos últimos
     DIAS_BUSCA dias com PDF anexo.
  3. Em cada PDF: acha a unidade pelo código, o código de barras (48 dígitos,
     de onde também sai o valor), o vencimento e o valor escrito.
  4. Lançamento do mesmo mês já existe (mesma pessoa, descrição parecida):
       - pago, ou com o mesmo código de barras -> não mexe;
       - a vencer -> completa vencimento, valor, código de barras e boleto.
     Não existe -> cria "A Vencer" com o boleto anexado.
  5. Manda um resumo por e-mail ao REMETENTE_EMAIL quando importou algo ou
     quando não conseguiu ler alguma fatura.

Simular (não grava nada, só mostra o que faria): SIMULAR=1.

Precisa no App Registration "Gestão Financeira Claude" (permissões de APLICATIVO,
com consentimento do administrador): Mail.Read (novo), Mail.Send e
Files.ReadWrite.All (já usadas pelos outros scripts).

Variáveis de ambiente:
  CAP_CLIENT_SECRET  (obrigatória)
  NEO_PDF_SENHAS     (opcional) senhas extras para PDF protegido, separadas por vírgula
  SIMULAR            (opcional) "1" para só simular
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
REMETENTE_EMAIL = "natanael.silva@sinerggia.com.br"
LINK_DASHBOARD = "https://sinerggia-dev.github.io/cap-contas-a-pagar/"
DIAS_BUSCA = 40
SIMULAR = os.environ.get("SIMULAR", "").strip() in ("1", "sim", "true", "yes")

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


def ler_tabela(headers, nome_tabela):
    url = f"{GRAPH}/drives/{DRIVE_ID}/items/{ITEM_ID}/workbook/tables/{nome_tabela}/rows"
    linhas = []
    while url:
        resp = requests.get(url, headers=headers)
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
    Os números têm de ser iguais: "Apto 102" e "Apto 201" são parecidos no texto,
    mas são unidades diferentes."""
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


# ---------------------------------------------------------------- unidades

def carregar_unidades(headers):
    """TabelaEmissao: [Id, Descrição, Despesa de, Tipo, Link, Usuário, Senha, Código, Dia, Nota, Atualizado]."""
    unidades = []
    for linha in ler_tabela(headers, "TabelaEmissao"):
        v = linha["values"]
        desc, nome, link, usuario, codigo = (str(v[i] or "").strip() for i in (1, 2, 4, 5, 7))
        if "neoenergia" not in (link + " " + desc).lower() or not so_digitos(codigo):
            continue
        unidades.append({
            "desc": desc, "nome": nome, "codigo": so_digitos(codigo).lstrip("0"),
            "doc_usuario": so_digitos(usuario),
        })
    return unidades


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


RE_LINHA_DIGITAVEL = re.compile(
    r"(8\d{10})\s*-?\s*(\d)\s*(\d{11})\s*-?\s*(\d)\s*(\d{11})\s*-?\s*(\d)\s*(\d{11})\s*-?\s*(\d)"
)


def ler_fatura(texto):
    """Devolve dict com barra, valor, venc (os que achar)."""
    achado = {}
    m = RE_LINHA_DIGITAVEL.search(texto)
    if m:
        g = m.groups()
        achado["barra"] = "".join(g)
        codigo44 = g[0] + g[2] + g[4] + g[6]
        # Arrecadação: 3º dígito 6 ou 8 = valor efetivo em centavos nas posições 5-15.
        if codigo44[2] in "68":
            achado["valor_barra"] = int(codigo44[4:15]) / 100

    plano = re.sub(r"[ \t]+", " ", texto)
    for padrao in (
        r"TOTAL\s+A\s+PAGAR[^\d]{0,40}?([\d.]+,\d{2})",
        r"VALOR\s+A\s+PAGAR[^\d]{0,40}?([\d.]+,\d{2})",
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
            achado["venc"] = datetime.strptime(m.group(1), "%d/%m/%Y").date()
        except ValueError:
            pass
    return achado


def unidade_da_fatura(texto, unidades):
    digitos = so_digitos(texto)
    return [u for u in unidades if u["codigo"] and u["codigo"] in digitos]


# ---------------------------------------------------------------- e-mail

def faturas_no_email(headers):
    desde = (HOJE - timedelta(days=DIAS_BUSCA)).isoformat() + "T00:00:00Z"
    url = (f"{GRAPH}/users/{REMETENTE_EMAIL}/messages"
           f"?$filter=receivedDateTime ge {desde} and hasAttachments eq true"
           f"&$select=id,subject,from,receivedDateTime&$top=50")
    mensagens = []
    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code == 403:
            raise RuntimeError("Sem permissão para ler o e-mail: falta a permissão de aplicativo Mail.Read "
                               "(com consentimento do administrador) no App Registration.")
        resp.raise_for_status()
        dados = resp.json()
        mensagens.extend(dados.get("value", []))
        url = dados.get("@odata.nextLink")
    neo = [m for m in mensagens
           if "neoenergia" in ((m.get("from") or {}).get("emailAddress") or {}).get("address", "").lower()]
    for m in sorted(neo, key=lambda x: x["receivedDateTime"]):
        resp = requests.get(f"{GRAPH}/users/{REMETENTE_EMAIL}/messages/{m['id']}/attachments", headers=headers)
        resp.raise_for_status()
        for anexo in resp.json().get("value", []):
            nome = anexo.get("name") or ""
            if anexo.get("@odata.type") == "#microsoft.graph.fileAttachment" and nome.lower().endswith(".pdf"):
                yield m, nome, base64.b64decode(anexo["contentBytes"])


def enviar_resumo(headers, linhas):
    corpo = "Importação de faturas da Neoenergia\n\n" + "\n".join(linhas) + f"\n\n🔗 {LINK_DASHBOARD}"
    body = {"message": {
        "subject": "CAP - Faturas Neoenergia importadas",
        "body": {"contentType": "Text", "content": corpo},
        "toRecipients": [{"emailAddress": {"address": REMETENTE_EMAIL}}],
    }, "saveToSentItems": "false"}
    resp = requests.post(f"{GRAPH}/users/{REMETENTE_EMAIL}/sendMail",
                         headers={**headers, "Content-Type": "application/json"}, json=body)
    resp.raise_for_status()


# ---------------------------------------------------------------- planilha

def subir_boleto(headers, conteudo, desc, venc):
    nome = f"boleto-neoenergia-{chave_desc(desc).replace(' ', '-')}-{venc.strftime('%Y-%m')}.pdf"
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
    unidades = carregar_unidades(headers)
    if not unidades:
        print("Nenhuma unidade da Neoenergia com código na Emissão de Documentos. Nada a fazer.")
        return
    print(f"{len(unidades)} unidade(s): " + ", ".join(f"{u['desc']} ({u['nome']})" for u in unidades))

    senhas = [s.strip() for s in os.environ.get("NEO_PDF_SENHAS", "").split(",") if s.strip()]
    for u in unidades:
        d = u["doc_usuario"]
        if d:
            senhas += [d[:5], d[:6], d[:4], d]
    senhas = list(dict.fromkeys(senhas))

    lancamentos = ler_tabela(headers, "TabelaLancamentos")
    resumo, problemas = [], []

    for msg, nome_pdf, conteudo in faturas_no_email(headers):
        recebido = msg["receivedDateTime"][:10]
        try:
            texto = texto_pdf(conteudo, senhas)
        except Exception as exc:
            problemas.append(f"⚠ {nome_pdf} ({recebido}): {exc}")
            continue
        achadas = unidade_da_fatura(texto, unidades)
        if len(achadas) != 1:
            problemas.append(f"⚠ {nome_pdf} ({recebido}): " +
                             ("nenhuma unidade conhecida no PDF" if not achadas else "mais de uma unidade no PDF"))
            continue
        un = achadas[0]
        f = ler_fatura(texto)
        valor = f.get("valor_barra") or f.get("valor_texto")
        if not f.get("venc") or not valor:
            problemas.append(f"⚠ {un['desc']} ({recebido}): não achei " +
                             " e ".join(x for x, ok in (("vencimento", f.get("venc")), ("valor", valor)) if not ok))
            continue
        if f.get("valor_barra") and f.get("valor_texto") and abs(f["valor_barra"] - f["valor_texto"]) > 0.01:
            problemas.append(f"⚠ {un['desc']}: valor do código de barras ({brl(f['valor_barra'])}) diferente "
                             f"do escrito ({brl(f['valor_texto'])}) — usei o do código de barras.")
        venc = f["venc"]
        barra = f.get("barra", "")

        # Lançamentos da mesma pessoa e descrição parecida.
        da_unidade = [l for l in lancamentos
                      if nome_da_linha(l["values"]).strip().lower() == un["nome"].strip().lower()
                      and parecida(desc_da_linha(l["values"]), un["desc"])]
        if barra and any(so_digitos(l["values"][10]) == barra for l in da_unidade):
            print(f"= {un['desc']} {venc:%m/%Y}: já importada (mesmo código de barras).")
            continue
        do_mes = [l for l in da_unidade
                  if (serial_para_data(l["values"][5]) or date.min).strftime("%Y-%m") == venc.strftime("%Y-%m")]
        # A descrição continua a do histórico (ex.: grafia antiga), para somar junto nos gráficos.
        historico = sorted(da_unidade, key=lambda l: serial_para_data(l["values"][5]) or date.min)
        desc_final = desc_da_linha(historico[-1]["values"]) if historico else un["desc"]
        nota = f"Importado da fatura digital Neoenergia ({recebido[8:10]}/{recebido[5:7]}/{recebido[:4]})"

        if do_mes:
            linha = do_mes[0]
            v = list(linha["values"])
            pago = serial_para_data(v[6]) or str(v[9] or "").strip().lower() == "pago"
            if pago:
                print(f"= {un['desc']} {venc:%m/%Y}: lançamento do mês já está pago — não mexi.")
                continue
            acao = f"✔ Completado: {desc_final} · {un['nome']} · vence {venc:%d/%m/%Y} · {brl(valor)}"
            if not SIMULAR:
                url_boleto = subir_boleto(headers, conteudo, desc_final, venc)
                v[5] = data_para_serial(venc)
                v[10] = barra
                v[11] = valor
                v[13] = url_boleto or v[13]
                v[16] = v[16] or nota
                gravar_linha(headers, linha["index"], v)
                linha["values"] = v
        else:
            acao = f"✚ Criado: {desc_final} · {un['nome']} · vence {venc:%d/%m/%Y} · {brl(valor)}"
            if not SIMULAR:
                url_boleto = subir_boleto(headers, conteudo, desc_final, venc)
                novo_id = max([int(float(l["values"][0])) for l in lancamentos
                               if str(l["values"][0]).replace(".", "", 1).isdigit()] or [0]) + 1
                v = [novo_id, "", "", un["nome"], "Envio de Boleto", data_para_serial(venc), "",
                     desc_final, "Energia", "A Vencer", barra, valor, "", url_boleto, "", "", nota]
                gravar_linha(headers, None, v)
                lancamentos.append({"index": len(lancamentos), "values": v})
        resumo.append(acao + (" (simulação)" if SIMULAR else ""))
        print(acao)

    for p in problemas:
        print(p)
    if not resumo and not problemas:
        print("Nenhuma fatura nova da Neoenergia.")
        return
    if not SIMULAR:
        try:
            enviar_resumo(headers, resumo + ([""] if resumo and problemas else []) + problemas)
        except Exception as exc:
            print(f"Aviso: resumo por e-mail não enviado: {exc}")


if __name__ == "__main__":
    main()
