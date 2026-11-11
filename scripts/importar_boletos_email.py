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
  3. Lançamento já existente (mesmo código de barras, ou mesma pessoa + descrição
     parecida + vencimento no mesmo mês / a até 10 dias) — inclusive feito à mão:
       a vencer -> acerta vencimento, valor, código de barras e anexa o boleto
                   (só o que estiver diferente; se já estiver igual, não mexe);
       pago     -> não mexe em valor/data; só anexa o boleto se faltar.
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


def assunto_bate(alvo, assunto):
    """'Assunto contém' do cadastro x assunto do e-mail (já sem acento). Vale o trecho
    exato ou, se não, todas as palavras com 3+ letras em qualquer ordem: "A sua fatura
    Claro chegou." bate com "Sua Fatura Digital Claro chegou"."""
    alvo = sem_acento(alvo).strip()
    if not alvo or alvo in assunto:
        return True
    palavras = [w for w in re.findall(r"\w+", alvo) if len(w) >= 3]
    do_assunto = set(re.findall(r"\w+", assunto))
    return bool(palavras) and all(w in do_assunto for w in palavras)


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


MESES = ["janeiro", "fevereiro", "marco", "abril", "maio", "junho", "julho", "agosto",
         "setembro", "outubro", "novembro", "dezembro"]


def linha_digitavel(c):
    """Grava o código com pontos/espaços, como nos lançamentos manuais. Só dígitos,
    o Excel vira número (3.41918E+46) e perde o código."""
    if len(c) == 47:
        return f"{c[0:5]}.{c[5:10]} {c[10:15]}.{c[15:21]} {c[21:26]}.{c[26:32]} {c[32]} {c[33:]}"
    if len(c) == 48:
        return " ".join(f"{c[i:i + 11]}-{c[i + 11]}" for i in range(0, 48, 12))
    return c


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
    if achado.get("valor_texto") == 0 and "barra" not in achado:
        # Fatura em debito automatico: "TOTAL A PAGAR R$ 0,00" e sem codigo de barras.
        # O valor de verdade fica no "TOTAL" dos itens ou no corpo ("Valor: R$ 544.76").
        del achado["valor_texto"]
        achado["debito_auto"] = True
        m = re.search(r"(?<![A-Z])TOTAL\s+(?!A\s+PAGAR)R?\$?\s*([\d.]+,\d{2})", plano) \
            or re.search(r"VALOR:\s*R\$\s*(\d+)[.,](\d{2})(?!\d)", plano, re.I)
        if m:
            v = valor_br(m.group(1)) if m.lastindex == 1 else valor_br(m.group(1) + "," + m.group(2))
            if v > 0:
                achado["valor_texto"] = v
    m = re.search(r"VENC(?:IMENTO|E)[^\d]{0,40}?(\d{2}/\d{2}/\d{4})", plano, re.I)
    if m:
        try:
            achado["venc_texto"] = datetime.strptime(m.group(1), "%d/%m/%Y").date()
        except ValueError:
            pass
    if "venc_texto" not in achado:
        # "Vence em 16 de novembro" / "vencimento: 16 de novembro de 2026" (fatura de app,
        # sem boleto). Sem o ano, fica o mais proximo de hoje.
        m = re.search(r"VENC(?:IMENTO|E)(?:\s+EM|\s*:)?\s*(\d{1,2})\s+DE\s+([A-ZÇa-zç]+)(?:\s+DE\s+(\d{4}))?",
                      sem_acento(plano), re.I)
        if m and sem_acento(m.group(2)) in MESES:
            mes, dia = MESES.index(sem_acento(m.group(2))) + 1, int(m.group(1))
            anos = [int(m.group(3))] if m.group(3) else [HOJE.year - 1, HOJE.year, HOJE.year + 1]
            datas = []
            for a in anos:
                try:
                    datas.append(date(a, mes, dia))
                except ValueError:
                    pass
            if datas:
                achado["venc_texto"] = min(datas, key=lambda d: abs((d - HOJE).days))
    if "valor_texto" not in achado:
        # Valor solto perto do vencimento: "R$ 568,08" ou "R$ 568 08" (centavos em
        # sobrescrito viram outro pedaco de texto).
        m = re.search(r"VENC(?:IMENTO|E)[\s\S]{0,120}?R\$\s*([\d.]+)(?:\s*,\s*|\s+)(\d{2})(?!\d)", plano, re.I) \
            or re.search(r"R\$\s*([\d.]+)(?:\s*,\s*|\s+)(\d{2})(?!\d)[\s\S]{0,120}?VENC(?:IMENTO|E)", plano, re.I)
        if m:
            achado["valor_texto"] = valor_br(m.group(1) + "," + m.group(2))
    return achado


# ---------------------------------------------------------------- e-mail

def remetente_de(msg):
    return (((msg.get("from") or {}).get("emailAddress") or {}).get("address") or "").lower()


def html_para_texto(html):
    t = re.sub(r"(?is)<(script|style).*?</\1>", " ", html or "")
    t = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>|</td>", "\n", t)
    t = re.sub(r"<[^>]+>", " ", t)
    for a, b in (("&nbsp;", " "), ("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"), ("&#36;", "$")):
        t = t.replace(a, b)
    return re.sub(r"[ \t]+", " ", t)


def eh_encaminhado(msg):
    return re.match(r"^\s*(fwd?|enc|encaminhada?)\s*:", msg.get("subject") or "", re.I) is not None


def baixar_pdf_do_link(html):
    """Boleto que vem como link no corpo ("Baixar boleto"): tenta baixar o PDF."""
    links = re.findall(r"""href=["']([^"']+)["']""", html or "", re.I)
    for url in links:
        if not re.search(r"boleto|fatura|\.pdf|segunda.?via|2.?via", url, re.I):
            continue
        try:
            r = requests.get(url.replace("&amp;", "&"), timeout=30, allow_redirects=True)
        except Exception:
            continue
        if r.ok and (r.content[:5] == b"%PDF-" or "pdf" in r.headers.get("Content-Type", "").lower()):
            return r.content
    return None


def emails_com_boleto(headers, regras):
    """Para cada e-mail de remetente cadastrado: devolve (msg, origem, pdf, texto_corpo).
    Boleto em PDF anexo, em PDF por link, ou so no corpo do e-mail (linha digitavel).
    E-mail encaminhado tambem vale: o remetente cadastrado e procurado no corpo."""
    desde = (HOJE - timedelta(days=DIAS_BUSCA)).isoformat() + "T00:00:00Z"
    url = (f"{GRAPH}/users/{CAIXA_EMAIL}/messages"
           f"?$filter=receivedDateTime ge {desde}"
           f"&$select=id,subject,from,receivedDateTime,hasAttachments&$top=100")
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
        if not any(r in de for r in remetentes) and not eh_encaminhado(m):
            continue
        resp = requests.get(f"{GRAPH}/users/{CAIXA_EMAIL}/messages/{m['id']}?$select=body", headers=headers)
        resp.raise_for_status()
        html = (resp.json().get("body") or {}).get("content", "")
        corpo = html_para_texto(html)
        if not any(r in de for r in remetentes):
            # Encaminhado: so segue se um remetente cadastrado aparece no corpo.
            achou = [r for r in remetentes if r in corpo.lower()]
            if not achou:
                continue
            m["_remetente_original"] = achou[0]
        pdfs = []
        if m.get("hasAttachments"):
            resp = requests.get(f"{GRAPH}/users/{CAIXA_EMAIL}/messages/{m['id']}/attachments", headers=headers)
            resp.raise_for_status()
            for anexo in resp.json().get("value", []):
                nome = anexo.get("name") or ""
                if anexo.get("@odata.type") == "#microsoft.graph.fileAttachment" and nome.lower().endswith(".pdf"):
                    pdfs.append((nome, base64.b64decode(anexo["contentBytes"])))
        if not pdfs:
            pdf_link = baixar_pdf_do_link(html)
            if pdf_link:
                pdfs.append(("boleto (link do e-mail)", pdf_link))
        if pdfs:
            for nome, conteudo in pdfs:
                yield m, nome, conteudo, corpo
        else:
            yield m, "corpo do e-mail", None, corpo


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


# ---------------------------------------------------------------- lançamento previsto

# Marcas na Observação que o app mostra como selo ("aguardando boleto" / "erro no boleto").
MARCA_AGUARDANDO = "⏳ Aguardando boleto"
MARCA_ERRO = "⚠ Erro no boleto"


def pascoa(ano):
    a, b, c = ano % 19, ano // 100, ano % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = ((h + l - 7 * m + 114) % 31) + 1
    return date(ano, mes, dia)


def feriados(ano):
    """Feriados nacionais (bancos fechados): fixos + Carnaval, Sexta-feira Santa e Corpus Christi."""
    fixos = [(1, 1), (4, 21), (5, 1), (9, 7), (10, 12), (11, 2), (11, 15), (11, 20), (12, 25)]
    p = pascoa(ano)
    return {date(ano, m, d) for m, d in fixos} | {p - timedelta(days=48), p - timedelta(days=47),
                                                   p - timedelta(days=2), p + timedelta(days=60)}


def dia_util(d):
    """Vencimento em fim de semana ou feriado passa para o próximo dia útil."""
    while d.weekday() >= 5 or d in feriados(d.year):
        d += timedelta(days=1)
    return d


def dia_tipico(linhas):
    """Dia de vencimento de sempre, pelos últimos lançamentos. Datas empurradas por
    fim de semana/feriado aparecem menos, então fica o dia mais comum (no empate, o menor)."""
    datas = sorted([serial_para_data(l["values"][5]) for l in linhas if serial_para_data(l["values"][5])])[-6:]
    if not datas:
        return None
    contagem = {}
    for d in datas:
        contagem[d.day] = contagem.get(d.day, 0) + 1
    maior = max(contagem.values())
    return min(dia for dia, n in contagem.items() if n == maior)


def data_prevista(ano, mes, dia):
    ultimo = (date(ano + (mes == 12), mes % 12 + 1, 1) - timedelta(days=1)).day
    return dia_util(date(ano, mes, min(dia, ultimo)))


def chave_conta(nome, desc):
    return (str(nome or "").strip().lower(), chave_desc(desc))


def garantir_previstos(headers, regras, lancamentos, erros, resumo):
    """Todo boleto cadastrado tem lançamento no mês atual. Se o boleto não chegou
    (ou deu erro), cria o lançamento zerado com o vencimento de sempre e a marca na
    Observação; quando o boleto chegar, o robô completa esse mesmo lançamento."""
    inicio_mes = HOJE.replace(day=1)
    vistos = set()
    for regra in regras:
        k = chave_conta(regra["nome"], regra["desc"])
        if k in vistos:
            continue
        vistos.add(k)
        da_conta = [l for l in lancamentos
                    if nome_da_linha(l["values"]).strip().lower() == regra["nome"].strip().lower()
                    and parecida(desc_da_linha(l["values"]), regra["desc"])]
        do_mes = [l for l in da_conta
                  if (serial_para_data(l["values"][5]) or date.min).strftime("%Y-%m") == HOJE.strftime("%Y-%m")]
        erro = erros.get(k)
        if do_mes:
            # Já existe. Se é o previsto e o boleto deu erro, troca a marca para erro.
            linha = do_mes[0]
            v = list(linha["values"]) + [""] * (17 - len(linha["values"]))
            obs = str(v[16] or "")
            if erro and MARCA_AGUARDANDO in obs:
                v[16] = f"{MARCA_ERRO}: {erro}"
                resumo.append(f"⚠ Erro no boleto: {desc_da_linha(v)} · {regra['nome']} — {erro}")
                if not SIMULAR:
                    gravar_linha(headers, linha["index"], v)
                    linha["values"] = v
            continue
        # Só conta que segue ativa: teve lançamento em algum dos 3 meses anteriores.
        recentes = [l for l in da_conta
                    if inicio_mes - timedelta(days=95) <= (serial_para_data(l["values"][5]) or date.min) < inicio_mes]
        if not recentes:
            continue
        dia = dia_tipico(da_conta)
        if not dia:
            continue
        venc = data_prevista(HOJE.year, HOJE.month, dia)
        historico = sorted(da_conta, key=lambda l: serial_para_data(l["values"][5]) or date.min)
        ultimo = historico[-1]["values"]
        desc_final = desc_da_linha(ultimo)
        tipo = str(ultimo[8] or "").strip() or regra["tipo"]
        obs = (f"{MARCA_ERRO}: {erro}" if erro else
               f"{MARCA_AGUARDANDO} — previsto pelo histórico (vence todo dia {dia})")
        acao = (f"{'⚠' if erro else '⏳'} Previsto (boleto {'com erro' if erro else 'ainda não chegou'}): "
                f"{desc_final} · {regra['nome']} · vence {venc:%d/%m/%Y} · R$ 0,00")
        if not SIMULAR:
            novo_id = max([int(float(l["values"][0])) for l in lancamentos
                           if str(l["values"][0]).replace(".", "", 1).isdigit()] or [0]) + 1
            v = [novo_id, "", "", regra["nome"], "Envio de Boleto", data_para_serial(venc), "",
                 desc_final, tipo, "A Vencer", "", 0, "", "", "", "", obs]
            gravar_linha(headers, None, v)
            lancamentos.append({"index": len(lancamentos), "values": v})
        resumo.append(acao + (" (simulação)" if SIMULAR else ""))
        print(acao)


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
    erros = {}   # conta -> motivo, para marcar o lançamento previsto como "erro no boleto"

    for msg, nome_pdf, conteudo, corpo in emails_com_boleto(headers, regras):
        recebido = msg["receivedDateTime"][:10]
        de = msg.get("_remetente_original") or remetente_de(msg)
        assunto = sem_acento(msg.get("subject"))
        candidatas = [r for r in regras if r["remetente"] in de and assunto_bate(r["assunto"], assunto)]
        if not candidatas:
            continue
        senhas = list(dict.fromkeys([s for r in candidatas for s in r["senhas"] if s] + extras))
        if conteudo:
            try:
                texto = texto_pdf(conteudo, senhas)
            except Exception as exc:
                problemas.append(f"⚠ {nome_pdf} de {de} ({recebido}): {exc}")
                if len(candidatas) == 1:
                    erros[chave_conta(candidatas[0]["nome"], candidatas[0]["desc"])] = str(exc)
                continue
            # O corpo ajuda quando o PDF nao traz algum dado (ou vice-versa).
            texto = texto + "\n" + corpo
        else:
            texto = corpo
        if not conteudo:
            # Sem PDF: so vale se o corpo tiver cara de boleto (codigo de barras, ou
            # vencimento + valor). E-mail de aviso/cadastro ("Cadastro de Fatura
            # Digital!") e ignorado sem gerar alerta.
            pre = ler_boleto(texto)
            if not (pre.get("barra") or (pre.get("venc_texto") and pre.get("valor_texto"))):
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
            motivo = "não achei " + " e ".join(x for x, ok in (("vencimento", venc), ("valor", valor)) if not ok)
            problemas.append(f"⚠ {regra['desc']} ({recebido}): {motivo}")
            erros[chave_conta(regra["nome"], regra["desc"])] = f"{motivo} no e-mail de {recebido[8:10]}/{recebido[5:7]}"
            continue
        if b.get("valor_barra") and b.get("valor_texto") and abs(b["valor_barra"] - b["valor_texto"]) > 0.01:
            problemas.append(f"⚠ {regra['desc']}: valor do código de barras ({brl(b['valor_barra'])}) diferente "
                             f"do escrito ({brl(b['valor_texto'])}) — usei o do código de barras.")
        barra = b.get("barra", "")

        da_conta = [l for l in lancamentos
                    if nome_da_linha(l["values"]).strip().lower() == regra["nome"].strip().lower()
                    and parecida(desc_da_linha(l["values"]), regra["desc"])]
        # Lançamento que já existe para este boleto: o de mesmo código de barras, ou
        # o do mesmo mês de vencimento (ou a até 10 dias, quando foi lançado à mão com
        # a data aproximada). Entre vários, o que ainda não está pago e sem boleto.
        def pago(l):
            return bool(serial_para_data(l["values"][6])) or str(l["values"][9] or "").strip().lower() == "pago"

        mesma_barra = [l for l in da_conta if barra and so_digitos(l["values"][10]) == barra]
        mesmo_mes = [l for l in da_conta
                     if (serial_para_data(l["values"][5]) or date.min).strftime("%Y-%m") == venc.strftime("%Y-%m")]
        # Mês vizinho só vale para lançamento manual em aberto e sem código de barras
        # (ex.: lançado 31/10 e o boleto vence 02/11) — nunca o do mês anterior já pago.
        perto = [l for l in da_conta if not pago(l) and not so_digitos(l["values"][10])
                 and abs(((serial_para_data(l["values"][5]) or date.min) - venc).days) <= 10]
        do_mes = mesma_barra or mesmo_mes or perto
        do_mes.sort(key=lambda l: (pago(l), bool(so_digitos(l["values"][10])), bool(str(l["values"][13] or "").strip()),
                                   abs(((serial_para_data(l["values"][5]) or date.min) - venc).days)))
        # A descrição segue a do histórico, para somar junto nos gráficos.
        historico = sorted(da_conta, key=lambda l: serial_para_data(l["values"][5]) or date.min)
        desc_final = desc_da_linha(historico[-1]["values"]) if historico else regra["desc"]
        nota = f"Importado do e-mail ({recebido[8:10]}/{recebido[5:7]}/{recebido[:4]})"
        # "TOTAL A PAGAR R$ 0,00" sem codigo de barras: a fatura ja esta quitada
        # (debito automatico ou pago antes). O lancamento entra/fica como Pago, com a
        # data do e-mail como pagamento.
        quitada = bool(b.get("debito_auto"))
        data_pag = datetime.strptime(recebido, "%Y-%m-%d").date()
        if quitada:
            nota += " - fatura com total a pagar R$ 0,00 (já paga)"

        if do_mes:
            linha = do_mes[0]
            v = list(linha["values"]) + [""] * (17 - len(linha["values"]))
            sem_anexo = conteudo and not str(v[13] or "").strip()
            if pago(linha):
                # Já pago: não mexe em valor nem data; só guarda o boleto se faltar.
                if not sem_anexo and (so_digitos(v[10]) == barra or not barra):
                    print(f"= {desc_final} {venc:%m/%Y}: já pago e com boleto — não mexi.")
                    continue
                acao = f"📎 Boleto/código conferido (já pago): {desc_final} · {regra['nome']} · {venc:%m/%Y}"
                if not SIMULAR:
                    if sem_anexo:
                        v[13] = subir_boleto(headers, conteudo, desc_final, venc)
                    v[10] = linha_digitavel(barra) if barra else v[10]
                    gravar_linha(headers, linha["index"], v)
                    linha["values"] = v
            else:
                # A vencer (lançado à mão ou pelo robô): acerta o que estiver diferente.
                venc_antes = serial_para_data(v[5])
                try:
                    valor_antes = float(str(v[11]).replace(",", ".")) if str(v[11]).strip() else None
                except ValueError:
                    valor_antes = None
                mudou = []
                if valor_antes is None or abs(valor_antes - valor) > 0.009:
                    mudou.append(f"valor {brl(valor_antes) if valor_antes is not None else 'vazio'} → {brl(valor)}")
                if venc_antes != venc:
                    mudou.append(f"vencimento {venc_antes:%d/%m/%Y} → {venc:%d/%m/%Y}" if venc_antes
                                 else f"vencimento → {venc:%d/%m/%Y}")
                if barra and so_digitos(v[10]) != barra:
                    mudou.append("código de barras")
                if sem_anexo:
                    mudou.append("boleto anexado")
                if quitada:
                    mudou.append("marcado como pago")
                if not mudou:
                    print(f"= {desc_final} {venc:%m/%Y}: lançamento já está igual ao boleto.")
                    continue
                acao = (f"✔ Atualizado: {desc_final} · {regra['nome']} · vence {venc:%d/%m/%Y} · {brl(valor)}"
                        f" ({', '.join(mudou)})")
                if not SIMULAR:
                    if sem_anexo:
                        v[13] = subir_boleto(headers, conteudo, desc_final, venc)
                    v[5], v[11] = data_para_serial(venc), valor
                    v[10] = linha_digitavel(barra) if barra else v[10]
                    if not v[16] or MARCA_AGUARDANDO in str(v[16]) or MARCA_ERRO in str(v[16]):
                        v[16] = nota
                    if quitada:
                        v[6], v[9], v[12] = data_para_serial(data_pag), "Pago", valor
                        if "já paga" not in str(v[16]):
                            v[16] = (str(v[16]) + " - " if v[16] else "") + "fatura com total a pagar R$ 0,00 (já paga)"
                    gravar_linha(headers, linha["index"], v)
                    linha["values"] = v
        else:
            acao = (f"✚ Criado: {desc_final} · {regra['nome']} · vence {venc:%d/%m/%Y} · {brl(valor)}"
                    + (" · já pago" if quitada else ""))
            if not SIMULAR:
                url_boleto = subir_boleto(headers, conteudo, desc_final, venc) if conteudo else ""
                novo_id = max([int(float(l["values"][0])) for l in lancamentos
                               if str(l["values"][0]).replace(".", "", 1).isdigit()] or [0]) + 1
                v = [novo_id, "", "", regra["nome"], "Envio de Boleto", data_para_serial(venc),
                     data_para_serial(data_pag) if quitada else "", desc_final, regra["tipo"],
                     "Pago" if quitada else "A Vencer", linha_digitavel(barra), valor,
                     valor if quitada else "", url_boleto, "", "", nota]
                gravar_linha(headers, None, v)
                lancamentos.append({"index": len(lancamentos), "values": v})
        resumo.append(acao + (" (simulação)" if SIMULAR else ""))
        print(acao)

    garantir_previstos(headers, regras, lancamentos, erros, resumo)

    for p in problemas:
        print(p)
    if not resumo and not problemas:
        print("Nenhum boleto novo nos e-mails.")
        return
    # Roda de hora em hora: o que foi importado sai na hora (não se repete, pois o
    # lançamento já fica igual ao boleto); os avisos de problema se repetiriam a cada
    # rodada, então só vão por e-mail na rodada das 07:30 (AVISAR_PROBLEMAS=1).
    if not os.environ.get("AVISAR_PROBLEMAS"):
        problemas = []
    if not SIMULAR and (resumo or problemas):
        try:
            enviar_resumo(headers, resumo + ([""] if resumo and problemas else []) + problemas)
        except Exception as exc:
            print(f"Aviso: resumo por e-mail não enviado: {exc}")


if __name__ == "__main__":
    main()
