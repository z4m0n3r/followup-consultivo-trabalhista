"""
Follow Up Semanal — Consultivo Trabalhista — D'Artibale Faria Advogados
=======================================================================
Versão simplificada do follow up para a área trabalhista.

Diferença fundamental em relação ao cível: todo o fluxo de trabalho do
time trabalhista converge para o bucket de concluído. Não há necessidade
de classificar buckets, tratar pendências, impedimentos ou carryover de
tarefas não iniciadas. O script lê as tarefas concluídas na semana
anterior (segunda a sexta) e monta um resumo único.

Fluxo:
  1. Lê a planilha de configuração com clientes e IDs de grupo/plano.
  2. Autentica via client_credentials no Microsoft Graph.
  3. Resolve os IDs dos 2 integrantes do consultivo trabalhista.
  4. Para cada cliente, varre todos os planos do grupo.
  5. Filtra tarefas atribuídas ao time trabalhista E concluídas na
     semana anterior (por completedDateTime).
  6. Monta e envia o e-mail de follow up via SMTP.

Defensivo por design: retries para throttling e erros transitórios,
parada limpa se faltar permissão, e nunca envia dado errado.
"""

import os
import sys
import time
import json
import base64
import smtplib
import requests
import openpyxl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from datetime import datetime, timedelta, timezone

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("America/Sao_Paulo")
except Exception:
    TZ = timezone(timedelta(hours=-3))


# ─── Configurações ────────────────────────────────────────────────────────────

TENANT_ID     = os.environ["TENANT_ID"]
CLIENT_ID     = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
REMETENTE     = os.environ["REMETENTE"]
SMTP_PASSWORD = os.environ["SMTP_PASSWORD"]
SMTP_HOST     = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT     = int(os.environ.get("SMTP_PORT", "587"))
PLANILHA      = os.environ.get("PLANILHA_TRABALHISTA", "followup-consultivo-trabalhista.xlsx")
GRAPH_BASE    = "https://graph.microsoft.com/v1.0"

TEST_LIMIT    = int(os.environ.get("TEST_LIMIT", "0"))

# Destinatários via secret — suporta valor único ou lista separada por vírgula
_dest_raw = os.environ.get("DESTINATARIOS_TRABALHISTA", "")
DESTINATARIOS = [e.strip() for e in _dest_raw.split(",") if e.strip()]
if not DESTINATARIOS:
    print("[erro] DESTINATARIOS_TRABALHISTA não definido nos secrets.")
    sys.exit(2)

# Os 2 integrantes do consultivo trabalhista
CONSULTIVO_TRABALHISTA_EMAILS = [
    "karin.gambaro@dartibalefaria.com",
    "moisesmiguel.garcia@dartibalefaria.com",
]


# ─── Utilidades de texto ──────────────────────────────────────────────────────

def sem_acento(s: str) -> str:
    mapa = str.maketrans(
        "áàâãäéèêëíìîïóòôõöúùûüçÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇ",
        "aaaaaeeeeiiiiooooouuuucAAAAAEEEEIIIIOOOOOUUUUC",
    )
    return s.translate(mapa)

def norm(s: str) -> str:
    return sem_acento((s or "").strip().lower())


# ─── Autenticação ─────────────────────────────────────────────────────────────

def obter_token() -> str:
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
    }
    r = requests.post(url, data=data, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]


def roles_do_token(token: str) -> list:
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        claims = json.loads(base64.urlsafe_b64decode(payload))
        return claims.get("roles", [])
    except Exception:
        return []


# ─── Chamadas Graph com retry ─────────────────────────────────────────────────

def _req(token: str, url: str, tentativas: int = 5):
    h = {"Authorization": f"Bearer {token}"}
    for i in range(tentativas):
        r = requests.get(url, headers=h, timeout=30)
        if r.status_code == 429 or 500 <= r.status_code < 600:
            espera = int(r.headers.get("Retry-After", 2 ** i))
            time.sleep(min(espera, 30))
            continue
        return r
    return r

def get_all(token: str, path: str) -> list:
    items = []
    url = f"{GRAPH_BASE}{path}"
    while url:
        r = _req(token, url)
        if r.status_code == 404:
            break
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


# ─── Resolução dos IDs do consultivo trabalhista ──────────────────────────────

def resolver_ids_trabalhista(token: str, group_ids: list) -> dict:
    alvo = {e.lower() for e in CONSULTIVO_TRABALHISTA_EMAILS}
    encontrados = {}
    for gid in group_ids:
        if len(encontrados) == len(alvo):
            break
        membros = get_all(
            token,
            f"/groups/{gid}/members/microsoft.graph.user"
            f"?$select=id,mail,userPrincipalName&$top=999",
        )
        for m in membros:
            for campo in ("mail", "userPrincipalName"):
                val = (m.get(campo) or "").lower()
                if val in alvo and val not in encontrados:
                    encontrados[val] = m["id"]
    faltando = sorted(alvo - set(encontrados.keys()))
    if faltando:
        print("[aviso] Não encontrei nos grupos os seguintes integrantes do trabalhista:")
        for e in faltando:
            print(f"   - {e}")
    if not encontrados:
        print("[erro] Nenhum integrante do trabalhista foi resolvido. "
              "Execução interrompida.")
        sys.exit(2)
    return {uid: email for email, uid in encontrados.items()}


# ─── Datas (semana anterior, segunda a sexta) ─────────────────────────────────

def janela_semana_anterior() -> dict:
    hoje = datetime.now(TZ).date()
    seg_corrente = hoje - timedelta(days=hoje.weekday())
    seg_anterior = seg_corrente - timedelta(days=7)
    sex_anterior = seg_corrente - timedelta(days=3)
    return {"ini": seg_anterior, "fim": sex_anterior}

def data_campo(t: dict, campo: str):
    val = t.get(campo)
    if not val:
        return None
    try:
        return datetime.fromisoformat(val.replace("Z", "+00:00")).date()
    except Exception:
        return None

def entre(d, ini, fim) -> bool:
    return d is not None and ini <= d <= fim


# ─── Buckets ignorados (ritual interno, backlog, modelos) ─────────────────────

BUCKETS_IGNORADOS = ("modelo", "backlog", "bloque")

def bucket_ignorado(nome: str) -> bool:
    n = norm(nome)
    return any(x in n for x in BUCKETS_IGNORADOS)


# ─── Títulos de ritual interno (nunca entram no follow up) ─────────────────────

TITULOS_EXCLUIDOS = (
    "follow-up semanal", "follow up semanal",
    "reuniao mensal", "reunião mensal",
    "follow-up semanal | consultivo trabalhista",
    "follow up consultivo trabalhista",
    "follow-up | consultivo trabalhista",
    "follow up - consultivo trabalhista",
)

def titulo_excluido(t: dict) -> bool:
    titulo = norm(t.get("title") or "")
    return any(x in titulo for x in (norm(e) for e in TITULOS_EXCLUIDOS))


# ─── Coleta de planos e tarefas ───────────────────────────────────────────────

def planos_do_grupo(token: str, group_id: str) -> list:
    planos = get_all(token, f"/groups/{group_id}/planner/plans")
    return [p["id"] for p in planos]

def buckets_do_plano(token: str, plan_id: str) -> dict:
    return {b["id"]: b["name"] for b in get_all(token, f"/planner/plans/{plan_id}/buckets")}

def tarefas_do_plano(token: str, plan_id: str) -> list:
    return get_all(token, f"/planner/plans/{plan_id}/tasks")


# ─── Filtro por responsável ───────────────────────────────────────────────────

def e_do_trabalhista(t: dict, ids_trabalhista: set) -> bool:
    atrib = t.get("assignments") or {}
    return any(uid in ids_trabalhista for uid in atrib.keys())


# ─── Montagem do follow up ────────────────────────────────────────────────────

def montar_followup(token, cliente, group_id, ids_trabalhista, jan, plan_id_fixo=None) -> str:
    if plan_id_fixo:
        plan_ids = [plan_id_fixo]
    else:
        plan_ids = planos_do_grupo(token, group_id)
    if not plan_ids:
        return None

    concluidas = []

    for plan_id in plan_ids:
        buckets = buckets_do_plano(token, plan_id)
        for t in tarefas_do_plano(token, plan_id):
            if not e_do_trabalhista(t, ids_trabalhista):
                continue
            if titulo_excluido(t):
                continue
            nome_bucket = buckets.get(t.get("bucketId", ""), "")
            if bucket_ignorado(nome_bucket):
                continue
            feito = data_campo(t, "completedDateTime")
            if entre(feito, jan["ini"], jan["fim"]):
                concluidas.append(t)

    return formatar_email(cliente, jan, concluidas)


def formatar_email(cliente, jan, concluidas):
    ini = jan["ini"].strftime("%d/%m/%Y")
    fim = jan["fim"].strftime("%d/%m/%Y")
    L = []
    L.append("[FOLLOW UP SEMANAL — CONSULTIVO TRABALHISTA]")
    L.append("")
    L.append("Boa tarde.")
    L.append("")
    L.append(
        f"Segue abaixo o resumo das demandas concluídas pelo time "
        f"D'Artibale Faria — Consultivo Trabalhista, referente à semana "
        f"de {ini} a {fim}."
    )
    L.append("")
    L.append("Demandas Concluídas na Semana")
    if not concluidas:
        L.append("  Sem demandas concluídas nesta semana.")
    else:
        for i, t in enumerate(concluidas):
            titulo = (t.get("title") or "(sem título)").strip()
            feito = data_campo(t, "completedDateTime")
            data_str = f" — {feito.strftime('%d/%m/%Y')}" if feito else ""
            marcador = chr(ord("a") + i) + ")" if i < 26 else "•"
            L.append(f"  {marcador} {titulo}{data_str}")
    L.append("")
    return "\n".join(L)


# ─── Envio por SMTP ───────────────────────────────────────────────────────────

def enviar_email(assunto: str, corpo: str):
    msg = MIMEMultipart()
    msg["From"]    = REMETENTE
    msg["To"]      = ", ".join(DESTINATARIOS)
    msg["Subject"] = Header(assunto, "utf-8")
    msg.attach(MIMEText(corpo, "plain", "utf-8"))
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60) as s:
        s.starttls()
        s.login(REMETENTE, SMTP_PASSWORD)
        s.sendmail(REMETENTE, DESTINATARIOS, msg.as_string())


# ─── Leitura da planilha ──────────────────────────────────────────────────────

def ler_clientes(caminho: str) -> list:
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb["Configuração"]
    clientes, cab = [], True
    for row in ws.iter_rows(values_only=True):
        if cab:
            cab = False
            continue
        cliente, planner_id, _emails, group_id = (list(row) + [None, None, None, None])[:4]
        if not cliente:
            continue
        gid = str(group_id).strip() if group_id else ""
        pid = str(planner_id).strip() if planner_id else ""
        if not gid:
            gid = pid
        if not gid:
            continue
        plan_id_fixo = pid if (pid and pid != gid) else None
        clientes.append({
            "cliente": str(cliente).strip(),
            "group_id": gid,
            "plan_id_fixo": plan_id_fixo,
        })
    wb.close()
    return clientes


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "=" * 60)
    print("  Follow Up Semanal — Consultivo Trabalhista")
    print(f"  {datetime.now(TZ).strftime('%d/%m/%Y %H:%M')}")
    print("=" * 60 + "\n")

    token = obter_token()
    print(f"Permissões do app: {roles_do_token(token) or 'NENHUMA'}\n")

    jan = janela_semana_anterior()
    print(f"Semana anterior: {jan['ini']:%d/%m} a {jan['fim']:%d/%m}\n")

    clientes = ler_clientes(PLANILHA)
    if TEST_LIMIT > 0:
        clientes = clientes[:TEST_LIMIT]
        print(f"[teste] MODO TESTE — apenas {len(clientes)} cliente(s).\n")
    print(f"[ok] {len(clientes)} clientes carregados.\n")

    todos_grupos = [c["group_id"] for c in ler_clientes(PLANILHA)]
    ids_trabalhista = set(resolver_ids_trabalhista(token, todos_grupos).keys())
    print(f"[ok] {len(ids_trabalhista)} integrantes do trabalhista resolvidos.\n")

    enviados, vazios, erros = 0, 0, []

    for i, c in enumerate(clientes, 1):
        print(f"[{i:02d}/{len(clientes)}] {c['cliente']}")
        try:
            corpo = montar_followup(
                token, c["cliente"], c["group_id"],
                ids_trabalhista, jan, c.get("plan_id_fixo"),
            )
            if corpo is None:
                print("   sem plano de Planner no grupo — pulado.")
                vazios += 1
                continue
            assunto = (
                f"[Follow Up Semanal — Trabalhista] "
                f"{c['cliente']} — {datetime.now(TZ):%d/%m/%Y}"
            )
            enviar_email(assunto, corpo)
            print(f"   [ok] enviado para {', '.join(DESTINATARIOS)}")
            enviados += 1
            time.sleep(2)
        except Exception as e:
            print(f"   [erro] {type(e).__name__}: {str(e)[:200]}")
            erros.append((c["cliente"], str(e)[:200]))

    print("\n" + "=" * 60)
    print(f"  Enviados: {enviados}  |  Sem plano: {vazios}  |  Erros: {len(erros)}")
    for nome, err in erros:
        print(f"   - {nome}: {err}")
    print("=" * 60 + "\n")

    if erros:
        sys.exit(1)


if __name__ == "__main__":
    main()