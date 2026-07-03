"""
Follow Up Semanal — D'Artibale Faria Advogados
==============================================
Para cada um dos clientes (linhas da planilha):
  1. Resolve o(s) plano(s) de Planner a partir do ID do grupo.
  2. Varre TODOS os planos do grupo (um grupo pode ter vários).
  3. Mantém apenas tarefas atribuídas às 8 pessoas do CONSULTIVO.
  4. Classifica por bucket + data e monta o follow up no formato padrão.
  5. Envia por SMTP (Gmail) APENAS para Rafael e Maria Clara.

Defensivo por design: retries para throttling/erros de rede, classificação
de bucket tolerante a variação de nome, e parada limpa com mensagem clara
se faltar permissão — nunca envia dado errado.

CORREÇÃO 2026-06: tarefas NÃO INICIADAS que venceram na SEMANA ANTERIOR e não
foram concluídas deixavam de aparecer em qualquer seção (limbo). Agora são
arrastadas para "Previstas para Esta Semana", marcadas como pendentes.
"""

import os
import re
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
PLANILHA      = os.environ.get("PLANILHA", "followup.xlsx")
GRAPH_BASE    = "https://graph.microsoft.com/v1.0"

# Limite opcional para teste (ex.: TEST_LIMIT=1 processa só o 1º cliente)
TEST_LIMIT    = int(os.environ.get("TEST_LIMIT", "0"))

# Destinatários FIXOS — a coluna de e-mails da planilha é ignorada de propósito
DESTINATARIOS = [
    "karin.gambaro@dartibalefaria.com",
    "moisesmiguel.garcia@dartibalefaria.com",
]

# Os integrantes do CONSULTIVO — o filtro mantém só tarefas destas pessoas
CONSULTIVO_EMAILS = [
    "karin.gambaro@dartibalefaria.com",
    "moisesmiguel.garcia@dartibalefaria.com",
]


# ─── Utilidades de texto ──────────────────────────────────────────────────────

def sem_acento(s: str) -> str:
    mapa = str.maketrans("áàâãäéèêëíìîïóòôõöúùûüçÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇ",
                         "aaaaaeeeeiiiiooooouuuucAAAAAEEEEIIIIOOOOOUUUUC")
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


# ─── Chamadas Graph com retry (trata 429 e erros transitórios) ────────────────

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

def get_one(token: str, path: str) -> dict:
    r = _req(token, f"{GRAPH_BASE}{path}")
    if r.status_code == 404:
        return {}
    r.raise_for_status()
    return r.json()

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


# ─── Resolução dos IDs do consultivo ──────────────────────────────────────────

def resolver_ids_consultivo(token: str, group_ids: list) -> dict:
    """Monta {user_id: email} das 8 pessoas do consultivo lendo os MEMBROS
    dos grupos (usa Group.Read.All, não depende de User.Read.All).
    Varre os grupos acumulando o mapa e para quando achar todos os 8."""
    alvo = {e.lower() for e in CONSULTIVO_EMAILS}
    encontrados = {}  # email_lower -> user_id
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
        print("[aviso] Não encontrei nos grupos os seguintes integrantes do consultivo:")
        for e in faltando:
            print(f"   - {e}")
        print("   (As tarefas dessas pessoas não entrarão no filtro.)")
    if not encontrados:
        print("[erro] Nenhum integrante do consultivo foi resolvido. "
              "Execução interrompida para não enviar follow up sem filtro.")
        sys.exit(2)
    # retorna {user_id: email}
    return {uid: email for email, uid in encontrados.items()}


# ─── Datas (semana anterior e corrente, segunda a sexta) ──────────────────────

def janelas_semana() -> dict:
    hoje = datetime.now(TZ).date()
    seg_corrente = hoje - timedelta(days=hoje.weekday())
    sex_corrente = seg_corrente + timedelta(days=4)
    seg_anterior = seg_corrente - timedelta(days=7)
    sex_anterior = seg_corrente - timedelta(days=3)
    return {
        "ini_ant": seg_anterior, "fim_ant": sex_anterior,
        "ini_cur": seg_corrente, "fim_cur": sex_corrente,
    }

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


# ─── Classificação de bucket tolerante a variação de nome ─────────────────────
# Categorias: concluido | nao_iniciado | em_andamento | validacao_gestor
#             assinatura | pendencia | outro

# Buckets explicitamente ignorados (não entram em nenhuma seção)
BUCKETS_IGNORADOS = ("modelo", "backlog", "bloque")

EM_ANDAMENTO_BUCKETS = ("andamento", "em andamento", "doing", "em progresso")

def categoria_bucket(nome: str) -> str:
    n = norm(nome)
    if any(x in n for x in BUCKETS_IGNORADOS):
        return "ignorar"
    if "assinatura" in n:
        return "assinatura"
    if "conclu" in n:
        return "concluido"
    if any(x in n for x in EM_ANDAMENTO_BUCKETS):
        return "em_andamento"
    if "valida" in n and "cliente" in n:
        return "pendencia"          # validação com o cliente
    if "pend" in n:
        return "pendencia"          # com pendência
    return "outro"


def categoria_status(t: dict) -> str:
    status = (t.get("status") or "").strip().lower()
    if status in ("inprogress", "in progress", "doing", "andamento", "em andamento", "em progresso"):
        return "em_andamento"
    if status in ("completed", "done", "concluido", "concluida", "finalizado", "finished"):
        return "concluido"
    if status in ("notstarted", "not started", "to do", "todo", "nao iniciado", "não iniciado", "a fazer", "pending"):
        return "nao_iniciado"

    percent = t.get("percentComplete")
    try:
        if percent is not None:
            percent = int(percent)
            if percent == 100:
                return "concluido"
            if 0 < percent < 100:
                return "em_andamento"
            return "nao_iniciado"
    except Exception:
        pass

    if t.get("completedDateTime"):
        return "concluido"
    return "outro"


# ─── Coleta de planos, buckets e tarefas do grupo ─────────────────────────────

def planos_do_grupo(token: str, group_id: str) -> list:
    planos = get_all(token, f"/groups/{group_id}/planner/plans")
    return [p["id"] for p in planos]

def buckets_do_plano(token: str, plan_id: str) -> dict:
    return {b["id"]: b["name"] for b in get_all(token, f"/planner/plans/{plan_id}/buckets")}

def tarefas_do_plano(token: str, plan_id: str) -> list:
    return get_all(token, f"/planner/plans/{plan_id}/tasks")


# ─── Filtro por responsável ───────────────────────────────────────────────────

def e_do_consultivo(t: dict, ids_consultivo: set) -> bool:
    atrib = t.get("assignments") or {}
    return any(uid in ids_consultivo for uid in atrib.keys())


# ─── Montagem do follow up de um cliente ──────────────────────────────────────

# Títulos que são ritual interno (nunca entram no follow up)
TITULOS_EXCLUIDOS = ("follow-up semanal", "follow up semanal", "reuniao mensal", "reunião mensal", "Follow-up Semanal | Consutivo Cível", "Follow Up Consultivo Cível", "Follow-up | Consultivo Empresarial", "Follow Up - Consultivo Cível")

def titulo_excluido(t: dict) -> bool:
    titulo = norm(t.get("title") or "")
    return any(x in titulo for x in (norm(e) for e in TITULOS_EXCLUIDOS))

def montar_followup(token, cliente, group_id, ids_consultivo, jan, plan_id_fixo=None) -> str:
    if plan_id_fixo:
        plan_ids = [plan_id_fixo]
    else:
        plan_ids = planos_do_grupo(token, group_id)
    if not plan_ids:
        return None  # nada a enviar

    resumo, previstas, assinatura, impedimentos = [], [], [], []

    for plan_id in plan_ids:
        buckets = buckets_do_plano(token, plan_id)
        for t in tarefas_do_plano(token, plan_id):
            if not e_do_consultivo(t, ids_consultivo):
                continue
            if titulo_excluido(t):
                continue
            cat = categoria_bucket(buckets.get(t.get("bucketId", ""), ""))
            if cat == "ignorar":
                continue
            venc = data_campo(t, "dueDateTime")
            feito = data_campo(t, "completedDateTime")
            status = categoria_status(t)

            # Pendente de Assinatura — por bucket, todas as datas
            if cat == "assinatura":
                assinatura.append((t, plan_id))
                continue

            # Impedimentos — pendência explícita ou status em andamento sem bucket de andamento
            if cat == "pendencia" or \
               (status == "em_andamento" and cat != "em_andamento") or \
               (cat == "em_andamento" and status != "em_andamento"):
                impedimentos.append((t, plan_id))
                continue

            # Resumo da última semana — concluídas na semana anterior (por data)
            if entre(feito, jan["ini_ant"], jan["fim_ant"]):
                resumo.append((t, plan_id))
            # Previstas para esta semana — vencimento na semana corrente (por data)
            elif entre(venc, jan["ini_cur"], jan["fim_cur"]):
                previstas.append((t, plan_id))
            # CORREÇÃO: carryover. Não iniciadas que venceram na SEMANA ANTERIOR e
            # não foram concluídas. Sem este ramo, caíam no limbo. Agora entram nas
            # previstas, marcadas como pendentes da semana passada.
            elif status == "nao_iniciado" and entre(venc, jan["ini_ant"], jan["fim_ant"]):
                t["_atrasada"] = True
                previstas.append((t, plan_id))
            # Opcional (descomente para resgatar QUALQUER atrasada não iniciada,
            # inclusive de semanas mais antigas, não só a imediatamente anterior):
            # elif status == "nao_iniciado" and venc is not None and venc < jan["ini_cur"]:
            #     t["_atrasada"] = True
            #     previstas.append((t, plan_id))

    return formatar_email(token, cliente, group_id, jan,
                          resumo, previstas, assinatura, impedimentos)


def linha_tarefa(t):
    titulo = (t.get("title") or "(sem título)").strip()
    venc = data_campo(t, "dueDateTime")
    sufixo = f" — {venc.strftime('%d/%m/%Y')}" if venc else ""
    if t.get("_atrasada"):
        sufixo += " (pendente da semana anterior — não iniciada)"
    return f"{titulo}{sufixo}"


def formatar_email(token, cliente, group_id, jan,
                   resumo, previstas, assinatura, impedimentos):
    ini = jan["ini_ant"].strftime("%d/%m/%Y")
    fim = jan["fim_ant"].strftime("%d/%m/%Y")
    L = []
    L.append("[FOLLOW UP SEMANAL]")
    L.append("")
    L.append("Boa tarde.")
    L.append("")
    L.append(f"Segue abaixo o resumo semanal das demandas tratadas pelo time "
             f"D'Artibale Faria, referente à semana anterior de {ini} a {fim}.")
    L.append("")

    def bloco(titulo, itens):
        L.append(titulo)
        if not itens:
            L.append("  Sem demandas nesta categoria.")
        else:
            for i, (t, _pid) in enumerate(itens):
                marcador = chr(ord('a') + i) + ")" if i < 26 else "•"
                L.append(f"  {marcador} {linha_tarefa(t)}")
        L.append("")

    bloco("Resumo das Principais Demandas da Última Semana", resumo)
    bloco("Previstas para Esta Semana", previstas)
    bloco("Pendente de Assinatura", assinatura)
    bloco("Impedimentos ou Pendências do Parceiro", impedimentos)

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
        # plan_id_fixo: usado quando PlannerID != GroupID (plano direto na planilha)
        plan_id_fixo = pid if (pid and pid != gid) else None
        clientes.append({"cliente": str(cliente).strip(), "group_id": gid, "plan_id_fixo": plan_id_fixo})
    wb.close()
    return clientes


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("\n" + "=" * 60)
    print("  Follow Up Semanal — D'Artibale Faria")
    print(f"  {datetime.now(TZ).strftime('%d/%m/%Y %H:%M')}")
    print("=" * 60 + "\n")

    token = obter_token()
    print(f"Permissões do app: {roles_do_token(token) or 'NENHUMA'}\n")

    jan = janelas_semana()
    print(f"Semana anterior: {jan['ini_ant']:%d/%m} a {jan['fim_ant']:%d/%m}  |  "
          f"corrente: {jan['ini_cur']:%d/%m} a {jan['fim_cur']:%d/%m}\n")

    clientes = ler_clientes(PLANILHA)
    if TEST_LIMIT > 0:
        clientes = clientes[:TEST_LIMIT]
        print(f"[teste] MODO TESTE — apenas {len(clientes)} cliente(s).\n")
    print(f"[ok] {len(clientes)} clientes carregados.\n")

    # Resolve os IDs do consultivo lendo membros de TODOS os grupos da planilha
    todos_grupos = [c["group_id"] for c in ler_clientes(PLANILHA)]
    ids_consultivo = set(resolver_ids_consultivo(token, todos_grupos).keys())
    print(f"[ok] {len(ids_consultivo)} integrantes do consultivo resolvidos.\n")

    enviados, vazios, erros = 0, 0, []

    for i, c in enumerate(clientes, 1):
        print(f"[{i:02d}/{len(clientes)}] {c['cliente']}")
        try:
            corpo = montar_followup(token, c["cliente"], c["group_id"], ids_consultivo, jan, c.get("plan_id_fixo"))
            if corpo is None:
                print("   sem plano de Planner no grupo — pulado.")
                vazios += 1
                continue
            assunto = f"[Follow Up Semanal] {c['cliente']} — {datetime.now(TZ):%d/%m/%Y}"
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