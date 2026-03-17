#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PipeLovers - Gerador de Dashboard PDI (v2)
==========================================
Arquivos necessários na mesma pasta:
  pdi.csv        -> planilha de PDIs (coluna H = e-mail(s) do usuário)
  consumo.csv    -> relatório de consumo
  usuarios.csv   -> base completa de usuários da plataforma
  clientes.csv   -> base de clientes (com Status e CSM)

Novidades v2:
  - Match por e-mail (coluna H do PDI) em vez de nome
  - PDIs de time: e-mails separados por ; geram um PDI por pessoa
  - Nova aba "Sem PDI": usuários de empresas ativas sem PDI atribuído
"""

import csv
import re
import json
import os
import sys
from collections import defaultdict
from datetime import datetime

# ─────────────────────────────────────────
#  CONFIGURAÇÃO
# ─────────────────────────────────────────
PDI_CSV       = "pdi.csv"
CONSUMO_CSV   = "consumo.csv"
USUARIOS_CSV  = "usuarios.csv"
CLIENTES_CSV  = "clientes.csv"
OUTPUT_HTML   = "index.html"

# ─────────────────────────────────────────
#  UTILITÁRIOS
# ─────────────────────────────────────────
def encontrar_csv(nome_config, prefixos):
    pasta = os.path.dirname(os.path.abspath(__file__))
    caminho = os.path.join(pasta, nome_config)
    if os.path.exists(caminho):
        return caminho
    for f in os.listdir(pasta):
        fl = f.lower()
        if fl.endswith('.csv') and any(fl.startswith(p.lower()) for p in prefixos):
            print(f"  Auto-detectado: {f}")
            return os.path.join(pasta, f)
    return None

def normalizar_titulo(t):
    t = t.lower().strip()
    t = re.sub(r'^\d+\.\d+\s*[-—]?\s*', '', t)
    t = re.sub(r'^\d+\)\s*', '', t)
    t = re.sub(r'[^\w\s]', ' ', t)
    t = re.sub(r'\s+', ' ', t)
    return t.strip()

def titulos_batem(pdi, consumido):
    np = normalizar_titulo(pdi)
    nc = normalizar_titulo(consumido)
    if not np or not nc:
        return False
    if np in nc or nc in np:
        return True
    wp = set(np.split())
    wc = set(nc.split())
    if len(wp) >= 3 and len(wc) >= 3:
        if len(wp & wc) / max(len(wp), len(wc)) >= 0.55:
            return True
    return False

def extrair_aulas_com_link(resumo):
    aulas = []
    atual = None
    for linha in resumo.split('\n'):
        ls = linha.strip()
        m = re.match(r'^\d+\)\s+(.+)', ls)
        if m:
            if atual:
                aulas.append(atual)
            titulo = m.group(1).strip()
            if titulo.startswith('http') or titulo.startswith('-') or titulo.startswith('Responsável'):
                atual = None
                continue
            titulo = re.sub(r'\s*[-—]\s*(Responsável|Link|https?).*', '', titulo)
            titulo = re.sub(r'\s*—\s*[A-Z][^—\n]+$', '', titulo)
            titulo = titulo.strip()
            if titulo and len(titulo) > 5:
                atual = {'titulo': titulo, 'link': ''}
            else:
                atual = None
        elif atual and 'https://app.hub.la' in ls:
            url = re.search(r'(https://app\.hub\.la\S+)', ls)
            if url:
                atual['link'] = url.group(1)
    if atual:
        aulas.append(atual)
    return aulas

def parse_emails(campo):
    """Extrai lista de e-mails do campo H (separados por ;)."""
    if not campo:
        return []
    emails = []
    for part in re.split(r'[;,]', campo):
        e = part.strip().lower()
        if '@' in e:
            emails.append(e)
    return emails

# ─────────────────────────────────────────
#  PROCESSAMENTO
# ─────────────────────────────────────────
def processar():
    print("\nPipeLovers - Dashboard PDI v2")
    print("=" * 40)

    p_pdi      = encontrar_csv(PDI_CSV,      ['pdi', 'controle_pdi'])
    p_consumo  = encontrar_csv(CONSUMO_CSV,  ['consumo', 'pipelovers_b2b', 'consumption'])
    p_usuarios = encontrar_csv(USUARIOS_CSV, ['usuarios', 'membros'])
    p_clientes = encontrar_csv(CLIENTES_CSV, ['clientes', 'base_de_clientes', 'planilha'])

    for nome, caminho in [('pdi', p_pdi), ('consumo', p_consumo)]:
        if not caminho:
            print(f"\nERRO: '{nome}.csv' nao encontrado.")
            sys.exit(1)

    print(f"\nArquivos:")
    print(f"  PDI:      {os.path.basename(p_pdi)}")
    print(f"  Consumo:  {os.path.basename(p_consumo)}")
    if p_usuarios: print(f"  Usuarios: {os.path.basename(p_usuarios)}")
    if p_clientes: print(f"  Clientes: {os.path.basename(p_clientes)}")

    # ── Consumo: indexar por e-mail e por nome ──
    print("\nCarregando consumo...")
    consumo_por_email = defaultdict(list)
    consumo_por_nome  = defaultdict(list)
    empresa_por_email = {}
    nome_por_email    = {}

    with open(p_consumo, encoding='utf-8') as f:
        for row in csv.DictReader(f):
            email   = row.get('user_email', '').strip().lower()
            nome    = row.get('user_full_name', '').strip()
            titulo  = row.get('title', '').strip()
            data    = row.get('first_consumed_at', '')
            empresa = row.get('company', '').strip()
            if not titulo: continue
            entrada = {
                'titulo': titulo,
                'data':   data[:10] if data else '',
                'mes':    data[:7]  if data else '',
            }
            if email:
                consumo_por_email[email].append(entrada)
                empresa_por_email[email] = empresa
                if nome: nome_por_email[email] = nome
            if nome:
                key = re.sub(r'\s+', ' ', nome.strip().lower())
                consumo_por_nome[key].append(entrada)

    print(f"  {len(consumo_por_email):,} usuários únicos")

    # ── PDIs: ler coluna H (e-mails) e expandir times ──
    print("\nCarregando PDIs...")

    # Estrutura intermediária: key -> {aulas, cargo, email_cs, emails_pdi}
    # key pode ser email (se coluna H preenchida) ou nome
    pdis_raw = {}

    with open(p_pdi, encoding='utf-8') as f:
        reader = csv.DictReader(f)
        colunas = reader.fieldnames or []

        # Detectar coluna H (índice 7) — pode ter nome variado
        col_email_pdi = None
        if len(colunas) > 7:
            col_email_pdi = colunas[7]  # coluna H = índice 7
        print(f"  Coluna H detectada: '{col_email_pdi}'")

        for row in reader:
            nome    = row.get('Nome da pessoa', '').strip()
            resumo  = row.get('Resumo do PDI', '')
            cargo   = row.get('Cargo', '').strip()
            email_cs = row.get('Email CS', '').strip()
            emails_h = parse_emails(row.get(col_email_pdi, '') if col_email_pdi else '')

            aulas = extrair_aulas_com_link(resumo)
            if not aulas or not nome:
                continue

            info = {'aulas': aulas, 'cargo': cargo, 'email_cs': email_cs,
                    'emails_pdi': emails_h, 'nome_original': nome}

            # Chave: se tem e-mails na coluna H, usa o primeiro e-mail
            # (times com múltiplos e-mails serão expandidos depois)
            if emails_h:
                key = emails_h[0]  # chave temporária
            else:
                key = nome

            # Manter o PDI mais completo por chave
            if key not in pdis_raw or len(aulas) > len(pdis_raw[key]['aulas']):
                pdis_raw[key] = info

    print(f"  {len(pdis_raw)} PDIs carregados")

    # ── Expandir PDIs de time ──
    # Para cada PDI com múltiplos e-mails na coluna H,
    # criar uma entrada individual por pessoa
    pdis_expandidos = {}  # email_ou_nome -> info

    for key, info in pdis_raw.items():
        emails = info['emails_pdi']
        if len(emails) <= 1:
            # PDI individual
            k = emails[0] if emails else info['nome_original']
            if k not in pdis_expandidos or len(info['aulas']) > len(pdis_expandidos[k]['aulas']):
                pdis_expandidos[k] = info
        else:
            # PDI de time: criar uma entrada por e-mail
            for email in emails:
                if email not in pdis_expandidos or len(info['aulas']) > len(pdis_expandidos[email]['aulas']):
                    pdis_expandidos[email] = info

    print(f"  {len(pdis_expandidos)} PDIs após expansão de times")

    # ── Cruzar PDIs com consumo ──
    print("\nCruzando dados...")
    dados_finais = []
    emails_com_pdi = set()

    for key, info in pdis_expandidos.items():
        # Buscar consumo por e-mail (prioritário) ou por nome
        if '@' in key:
            email = key
            entradas = consumo_por_email.get(email, [])
            empresa  = empresa_por_email.get(email, '')
            nome_exib = nome_por_email.get(email, info['nome_original'])
            emails_com_pdi.add(email)
        else:
            email = ''
            norm = re.sub(r'\s+', ' ', key.strip().lower())
            entradas = consumo_por_nome.get(norm, [])
            empresa  = entradas[0].get('empresa', '') if entradas else '' # type: ignore
            nome_exib = key

        titulos_consumidos = [e['titulo'] for e in entradas]
        consumo_por_mes = defaultdict(int)
        for e in entradas:
            if e['mes']:
                consumo_por_mes[e['mes']] += 1

        aulas_status = []
        for aula in info['aulas']:
            consumida = any(titulos_batem(aula['titulo'], ct) for ct in titulos_consumidos)
            aulas_status.append({
                'title':    aula['titulo'],
                'link':     aula.get('link', ''),
                'consumed': consumida,
            })

        total_pdi  = len(aulas_status)
        feitas_pdi = sum(1 for a in aulas_status if a['consumed'])
        pct_pdi    = round(feitas_pdi / total_pdi * 100) if total_pdi > 0 else 0

        dados_finais.append({
            'name':              nome_exib,
            'email':             email,
            'cargo':             info['cargo'],
            'email_cs':          info['email_cs'],
            'company':           empresa,
            'total_consumed':    len(titulos_consumidos),
            'consumed_by_month': dict(sorted(consumo_por_mes.items())),
            'pdi_lessons':       aulas_status,
            'pdi_total':         total_pdi,
            'pdi_done':          feitas_pdi,
            'pdi_pct':           pct_pdi,
        })

    print(f"  {len(dados_finais)} pessoas com PDI")
    print(f"  {sum(1 for d in dados_finais if d['total_consumed'] > 0)} com consumo encontrado")

    # ── Aba Sem PDI ──
    sem_pdi = []

    if p_usuarios and p_clientes:
        print("\nCarregando base para aba 'Sem PDI'...")

        # Base de clientes: status prioritário
        STATUS_PRIORITY = {'Ativo': 0, 'Try and Buy': 1, 'Inativo': 2, 'Churn': 3, '': 4}
        clientes = {}
        with open(p_clientes, encoding='utf-8') as f:
            for row in csv.DictReader(f):
                emp = row.get('Empresa', '').strip()
                if not emp: continue
                st  = (row.get('Status') or row.get(' Status') or '').strip()
                csm = row.get('CSM', '').strip()
                if emp not in clientes or STATUS_PRIORITY.get(st, 4) < STATUS_PRIORITY.get(clientes[emp]['status'], 4):
                    clientes[emp] = {'status': st, 'csm': csm}

        ativas = {e for e, d in clientes.items() if d['status'] == 'Ativo'}
        ativas_lower = {e.lower(): e for e in ativas}

        def get_empresa_ativa(nome_emp):
            if nome_emp in ativas: return nome_emp
            return ativas_lower.get(nome_emp.lower())

        # Usuários de empresas ativas que NÃO têm PDI
        with open(p_usuarios, encoding='utf-8') as f:
            for row in csv.DictReader(f):
                email = (row.get('Url do E-mail do Membro') or
                         row.get('email') or row.get('Email') or '').strip().lower()
                if '@' not in email: continue

                # Pular quem já tem PDI
                if email in emails_com_pdi: continue

                nome    = (row.get('Nome Completo') or '').strip()
                company = (row.get('Nome da Empresa') or '').strip()

                emp_ativa = get_empresa_ativa(company)
                if not emp_ativa: continue  # empresa não está ativa

                info_co = clientes[emp_ativa]
                sem_pdi.append({
                    'email':   email,
                    'name':    nome,
                    'company': emp_ativa,
                    'csm':     info_co['csm'],
                })

        print(f"  {len(sem_pdi)} usuários de empresas ativas sem PDI")
    else:
        print("\n  (usuarios.csv ou clientes.csv nao encontrados — aba 'Sem PDI' vazia)")

    return dados_finais, sem_pdi

# ─────────────────────────────────────────
#  GERAÇÃO DO HTML
# ─────────────────────────────────────────
def gerar_html(dados_finais, sem_pdi):
    json_dados   = json.dumps(dados_finais, ensure_ascii=False, separators=(',', ':'))
    json_sem_pdi = json.dumps(sem_pdi,      ensure_ascii=False, separators=(',', ':'))
    agora = datetime.now().strftime('%d/%m/%Y %H:%M')

    pasta = os.path.dirname(os.path.abspath(__file__))
    caminho = os.path.join(pasta, OUTPUT_HTML)
    print(f"\nGerando HTML...")

    with open(caminho, 'w', encoding='utf-8') as f:
        f.write('''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PipeLovers - Dashboard PDI</title>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{
  --navy:#0f2952;--mid:#2563b0;--sky:#3b82f6;--sky-l:#eff6ff;--sky-p:#f8faff;
  --g50:#f9fafb;--g100:#f1f5f9;--g200:#e2e8f0;--g400:#94a3b8;--g500:#64748b;
  --g600:#475569;--g700:#334155;--g800:#1e293b;
  --green:#059669;--gl:#ecfdf5;--gb:#6ee7b7;
  --yel:#d97706;--yl:#fffbeb;--amber:#f59e0b;
  --red:#dc2626;--rl:#fef2f2;
  --ora:#ea580c;
  --r:12px;--sh:0 1px 3px rgba(15,41,82,.07),0 1px 2px rgba(15,41,82,.04);
}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:"Plus Jakarta Sans",sans-serif;background:var(--g50);color:var(--g800);min-height:100vh;font-size:14px;}
.hdr{background:var(--navy);height:56px;padding:0 24px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:400;box-shadow:0 2px 10px rgba(15,41,82,.2);}
.brand{display:flex;align-items:center;gap:10px;}
.bico{width:30px;height:30px;background:var(--sky);border-radius:7px;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:800;color:#fff;}
.bname{font-size:16px;font-weight:700;color:#fff;}.btag{font-size:11px;color:rgba(255,255,255,.4);margin-top:1px;}
.hright{display:flex;align-items:center;gap:12px;}
.upd{font-size:11px;color:rgba(255,255,255,.35);}
.btn-hdr{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:rgba(255,255,255,.85);padding:6px 14px;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;transition:all .2s;}
.btn-hdr:hover{background:rgba(255,255,255,.18);color:#fff;}
.nav{background:#fff;border-bottom:1px solid var(--g200);padding:0 24px;display:flex;position:sticky;top:56px;z-index:300;box-shadow:var(--sh);}
.ntab{padding:13px 16px;font-size:13px;font-weight:600;color:var(--g500);cursor:pointer;border-bottom:2px solid transparent;transition:all .15s;white-space:nowrap;}
.ntab:hover{color:var(--mid);}.ntab.on{color:var(--mid);border-bottom-color:var(--sky);}
.fbar{background:#fff;border-bottom:1px solid var(--g200);padding:10px 24px;display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap;}
.fg{display:flex;flex-direction:column;gap:3px;}
.fl{font-size:10px;font-weight:700;color:var(--g400);text-transform:uppercase;letter-spacing:.5px;}
select,input[type=text]{background:var(--g50);border:1.5px solid var(--g200);color:var(--g800);padding:6px 10px;border-radius:7px;font-size:12px;font-family:inherit;outline:none;min-width:145px;transition:border-color .15s;}
select:focus,input:focus{border-color:var(--sky);background:#fff;}
.btnf{background:#fff;border:1.5px solid var(--g200);color:var(--g600);padding:6px 12px;border-radius:7px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;transition:all .15s;align-self:flex-end;}
.btnf:hover{border-color:var(--sky);color:var(--mid);}
.ctag{font-size:12px;font-weight:600;color:var(--mid);background:var(--sky-l);border:1px solid #bfdbfe;padding:4px 10px;border-radius:20px;align-self:flex-end;}
/* PDI TAB */
.pdi-main{display:flex;height:calc(100vh - 149px);}
.list-panel{width:300px;flex-shrink:0;background:#fff;border-right:1px solid var(--g200);overflow-y:auto;}
.lhdr{padding:12px 16px 8px;font-size:10px;font-weight:700;color:var(--g400);text-transform:uppercase;letter-spacing:.7px;border-bottom:1px solid var(--g100);background:#fff;position:sticky;top:0;z-index:10;}
.prow{padding:12px 16px;cursor:pointer;border-bottom:1px solid var(--g100);transition:background .12s;border-left:3px solid transparent;}
.prow:hover{background:var(--sky-p);}.prow.on{background:var(--sky-l);border-left-color:var(--sky);}
.pr-name{font-size:13px;font-weight:600;margin-bottom:3px;}
.pr-meta{display:flex;justify-content:space-between;align-items:center;gap:6px;}
.pr-co{font-size:11px;color:var(--g400);max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.pr-pdi{display:flex;align-items:center;gap:5px;}
.pr-bar{width:44px;height:4px;background:var(--g200);border-radius:2px;overflow:hidden;flex-shrink:0;}
.pr-fill{height:100%;border-radius:2px;}
.pr-pct{font-size:11px;font-weight:700;}
.det-panel{flex:1;overflow-y:auto;padding:24px 28px;background:var(--g50);}
.empty-st{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:12px;}
.empty-ico{width:60px;height:60px;background:var(--sky-l);border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:26px;}
.empty-txt{font-size:15px;font-weight:600;color:var(--g600);}
.empty-sub{font-size:13px;color:var(--g400);}
.det-name{font-size:22px;font-weight:800;color:var(--navy);margin-bottom:6px;}
.badges{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:20px;}
.badge{padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600;}
.b-co{background:var(--sky-l);color:var(--mid);border:1px solid #bfdbfe;}
.b-role{background:var(--g100);color:var(--g600);border:1px solid var(--g200);}
.krow{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px;}
.kpi{background:#fff;border:1.5px solid var(--g200);border-radius:var(--r);padding:14px 16px;box-shadow:var(--sh);}
.klbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--g400);margin-bottom:6px;}
.kval{font-size:26px;font-weight:800;line-height:1;letter-spacing:-.5px;}
.kv-b{color:var(--mid);}.kv-g{color:var(--green);}.kv-y{color:var(--yel);}.kv-o{color:var(--ora);}
.sec{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--g400);margin-bottom:10px;display:flex;align-items:center;gap:8px;}
.sec::after{content:"";flex:1;height:1px;background:var(--g200);}
.prog-card{background:#fff;border:1.5px solid var(--g200);border-radius:var(--r);padding:18px 22px;margin-bottom:18px;display:flex;align-items:center;gap:20px;box-shadow:var(--sh);}
.prog-pct{font-size:48px;font-weight:800;line-height:1;letter-spacing:-2px;}
.prog-lbl{font-size:12px;color:var(--g400);margin-top:3px;}
.prog-bar-out{flex:1;background:var(--g100);border-radius:8px;height:12px;overflow:hidden;}
.prog-bar-in{height:100%;border-radius:8px;transition:width .7s cubic-bezier(.4,0,.2,1);}
.chart-card{background:#fff;border:1.5px solid var(--g200);border-radius:var(--r);padding:18px 22px;margin-bottom:18px;box-shadow:var(--sh);}
.chart-bars{display:flex;align-items:flex-end;gap:8px;height:100px;}
.bar-col{flex:1;display:flex;flex-direction:column;align-items:center;gap:4px;min-width:0;}
.bar-body{width:100%;min-height:4px;border-radius:4px 4px 0 0;background:var(--sky);position:relative;}
.bar-body:hover{background:var(--mid);}
.bar-tip{position:absolute;bottom:calc(100% + 4px);left:50%;transform:translateX(-50%);background:var(--navy);color:#fff;padding:2px 7px;border-radius:5px;font-size:10px;font-weight:600;white-space:nowrap;opacity:0;pointer-events:none;transition:opacity .15s;z-index:20;}
.bar-body:hover .bar-tip{opacity:1;}
.bar-lbl{font-size:9px;color:var(--g400);font-weight:600;text-align:center;overflow:hidden;max-width:100%;}
.no-chart{text-align:center;padding:24px;color:var(--g400);font-size:13px;}
.lgrid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:20px;}
.lcard{background:#fff;border:1.5px solid var(--g200);border-radius:10px;padding:12px 14px;display:flex;align-items:flex-start;gap:10px;box-shadow:var(--sh);}
.lcard-ok{border-color:var(--gb);background:var(--gl);}
.lcard-nd{opacity:.75;}
.lico{width:22px;height:22px;border-radius:50%;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:11px;margin-top:1px;}
.lico-ok{background:#d1fae5;color:var(--green);}
.lico-nd{background:var(--g100);color:var(--g400);border:1px solid var(--g200);}
.l-title{font-size:12px;line-height:1.45;font-weight:500;}
.l-st{font-size:10px;font-weight:700;margin-top:3px;}
.l-ok{color:var(--green);}.l-nd{color:var(--g400);}
.l-link{display:inline-block;margin-top:4px;font-size:11px;color:var(--sky);font-weight:600;text-decoration:none;}
.l-link:hover{text-decoration:underline;}
/* SEM PDI TAB */
.pg{display:none;padding:20px 24px;}.pg.on{display:block;}
.twrap{background:#fff;border:1.5px solid var(--g200);border-radius:var(--r);overflow:hidden;box-shadow:var(--sh);margin-bottom:20px;}
.thdr{padding:12px 16px;border-bottom:1px solid var(--g100);display:flex;justify-content:space-between;align-items:center;}
.ttl{font-size:13px;font-weight:700;color:var(--g700);}.tcnt{font-size:12px;color:var(--g400);}
.tscr{overflow-x:auto;max-height:600px;overflow-y:auto;}
table{width:100%;border-collapse:collapse;}
th{padding:9px 12px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--g400);background:var(--g50);border-bottom:1px solid var(--g200);white-space:nowrap;position:sticky;top:0;z-index:1;}
td{padding:10px 12px;font-size:12px;border-bottom:1px solid var(--g100);vertical-align:middle;}
tr:last-child td{border-bottom:none;}tr:hover td{background:var(--sky-p);}
.pill{display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;}
.p-y{background:var(--yl);color:var(--yel);border:1px solid #fcd34d;}
.krow3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px;max-width:560px;}
::-webkit-scrollbar{width:5px;height:5px;}::-webkit-scrollbar-track{background:transparent;}::-webkit-scrollbar-thumb{background:var(--g200);border-radius:3px;}
@media print{.hdr,.nav,.fbar,.list-panel{display:none!important;}.det-panel,.pg{padding:16px;overflow:visible;}.pdi-main{height:auto;}}
</style>
</head>
<body>
<div class="hdr">
  <div class="brand">
    <div class="bico">PL</div>
    <div><div class="bname">PipeLovers</div><div class="btag">Dashboard de PDIs</div></div>
  </div>
  <div class="hright">
    <span class="upd" id="upd-lbl"></span>
    <button class="btn-hdr" onclick="window.print()">&#128424; Imprimir</button>
  </div>
</div>
<div class="nav">
  <div class="ntab on" onclick="goTab(\'pdi\',this)">&#128203; PDIs</div>
  <div class="ntab"    onclick="goTab(\'sempdi\',this)">&#9888;&#65039; Sem PDI</div>
</div>

<!-- ═══ TAB PDI ═══ -->
<div id="pg-pdi" style="display:block;">
  <div class="fbar">
    <div class="fg"><div class="fl">Buscar pessoa</div><input type="text" id="s-name" placeholder="Nome ou e-mail..."></div>
    <div class="fg"><div class="fl">Empresa</div><select id="s-co"><option value="">Todas</option></select></div>
    <div class="fg"><div class="fl">Progresso PDI</div>
      <select id="s-pr">
        <option value="">Todos</option>
        <option value="ns">Nao iniciado (0%)</option>
        <option value="ip">Em andamento (1-99%)</option>
        <option value="ok">Concluido (100%)</option>
      </select>
    </div>
    <button class="btnf" id="pdi-rst">Limpar</button>
    <div class="ctag" id="pdi-ctag"></div>
  </div>
  <div class="pdi-main">
    <div class="list-panel">
      <div class="lhdr">Colaboradores</div>
      <div id="plist"></div>
    </div>
    <div class="det-panel" id="dpanel">
      <div class="empty-st">
        <div class="empty-ico">&#128203;</div>
        <div class="empty-txt">Selecione um colaborador</div>
        <div class="empty-sub">Clique em um nome na lista para ver o PDI</div>
      </div>
    </div>
  </div>
</div>

<!-- ═══ TAB SEM PDI ═══ -->
<div id="pg-sempdi" class="pg">
  <div class="fbar">
    <div class="fg"><div class="fl">Buscar</div><input type="text" id="sp-q" placeholder="Nome ou e-mail..."></div>
    <div class="fg"><div class="fl">Empresa</div><select id="sp-co"><option value="">Todas</option></select></div>
    <div class="fg"><div class="fl">CSM</div><select id="sp-csm"><option value="">Todos</option></select></div>
    <button class="btnf" id="sp-rst">Limpar</button>
    <div class="ctag" id="sp-ctag"></div>
  </div>
  <div id="sp-body" style="padding:20px 24px;"></div>
</div>

<script>
''')

        f.write('var D=')
        f.write(json_dados)
        f.write(';\nvar SP=')
        f.write(json_sem_pdi)
        f.write(';\nvar UPD="')
        f.write(agora)
        f.write('";\n')

        f.write(r'''
document.getElementById("upd-lbl").textContent="Atualizado em "+UPD;
function esc(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function pct(a,b){return b>0?Math.round(a/b*100):0;}
function barCol(p){if(p===100)return"#059669";if(p>50)return"#2563b0";if(p>0)return"#f59e0b";return"#94a3b8";}

function goTab(n,el){
  document.getElementById("pg-pdi").style.display=(n==="pdi")?"block":"none";
  document.getElementById("pg-sempdi").className=(n==="sempdi")?"pg on":"pg";
  document.querySelectorAll(".ntab").forEach(function(t){t.classList.remove("on");});
  el.classList.add("on");
}

/* ══ PDI TAB ══ */
var fd=D.slice(), sel=null;
var cos=[...new Set(D.map(function(d){return d.company;}))].sort();
var coSel=document.getElementById("s-co");
cos.forEach(function(c){var o=document.createElement("option");o.value=c;o.textContent=c;coSel.appendChild(o);});

function filterPdi(){
  var q=document.getElementById("s-name").value.toLowerCase();
  var co=document.getElementById("s-co").value;
  var pr=document.getElementById("s-pr").value;
  fd=D.filter(function(d){
    if(q&&d.name.toLowerCase().indexOf(q)<0&&d.email.toLowerCase().indexOf(q)<0)return false;
    if(co&&d.company!==co)return false;
    if(pr==="ns"&&d.pdi_pct!==0)return false;
    if(pr==="ip"&&(d.pdi_pct===0||d.pdi_pct===100))return false;
    if(pr==="ok"&&d.pdi_pct!==100)return false;
    return true;
  });
  document.getElementById("pdi-ctag").textContent=fd.length+" pessoas";
  renderList();
}
function resetPdi(){
  document.getElementById("s-name").value="";
  document.getElementById("s-co").value="";
  document.getElementById("s-pr").value="";
  filterPdi();
}
function renderList(){
  var c=document.getElementById("plist");c.innerHTML="";
  if(!fd.length){c.innerHTML='<div style="padding:30px;text-align:center;color:var(--g400);">Nenhum resultado</div>';return;}
  fd.forEach(function(p){
    var div=document.createElement("div");
    div.className="prow"+(sel===p.email||sel===p.name?" on":"");
    var col=barCol(p.pdi_pct);
    div.innerHTML='<div class="pr-name">'+esc(p.name)+'</div>'+
      '<div class="pr-meta">'+
        '<span class="pr-co">'+esc(p.company||"\u2014")+'</span>'+
        '<div class="pr-pdi">'+
          '<div class="pr-bar"><div class="pr-fill" style="width:'+p.pdi_pct+'%;background:'+col+'"></div></div>'+
          '<span class="pr-pct" style="color:'+col+'">'+p.pdi_pct+'%</span>'+
        '</div>'+
      '</div>';
    div.addEventListener("click",function(){pickPerson(p);});
    c.appendChild(div);
  });
}
function pickPerson(p){
  sel=p.email||p.name;
  renderList();
  renderDetail(p);
}
function renderDetail(p){
  var panel=document.getElementById("dpanel");
  var col=barCol(p.pdi_pct);
  var kpi='<div class="krow">'+
    '<div class="kpi"><div class="klbl">Aulas assistidas</div><div class="kval kv-b">'+p.total_consumed+'</div></div>'+
    '<div class="kpi"><div class="klbl">PDI concluido</div><div class="kval kv-g">'+p.pdi_done+' / '+p.pdi_total+'</div></div>'+
    '<div class="kpi"><div class="klbl">Evolucao</div><div class="kval" style="color:'+col+'">'+p.pdi_pct+'%</div></div>'+
    '<div class="kpi"><div class="klbl">Pendentes</div><div class="kval kv-y">'+(p.pdi_total-p.pdi_done)+'</div></div>'+
  '</div>';
  var prog='<div class="sec">Progresso do PDI</div>'+
    '<div class="prog-card">'+
      '<div><div class="prog-pct" style="color:'+col+'">'+p.pdi_pct+'%</div>'+
      '<div class="prog-lbl">'+p.pdi_done+' de '+p.pdi_total+' aulas concluidas</div></div>'+
      '<div class="prog-bar-out"><div class="prog-bar-in" style="width:'+p.pdi_pct+'%;background:'+col+'"></div></div>'+
    '</div>';
  var months=Object.keys(p.consumed_by_month||{});
  var chart='<div class="sec">Consumo mensal</div><div class="chart-card">';
  if(months.length){
    var mx=1;months.forEach(function(m){if(p.consumed_by_month[m]>mx)mx=p.consumed_by_month[m];});
    chart+='<div class="chart-bars">';
    months.forEach(function(m){
      var v=p.consumed_by_month[m],h=Math.max(4,Math.round(v/mx*90));
      chart+='<div class="bar-col"><div class="bar-body" style="height:'+h+'px"><div class="bar-tip">'+v+' aulas</div></div>'+
        '<div class="bar-lbl">'+m.replace(/-/g,"/")+'</div></div>';
    });
    chart+='</div>';
  }else{chart+='<div class="no-chart">Sem dados de consumo ainda</div>';}
  chart+='</div>';
  var lessons='<div class="sec">Aulas do PDI</div><div class="lgrid">';
  p.pdi_lessons.forEach(function(l){
    var ok=l.consumed;
    var lnk=l.link?'<a class="l-link" href="'+esc(l.link)+'" target="_blank">Acessar aula &rarr;</a>':"";
    lessons+='<div class="lcard '+(ok?"lcard-ok":"lcard-nd")+'">'+
      '<div class="lico '+(ok?"lico-ok":"lico-nd")+'">'+(ok?"&#10003;":"&#9679;")+'</div>'+
      '<div><div class="l-title">'+esc(l.title)+'</div>'+
      '<div class="l-st '+(ok?"l-ok":"l-nd")+'">'+(ok?"Concluida":"Pendente")+'</div>'+lnk+'</div></div>';
  });
  lessons+='</div>';
  var bgs='<div class="badges">';
  if(p.company)bgs+='<span class="badge b-co">&#127970; '+esc(p.company)+'</span>';
  if(p.email)  bgs+='<span class="badge b-role">'+esc(p.email)+'</span>';
  if(p.cargo)  bgs+='<span class="badge b-role">'+esc(p.cargo)+'</span>';
  if(p.email_cs)bgs+='<span class="badge b-role">CS: '+esc(p.email_cs)+'</span>';
  bgs+='</div>';
  panel.innerHTML='<div class="det-name">'+esc(p.name)+'</div>'+bgs+kpi+prog+chart+lessons;
}
document.getElementById("s-name").addEventListener("input",filterPdi);
["s-co","s-pr"].forEach(function(id){document.getElementById(id).addEventListener("change",filterPdi);});
document.getElementById("pdi-rst").addEventListener("click",resetPdi);

/* ══ SEM PDI TAB ══ */
var scos=[...new Set(SP.map(function(u){return u.company;}))].sort();
var scsms=[...new Set(SP.map(function(u){return u.csm;}).filter(Boolean))].sort();
function pop(id,vals){var s=document.getElementById(id);vals.forEach(function(v){var o=document.createElement("option");o.value=v;o.textContent=v;s.appendChild(o);});}
pop("sp-co",scos);pop("sp-csm",scsms);

function filterSp(){
  var q=document.getElementById("sp-q").value.toLowerCase();
  var co=document.getElementById("sp-co").value;
  var csm=document.getElementById("sp-csm").value;
  var fd=SP.filter(function(u){
    if(q&&u.name.toLowerCase().indexOf(q)<0&&u.email.toLowerCase().indexOf(q)<0)return false;
    if(co&&u.company!==co)return false;
    if(csm&&u.csm!==csm)return false;
    return true;
  });
  document.getElementById("sp-ctag").textContent=fd.length+" usuarios";
  var kpis='<div class="krow3">'+
    '<div class="kpi"><div class="klbl">Sem PDI</div><div class="kval kv-y">'+fd.length+'</div><div style="font-size:11px;color:var(--g400);margin-top:3px;">usuarios sem PDI ativo</div></div>'+
    '<div class="kpi"><div class="klbl">Empresas</div><div class="kval kv-b">'+[...new Set(fd.map(function(u){return u.company;}))].length+'</div></div>'+
    '<div class="kpi"><div class="klbl">CSMs</div><div class="kval" style="color:var(--g700);">'+[...new Set(fd.map(function(u){return u.csm;}).filter(Boolean))].length+'</div></div>'+
  '</div>';
  var rows=fd.map(function(u){
    return '<tr><td><div style="font-weight:600;">'+esc(u.name)+'</div>'+
      '<div style="font-size:11px;color:var(--g400);">'+esc(u.email)+'</div></td>'+
      '<td>'+esc(u.company)+'</td>'+
      '<td>'+esc(u.csm||"\u2014")+'</td>'+
      '<td><span class="pill p-y">&#9888; Sem PDI</span></td></tr>';
  }).join('');
  document.getElementById("sp-body").innerHTML=kpis+
    '<div class="twrap"><div class="thdr"><div class="ttl">Usuarios de empresas ativas sem PDI atribuido</div>'+
    '<div class="tcnt">'+fd.length+' usuarios</div></div>'+
    '<div class="tscr"><table><thead><tr><th>Usuario</th><th>Empresa</th><th>CSM</th><th>Status</th></tr></thead>'+
    '<tbody>'+rows+'</tbody></table></div></div>';
}
document.getElementById("sp-q").addEventListener("input",filterSp);
["sp-co","sp-csm"].forEach(function(id){document.getElementById(id).addEventListener("change",filterSp);});
document.getElementById("sp-rst").addEventListener("click",function(){
  document.getElementById("sp-q").value="";
  ["sp-co","sp-csm"].forEach(function(id){document.getElementById(id).value="";});
  filterSp();
});

filterPdi();
filterSp();
''')
        f.write('</script>\n</body>\n</html>\n')

    print(f"  {OUTPUT_HTML} gerado ({os.path.getsize(caminho)//1024} KB)")
    return caminho


# ─────────────────────────────────────────
#  ENTRADA
# ─────────────────────────────────────────
if __name__ == '__main__':
    dados, sem_pdi = processar()
    caminho = gerar_html(dados, sem_pdi)
    print(f"\nPronto!")
    print(f"  {len(dados)} PDIs no dashboard")
    print(f"  {len(sem_pdi)} usuarios sem PDI")
    print("\nProximos passos:")
    print("  1. Confira o index.html no navegador")
    print("  2. Suba no GitHub para publicar")
