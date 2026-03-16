#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PipeLovers - Gerador de Dashboard PDI
======================================
Como usar:
  1. Coloque os dois CSVs na mesma pasta que este script
  2. Rode: python gerar_dashboard.py
  3. O arquivo index.html será gerado/atualizado
"""

import csv
import re
import json
import os
import sys
from collections import defaultdict
from datetime import datetime

# ─────────────────────────────────────────
#  CONFIGURAÇÃO — ajuste os nomes dos arquivos se precisar
# ─────────────────────────────────────────
PDI_CSV      = "pdi.csv"          # Planilha de PDIs (Controle_PDIs_...)
CONSUMO_CSV  = "consumo.csv"      # Relatório de consumo (pipelovers_b2b_consumption_...)
OUTPUT_HTML  = "index.html"       # Arquivo gerado (não precisa mudar)

def encontrar_csv(prefixo_opcoes):
    """Tenta encontrar o arquivo CSV pelo nome exato ou por prefixo."""
    pasta = os.path.dirname(os.path.abspath(__file__))
    for nome in os.listdir(pasta):
        nome_lower = nome.lower()
        for prefixo in prefixo_opcoes:
            if nome_lower.startswith(prefixo.lower()) and nome_lower.endswith('.csv'):
                return os.path.join(pasta, nome)
    return None

def resolver_caminho(nome_config, prefixos_fallback):
    pasta = os.path.dirname(os.path.abspath(__file__))
    caminho_exato = os.path.join(pasta, nome_config)
    if os.path.exists(caminho_exato):
        return caminho_exato
    encontrado = encontrar_csv(prefixos_fallback)
    if encontrado:
        print(f"  Encontrado automaticamente: {os.path.basename(encontrado)}")
        return encontrado
    return None

# ─────────────────────────────────────────
#  EXTRAÇÃO DE AULAS DO TEXTO DO PDI
# ─────────────────────────────────────────
def extrair_aulas(resumo):
    aulas = []
    for linha in resumo.split('\n'):
        linha = linha.strip()
        m = re.match(r'^\d+\)\s+(.+)', linha)
        if not m:
            continue
        titulo = m.group(1).strip()
        if titulo.startswith('http') or titulo.startswith('-') or titulo.startswith('Responsável'):
            continue
        titulo = re.sub(r'\s*[-—]\s*(Responsável|Link|https?).*', '', titulo)
        titulo = re.sub(r'\s*—\s*[A-Z][^—\n]+$', '', titulo)
        titulo = titulo.strip()
        if titulo and len(titulo) > 5:
            aulas.append(titulo)
    return aulas

def extrair_aulas_com_link(resumo):
    aulas = []
    linhas = resumo.split('\n')
    atual = None
    for linha in linhas:
        linha_s = linha.strip()
        m = re.match(r'^\d+\)\s+(.+)', linha_s)
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
        elif atual and 'https://app.hub.la' in linha_s:
            url = re.search(r'(https://app\.hub\.la\S+)', linha_s)
            if url:
                atual['link'] = url.group(1)
    if atual:
        aulas.append(atual)
    return aulas

# ─────────────────────────────────────────
#  NORMALIZAÇÃO PARA MATCHING DE TÍTULOS
# ─────────────────────────────────────────
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
        sobreposicao = len(wp & wc) / max(len(wp), len(wc))
        if sobreposicao >= 0.55:
            return True
    return False

def normalizar_nome(n):
    return re.sub(r'\s+', ' ', n.strip().lower())

# ─────────────────────────────────────────
#  PROCESSAMENTO PRINCIPAL
# ─────────────────────────────────────────
def processar():
    print("\nPipeLovers - Gerador de Dashboard PDI")
    print("=" * 40)

    # Localizar arquivos
    caminho_pdi = resolver_caminho(PDI_CSV, ['controle_pdi', 'pdi'])
    caminho_consumo = resolver_caminho(CONSUMO_CSV, ['pipelovers_b2b', 'consumo', 'consumption'])

    if not caminho_pdi:
        print(f"\nERRO: Arquivo de PDI nao encontrado.")
        print(f"  Renomeie seu arquivo para '{PDI_CSV}' ou coloque-o na mesma pasta.")
        sys.exit(1)
    if not caminho_consumo:
        print(f"\nERRO: Arquivo de consumo nao encontrado.")
        print(f"  Renomeie seu arquivo para '{CONSUMO_CSV}' ou coloque-o na mesma pasta.")
        sys.exit(1)

    print(f"\nArquivos encontrados:")
    print(f"  PDI:     {os.path.basename(caminho_pdi)}")
    print(f"  Consumo: {os.path.basename(caminho_consumo)}")

    # ── Carregar consumo ──
    print("\nCarregando dados de consumo...")
    consumo_por_email = defaultdict(list)
    consumo_por_nome  = defaultdict(list)
    empresa_por_email = {}
    nome_por_email    = {}

    with open(caminho_consumo, encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            email   = row.get('user_email', '').strip().lower()
            nome    = row.get('user_full_name', '').strip()
            titulo  = row.get('title', '').strip()
            data    = row.get('first_consumed_at', '')
            empresa = row.get('company', '').strip()
            if not titulo:
                continue
            entrada = {
                'titulo':  titulo,
                'data':    data[:10] if data else '',
                'mes':     data[:7]  if data else '',
                'empresa': empresa,
            }
            if email:
                consumo_por_email[email].append(entrada)
                empresa_por_email[email] = empresa
                if nome:
                    nome_por_email[email] = nome
            if nome:
                consumo_por_nome[normalizar_nome(nome)].append(entrada)

    print(f"  {sum(len(v) for v in consumo_por_email.values()):,} registros de consumo")
    print(f"  {len(consumo_por_email):,} usuarios unicos")

    # ── Carregar PDIs ──
    print("\nCarregando PDIs...")
    pdis = {}
    with open(caminho_pdi, encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            nome    = row.get('Nome da pessoa', '').strip()
            resumo  = row.get('Resumo do PDI', '')
            cargo   = row.get('Cargo', '').strip()
            email_cs = row.get('Email CS', '').strip()
            if not nome:
                continue
            aulas = extrair_aulas_com_link(resumo)
            if aulas and (nome not in pdis or len(aulas) > len(pdis[nome]['aulas'])):
                pdis[nome] = {'aulas': aulas, 'cargo': cargo, 'email_cs': email_cs}

    print(f"  {len(pdis)} pessoas com PDI")

    # ── Cruzamento ──
    print("\nCruzando dados...")
    dados_finais = []

    for nome_pdi, info_pdi in pdis.items():
        if not info_pdi['aulas']:
            continue

        # Encontrar consumo do usuário (por email ou nome)
        entradas_consumo = []
        empresa = ''
        nome_exibicao = nome_pdi

        if '@' in nome_pdi:
            email = nome_pdi.strip().lower()
            entradas_consumo = consumo_por_email.get(email, [])
            empresa = empresa_por_email.get(email, '')
            nome_exibicao = nome_por_email.get(email, nome_pdi)
        else:
            norm = normalizar_nome(nome_pdi)
            entradas_consumo = consumo_por_nome.get(norm, [])
            if entradas_consumo:
                empresa = entradas_consumo[0].get('empresa', '')

        titulos_consumidos = [e['titulo'] for e in entradas_consumo]

        consumo_por_mes = defaultdict(int)
        for e in entradas_consumo:
            if e['mes']:
                consumo_por_mes[e['mes']] += 1

        # Verificar quais aulas do PDI foram consumidas
        aulas_status = []
        for aula in info_pdi['aulas']:
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
            'name':             nome_exibicao,
            'pdi_name':         nome_pdi,
            'cargo':            info_pdi['cargo'],
            'email_cs':         info_pdi['email_cs'],
            'company':          empresa,
            'total_consumed':   len(titulos_consumidos),
            'consumed_by_month': dict(sorted(consumo_por_mes.items())),
            'pdi_lessons':      aulas_status,
            'pdi_total':        total_pdi,
            'pdi_done':         feitas_pdi,
            'pdi_pct':          pct_pdi,
        })

    com_consumo   = sum(1 for d in dados_finais if d['total_consumed'] > 0)
    com_progresso = sum(1 for d in dados_finais if d['pdi_done'] > 0)
    print(f"  {len(dados_finais)} pessoas no dashboard")
    print(f"  {com_consumo} com dados de consumo encontrados")
    print(f"  {com_progresso} com pelo menos 1 aula do PDI concluida")

    return dados_finais

# ─────────────────────────────────────────
#  GERAÇÃO DO HTML
# ─────────────────────────────────────────
def gerar_html(dados):
    json_dados = json.dumps(dados, ensure_ascii=False, separators=(',', ':'))
    agora = datetime.now().strftime('%d/%m/%Y')

    pasta_saida = os.path.dirname(os.path.abspath(__file__))
    caminho_html = os.path.join(pasta_saida, OUTPUT_HTML)

    print(f"\nGerando HTML...")

    HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PipeLovers - Dashboard PDI</title>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root {
  --navy:#0f2952;--blue:#1a4a8a;--blue-mid:#2563b0;--sky:#3b82f6;
  --sky-light:#eff6ff;--sky-pale:#f8faff;--white:#ffffff;
  --gray-50:#f9fafb;--gray-100:#f1f5f9;--gray-200:#e2e8f0;
  --gray-400:#94a3b8;--gray-600:#475569;--gray-800:#1e293b;
  --green:#059669;--green-bg:#ecfdf5;--green-bd:#a7f3d0;
  --amber:#d97706;--shadow-sm:0 1px 3px rgba(15,41,82,.07);
  --shadow:0 4px 16px rgba(15,41,82,.09);--radius:12px;
}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:"Plus Jakarta Sans",sans-serif;background:var(--gray-50);color:var(--gray-800);min-height:100vh;font-size:14px;}
.header{background:var(--navy);padding:0 28px;height:60px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:200;box-shadow:0 2px 12px rgba(15,41,82,.18);}
.brand{display:flex;align-items:center;gap:10px;}
.brand-icon{width:32px;height:32px;background:var(--sky);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:800;color:white;}
.brand-name{font-size:17px;font-weight:700;color:white;letter-spacing:-.3px;}
.brand-tag{font-size:12px;color:rgba(255,255,255,.5);margin-top:1px;}
.header-right{display:flex;align-items:center;gap:10px;}
.updated-chip{font-size:12px;color:rgba(255,255,255,.45);}
.btn-print{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:rgba(255,255,255,.85);padding:6px 14px;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;transition:all .2s;}
.btn-print:hover{background:rgba(255,255,255,.18);color:white;}
.filter-bar{background:white;border-bottom:1px solid var(--gray-200);padding:12px 28px;display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap;box-shadow:var(--shadow-sm);}
.fg{display:flex;flex-direction:column;gap:3px;}
.fl{font-size:10px;font-weight:700;color:var(--gray-400);text-transform:uppercase;letter-spacing:.6px;}
select,input[type="text"]{background:var(--gray-50);border:1.5px solid var(--gray-200);color:var(--gray-800);padding:7px 11px;border-radius:8px;font-size:13px;font-family:inherit;outline:none;min-width:155px;transition:border-color .15s;}
select:focus,input:focus{border-color:var(--sky);background:white;}
.btn-reset{background:white;border:1.5px solid var(--gray-200);color:var(--gray-600);padding:7px 14px;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;transition:all .15s;}
.btn-reset:hover{border-color:var(--sky);color:var(--blue);}
.count-tag{font-size:12px;font-weight:600;color:var(--blue-mid);background:var(--sky-light);border:1px solid #bfdbfe;padding:5px 10px;border-radius:20px;align-self:flex-end;}
.main{display:flex;height:calc(100vh - 109px);}
.list-panel{width:300px;flex-shrink:0;background:white;border-right:1px solid var(--gray-200);overflow-y:auto;display:flex;flex-direction:column;}
.detail-panel{flex:1;overflow-y:auto;padding:28px 32px;background:var(--gray-50);}
.list-section-title{padding:12px 16px 8px;font-size:10px;font-weight:700;color:var(--gray-400);text-transform:uppercase;letter-spacing:.7px;border-bottom:1px solid var(--gray-100);background:white;position:sticky;top:0;z-index:10;}
.person-row{padding:12px 16px;cursor:pointer;border-bottom:1px solid var(--gray-100);transition:background .12s;border-left:3px solid transparent;}
.person-row:hover{background:var(--sky-pale);}
.person-row.active{background:var(--sky-light);border-left-color:var(--sky);}
.pr-name{font-size:13px;font-weight:600;color:var(--gray-800);margin-bottom:3px;line-height:1.3;}
.pr-meta{display:flex;justify-content:space-between;align-items:center;gap:6px;}
.pr-company{font-size:11px;color:var(--gray-400);max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.pr-pdi{display:flex;align-items:center;gap:5px;}
.pr-bar{width:44px;height:4px;background:var(--gray-200);border-radius:2px;overflow:hidden;flex-shrink:0;}
.pr-fill{height:100%;border-radius:2px;}
.pr-pct{font-size:11px;font-weight:700;}
.no-results{padding:40px 16px;text-align:center;color:var(--gray-400);font-size:13px;}
.empty-state{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:12px;}
.empty-icon{width:64px;height:64px;background:var(--sky-light);border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:28px;}
.empty-text{font-size:15px;font-weight:600;color:var(--gray-600);}
.empty-sub{font-size:13px;color:var(--gray-400);}
.det-name{font-size:24px;font-weight:800;color:var(--navy);letter-spacing:-.4px;margin-bottom:6px;}
.det-badges{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:24px;}
.badge{padding:4px 10px;border-radius:20px;font-size:11px;font-weight:600;}
.badge-co{background:var(--sky-light);color:var(--blue-mid);border:1px solid #bfdbfe;}
.badge-role{background:var(--gray-100);color:var(--gray-600);border:1px solid var(--gray-200);}
.kpi-row{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px;}
.kpi{background:white;border:1.5px solid var(--gray-200);border-radius:var(--radius);padding:16px 18px;box-shadow:var(--shadow-sm);}
.kpi-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);margin-bottom:8px;}
.kpi-value{font-size:28px;font-weight:800;line-height:1;letter-spacing:-.5px;}
.kv-blue{color:var(--blue-mid);}
.kv-green{color:var(--green);}
.kv-amber{color:var(--amber);}
.sec-title{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--gray-400);margin-bottom:12px;display:flex;align-items:center;gap:8px;}
.sec-title::after{content:"";flex:1;height:1px;background:var(--gray-200);}
.progress-card{background:white;border:1.5px solid var(--gray-200);border-radius:var(--radius);padding:20px 24px;margin-bottom:20px;display:flex;align-items:center;gap:24px;box-shadow:var(--shadow-sm);}
.prog-pct{font-size:52px;font-weight:800;letter-spacing:-2px;line-height:1;}
.prog-label{font-size:13px;color:var(--gray-400);margin-top:4px;font-weight:500;}
.prog-bar-wrap{flex:1;}
.prog-bar-outer{background:var(--gray-100);border-radius:8px;height:12px;overflow:hidden;}
.prog-bar-inner{height:100%;border-radius:8px;transition:width .7s cubic-bezier(.4,0,.2,1);}
.chart-card{background:white;border:1.5px solid var(--gray-200);border-radius:var(--radius);padding:20px 24px;margin-bottom:20px;box-shadow:var(--shadow-sm);}
.chart-bars{display:flex;align-items:flex-end;gap:8px;height:100px;}
.bar-col{flex:1;display:flex;flex-direction:column;align-items:center;gap:5px;min-width:0;}
.bar-body{width:100%;min-height:4px;border-radius:5px 5px 0 0;background:var(--sky);position:relative;cursor:default;transition:background .15s;}
.bar-body:hover{background:var(--blue-mid);}
.bar-tip{position:absolute;bottom:calc(100% + 5px);left:50%;transform:translateX(-50%);background:var(--navy);color:white;padding:3px 8px;border-radius:6px;font-size:11px;font-weight:600;white-space:nowrap;opacity:0;pointer-events:none;transition:opacity .15s;z-index:20;}
.bar-body:hover .bar-tip{opacity:1;}
.bar-lbl{font-size:9px;color:var(--gray-400);font-weight:600;text-align:center;white-space:nowrap;overflow:hidden;max-width:100%;}
.no-chart{text-align:center;padding:28px;color:var(--gray-400);font-size:13px;}
.lessons-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:24px;}
.lesson{background:white;border:1.5px solid var(--gray-200);border-radius:10px;padding:13px 15px;display:flex;align-items:flex-start;gap:10px;box-shadow:var(--shadow-sm);}
.lesson-done{border-color:var(--green-bd);background:var(--green-bg);}
.lesson-pend{opacity:.78;}
.l-icon{width:22px;height:22px;border-radius:50%;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:11px;margin-top:1px;}
.l-icon-done{background:#d1fae5;color:var(--green);}
.l-icon-pend{background:var(--gray-100);color:var(--gray-400);border:1px solid var(--gray-200);}
.l-title{font-size:12px;line-height:1.45;font-weight:500;color:var(--gray-800);}
.l-status{font-size:10px;font-weight:700;margin-top:3px;}
.l-status-done{color:var(--green);}
.l-status-pend{color:var(--gray-400);}
.l-link{display:inline-block;margin-top:5px;font-size:11px;color:var(--sky);font-weight:600;text-decoration:none;}
.l-link:hover{text-decoration:underline;}
::-webkit-scrollbar{width:5px;}::-webkit-scrollbar-track{background:transparent;}::-webkit-scrollbar-thumb{background:var(--gray-200);border-radius:3px;}
@media print{.header,.filter-bar,.list-panel{display:none!important;}.detail-panel{padding:16px;overflow:visible;}.main{height:auto;}body{background:white;}}
</style>
</head>
<body>
<div class="header">
  <div class="brand">
    <div class="brand-icon">PL</div>
    <div><div class="brand-name">PipeLovers</div><div class="brand-tag">Dashboard de PDIs</div></div>
  </div>
  <div class="header-right">
    <span class="updated-chip" id="upd"></span>
    <button class="btn-print" id="print-btn">&#128424; Imprimir / PDF</button>
  </div>
</div>
<div class="filter-bar">
  <div class="fg"><div class="fl">Buscar pessoa</div><input type="text" id="s-name" placeholder="Nome do colaborador..."></div>
  <div class="fg"><div class="fl">Empresa</div><select id="s-co"><option value="">Todas as empresas</option></select></div>
  <div class="fg"><div class="fl">Progresso PDI</div>
    <select id="s-pr">
      <option value="">Todos</option>
      <option value="ns">Nao iniciado (0%)</option>
      <option value="ip">Em andamento (1-99%)</option>
      <option value="ok">Concluido (100%)</option>
    </select>
  </div>
  <button class="btn-reset" id="reset-btn">Limpar</button>
  <div class="count-tag" id="ctag"></div>
</div>
<div class="main">
  <div class="list-panel">
    <div class="list-section-title">Colaboradores</div>
    <div id="plist"></div>
  </div>
  <div class="detail-panel" id="dpanel">
    <div class="empty-state">
      <div class="empty-icon">&#128203;</div>
      <div class="empty-text">Selecione um colaborador</div>
      <div class="empty-sub">Clique em um nome na lista para ver o PDI</div>
    </div>
  </div>
</div>
<script>
PLACEHOLDER_DATA
PLACEHOLDER_DATE
var fd=D.slice(),sel=null;
document.getElementById("upd").textContent="Atualizado em "+UPD;
function pc(p){if(p===100)return"#059669";if(p>50)return"#2563b0";if(p>0)return"#3b82f6";return"#94a3b8";}
function esc(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
var cs=[];D.forEach(function(d){if(d.company&&cs.indexOf(d.company)<0)cs.push(d.company);});cs.sort();
var csel=document.getElementById("s-co");
cs.forEach(function(c){var o=document.createElement("option");o.value=c;o.textContent=c;csel.appendChild(o);});
function filter(){
  var q=document.getElementById("s-name").value.toLowerCase();
  var co=document.getElementById("s-co").value;
  var pr=document.getElementById("s-pr").value;
  fd=D.filter(function(d){
    if(q&&d.name.toLowerCase().indexOf(q)<0)return false;
    if(co&&d.company!==co)return false;
    if(pr==="ns"&&d.pdi_pct!==0)return false;
    if(pr==="ip"&&(d.pdi_pct===0||d.pdi_pct===100))return false;
    if(pr==="ok"&&d.pdi_pct!==100)return false;
    return true;
  });
  document.getElementById("ctag").textContent=fd.length+" pessoas";
  renderList();
}
function reset(){document.getElementById("s-name").value="";document.getElementById("s-co").value="";document.getElementById("s-pr").value="";filter();}
function renderList(){
  var c=document.getElementById("plist");c.innerHTML="";
  if(!fd.length){c.innerHTML='<div class="no-results">Nenhum resultado</div>';return;}
  fd.forEach(function(p){
    var div=document.createElement("div");
    div.className="person-row"+(sel===p.name?" active":"");
    var col=pc(p.pdi_pct);
    div.innerHTML='<div class="pr-name">'+esc(p.name)+'</div><div class="pr-meta"><span class="pr-company">'+esc(p.company||"&mdash;")+'</span><div class="pr-pdi"><div class="pr-bar"><div class="pr-fill" style="width:'+p.pdi_pct+'%;background:'+col+'"></div></div><span class="pr-pct" style="color:'+col+'">'+p.pdi_pct+'%</span></div></div>';
    div.addEventListener("click",function(){pick(p);});
    c.appendChild(div);
  });
}
function pick(p){sel=p.name;renderList();renderDetail(p);}
function renderDetail(p){
  var panel=document.getElementById("dpanel");
  var col=pc(p.pdi_pct);
  var kpi='<div class="kpi-row"><div class="kpi"><div class="kpi-label">Aulas assistidas</div><div class="kpi-value kv-blue">'+p.total_consumed+'</div></div><div class="kpi"><div class="kpi-label">PDI concluido</div><div class="kpi-value kv-green">'+p.pdi_done+' / '+p.pdi_total+'</div></div><div class="kpi"><div class="kpi-label">Evolucao</div><div class="kpi-value" style="color:'+col+'">'+p.pdi_pct+'%</div></div><div class="kpi"><div class="kpi-label">Pendentes</div><div class="kpi-value kv-amber">'+( p.pdi_total-p.pdi_done)+'</div></div></div>';
  var prog='<div class="sec-title">Evolucao do PDI</div><div class="progress-card"><div><div class="prog-pct" style="color:'+col+'">'+p.pdi_pct+'%</div><div class="prog-label">'+p.pdi_done+' de '+p.pdi_total+' aulas concluidas</div></div><div class="prog-bar-wrap"><div class="prog-bar-outer"><div class="prog-bar-inner" style="width:'+p.pdi_pct+'%;background:'+col+'"></div></div></div></div>';
  var months=Object.keys(p.consumed_by_month||{});
  var chart='<div class="sec-title">Consumo mensal de aulas</div><div class="chart-card">';
  if(months.length){
    var mx=1;months.forEach(function(m){if(p.consumed_by_month[m]>mx)mx=p.consumed_by_month[m];});
    chart+='<div class="chart-bars">';
    months.forEach(function(m){var v=p.consumed_by_month[m],h=Math.max(5,Math.round(v/mx*90));chart+='<div class="bar-col"><div class="bar-body" style="height:'+h+'px"><div class="bar-tip">'+v+' aulas</div></div><div class="bar-lbl">'+m.replace(/-/g,"/")+'</div></div>';});
    chart+='</div>';
  }else{chart+='<div class="no-chart">Sem historico de consumo registrado ainda</div>';}
  chart+='</div>';
  var lessons='<div class="sec-title">Aulas do PDI</div><div class="lessons-grid">';
  p.pdi_lessons.forEach(function(l){
    var ok=l.consumed;
    var lnk=l.link?'<a class="l-link" href="'+esc(l.link)+'" target="_blank">Acessar aula &rarr;</a>':"";
    lessons+='<div class="lesson '+( ok?"lesson-done":"lesson-pend")+'"><div class="l-icon '+( ok?"l-icon-done":"l-icon-pend")+'">'+( ok?"&#10003;":"&#9679;")+'</div><div><div class="l-title">'+esc(l.title)+'</div><div class="l-status '+( ok?"l-status-done":"l-status-pend")+'">'+( ok?"Concluida":"Pendente")+'</div>'+lnk+'</div></div>';
  });
  lessons+='</div>';
  var bgs='<div class="det-badges">';
  if(p.company)bgs+='<span class="badge badge-co">&#127970; '+esc(p.company)+'</span>';
  if(p.cargo)bgs+='<span class="badge badge-role">'+esc(p.cargo)+'</span>';
  if(p.email_cs)bgs+='<span class="badge badge-role">CS: '+esc(p.email_cs)+'</span>';
  bgs+='</div>';
  panel.innerHTML='<div class="det-name">'+esc(p.name)+'</div>'+bgs+kpi+prog+chart+lessons;
}
document.getElementById("s-name").addEventListener("input",filter);
document.getElementById("s-co").addEventListener("change",filter);
document.getElementById("s-pr").addEventListener("change",filter);
document.getElementById("reset-btn").addEventListener("click",reset);
document.getElementById("print-btn").addEventListener("click",function(){window.print();});
filter();
</script>
</body>
</html>"""

    html = HTML_TEMPLATE.replace("PLACEHOLDER_DATA", "var D = " + json_dados + ";")
    html = html.replace("PLACEHOLDER_DATE", 'var UPD = "' + agora + '";\n')

    with open(caminho_html, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"  HTML gerado: {OUTPUT_HTML}")
    return caminho_html


# ─────────────────────────────────────────
#  ENTRADA
# ─────────────────────────────────────────
if __name__ == '__main__':
    dados = processar()
    caminho = gerar_html(dados)
    print(f"\nPronto! Arquivo gerado: {os.path.basename(caminho)}")
    print(f"  {len(dados)} pessoas no dashboard")
    print("\nProximos passos:")
    print("  1. Abra o index.html no navegador para conferir")
    print("  2. Faca o commit/push para o GitHub para publicar")
    print("\nAte logo!")
