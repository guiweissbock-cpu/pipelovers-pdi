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
    agora = datetime.now().strftime('%d/%m/%Y %H:%M')

    pasta_saida = os.path.dirname(os.path.abspath(__file__))
    caminho_html = os.path.join(pasta_saida, OUTPUT_HTML)

    print(f"\nGerando HTML...")

    with open(caminho_html, 'w', encoding='utf-8') as f:
        f.write('''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PipeLovers - Dashboard PDI</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=Inter:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
  :root {
    --bg:#0d0f14;--surface:#161921;--surface2:#1e2230;
    --border:#2a2f3d;--accent:#ff4f00;--accent2:#ff8c42;
    --green:#00c896;--yellow:#ffc857;--text:#e8eaf0;--muted:#7b8299;--radius:12px;
  }
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:"Inter",sans-serif;background:var(--bg);color:var(--text);min-height:100vh;}
  .header{background:var(--surface);border-bottom:1px solid var(--border);padding:18px 32px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;}
  .logo{display:flex;align-items:center;gap:12px;}
  .logo-icon{width:36px;height:36px;background:var(--accent);border-radius:8px;display:flex;align-items:center;justify-content:center;font-family:"Syne",sans-serif;font-weight:800;font-size:18px;color:white;}
  .logo-text{font-family:"Syne",sans-serif;font-weight:700;font-size:20px;}
  .logo-sub{font-size:12px;color:var(--muted);margin-top:1px;}
  .updated{font-size:12px;color:var(--muted);}
  .filters{padding:14px 32px;background:var(--surface);border-bottom:1px solid var(--border);display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end;}
  .filter-group{display:flex;flex-direction:column;gap:4px;}
  .filter-label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;font-weight:500;}
  select,input[type="text"]{background:var(--surface2);border:1px solid var(--border);color:var(--text);padding:8px 12px;border-radius:8px;font-size:13px;font-family:inherit;outline:none;cursor:pointer;min-width:160px;}
  select:focus,input:focus{border-color:var(--accent);}
  .btn{background:transparent;border:1px solid var(--border);color:var(--muted);padding:8px 16px;border-radius:8px;font-size:13px;cursor:pointer;font-family:inherit;transition:all .2s;}
  .btn:hover{border-color:var(--accent);color:var(--accent);}
  .filter-count{font-size:13px;color:var(--muted);padding-bottom:8px;}
  .main{display:flex;height:calc(100vh - 121px);}
  .list-panel{width:320px;flex-shrink:0;border-right:1px solid var(--border);overflow-y:auto;background:var(--surface);}
  .detail-panel{flex:1;overflow-y:auto;padding:28px 32px;}
  .list-header{padding:14px 20px 8px;font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;font-weight:600;border-bottom:1px solid var(--border);}
  .person-item{padding:13px 20px;cursor:pointer;border-bottom:1px solid var(--border);transition:background .15s;border-left:3px solid transparent;}
  .person-item:hover{background:var(--surface2);}
  .person-item.active{background:var(--surface2);border-left-color:var(--accent);}
  .person-name{font-size:14px;font-weight:500;margin-bottom:4px;}
  .person-meta{font-size:12px;color:var(--muted);display:flex;justify-content:space-between;align-items:center;}
  .pdi-mini{display:flex;align-items:center;gap:5px;}
  .pdi-mini-bar{width:46px;height:4px;background:var(--border);border-radius:2px;overflow:hidden;}
  .pdi-mini-fill{height:100%;border-radius:2px;}
  .stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:24px;}
  .stat-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:18px 20px;}
  .stat-label{font-size:11px;color:var(--muted);margin-bottom:8px;text-transform:uppercase;letter-spacing:.5px;}
  .stat-value{font-family:"Syne",sans-serif;font-size:30px;font-weight:700;line-height:1;}
  .detail-name{font-family:"Syne",sans-serif;font-size:26px;font-weight:800;margin-bottom:8px;}
  .detail-badges{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:24px;}
  .badge{padding:4px 10px;border-radius:20px;font-size:12px;font-weight:500;background:var(--surface2);border:1px solid var(--border);}
  .badge-company{color:var(--accent2);border-color:rgba(255,140,66,.3);background:rgba(255,140,66,.08);}
  .section-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:var(--muted);margin-bottom:14px;display:flex;align-items:center;gap:10px;}
  .section-title::after{content:"";flex:1;height:1px;background:var(--border);}
  .pdi-progress-wrap{display:flex;align-items:center;gap:20px;margin-bottom:24px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:20px 24px;}
  .pdi-big-pct{font-family:"Syne",sans-serif;font-size:44px;font-weight:800;line-height:1;white-space:nowrap;}
  .pdi-big-label{font-size:13px;color:var(--muted);margin-top:4px;}
  .pdi-bar-outer{flex:1;background:var(--surface2);border-radius:8px;height:14px;overflow:hidden;}
  .pdi-bar-inner{height:100%;border-radius:8px;transition:width .6s ease;}
  .lessons-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:24px;}
  .lesson-card{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:14px 16px;display:flex;align-items:flex-start;gap:10px;}
  .lesson-done{border-color:rgba(0,200,150,.25);background:rgba(0,200,150,.04);}
  .lesson-pending{opacity:.72;}
  .lesson-icon{width:22px;height:22px;border-radius:50%;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:11px;margin-top:1px;}
  .icon-done{background:rgba(0,200,150,.2);color:var(--green);}
  .icon-pending{background:var(--surface2);color:var(--muted);border:1px solid var(--border);}
  .lesson-title-text{font-size:13px;line-height:1.4;}
  .lesson-status{font-size:10px;font-weight:600;margin-top:4px;}
  .status-done{color:var(--green);}
  .status-pending{color:var(--muted);}
  .lesson-link{display:inline-block;margin-top:5px;font-size:11px;color:var(--accent2);text-decoration:none;}
  .lesson-link:hover{text-decoration:underline;}
  .chart-wrap{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:20px;margin-bottom:24px;}
  .chart-bars{display:flex;align-items:flex-end;gap:8px;height:110px;}
  .bar-group{flex:1;display:flex;flex-direction:column;align-items:center;gap:4px;min-width:0;}
  .bar-fill{width:100%;border-radius:4px 4px 0 0;min-height:3px;background:var(--accent);position:relative;cursor:default;}
  .bar-fill:hover{background:var(--accent2);}
  .bar-tip{position:absolute;bottom:calc(100% + 4px);left:50%;transform:translateX(-50%);background:var(--surface2);border:1px solid var(--border);padding:3px 7px;border-radius:4px;font-size:11px;white-space:nowrap;opacity:0;pointer-events:none;transition:opacity .15s;z-index:10;}
  .bar-fill:hover .bar-tip{opacity:1;}
  .bar-label{font-size:10px;color:var(--muted);text-align:center;white-space:nowrap;overflow:hidden;max-width:100%;}
  .no-data{text-align:center;color:var(--muted);font-size:14px;padding:30px;}
  .empty-state{text-align:center;padding:80px 20px;color:var(--muted);}
  .empty-icon{font-size:48px;margin-bottom:16px;}
  ::-webkit-scrollbar{width:5px;}::-webkit-scrollbar-track{background:transparent;}::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px;}
  @media print{.header,.filters,.list-panel{display:none!important;}.detail-panel{padding:16px;}.main{height:auto;}}
</style>
</head>
<body>
<div class="header">
  <div class="logo">
    <div class="logo-icon">P</div>
    <div>
      <div class="logo-text">PipeLovers</div>
      <div class="logo-sub">Dashboard de PDIs</div>
    </div>
  </div>
  <div style="display:flex;align-items:center;gap:16px;">
    <span class="updated" id="updated-label"></span>
    <button class="btn" id="print-btn">Imprimir / PDF</button>
  </div>
</div>
<div class="filters">
  <div class="filter-group">
    <div class="filter-label">Buscar pessoa</div>
    <input type="text" id="search-name" placeholder="Digite o nome...">
  </div>
  <div class="filter-group">
    <div class="filter-label">Empresa</div>
    <select id="filter-company"><option value="">Todas as empresas</option></select>
  </div>
  <div class="filter-group">
    <div class="filter-label">Progresso PDI</div>
    <select id="filter-progress">
      <option value="">Todos</option>
      <option value="not_started">Nao iniciado (0%)</option>
      <option value="in_progress">Em andamento (1-99%)</option>
      <option value="completed">Concluido (100%)</option>
    </select>
  </div>
  <button class="btn" id="reset-btn">Limpar filtros</button>
  <span class="filter-count" id="filter-count"></span>
</div>
<div class="main">
  <div class="list-panel">
    <div class="list-header">Colaboradores</div>
    <div id="person-list"></div>
  </div>
  <div class="detail-panel" id="detail-panel">
    <div class="empty-state">
      <div class="empty-icon">&#128072;</div>
      <div>Selecione um colaborador para ver o PDI</div>
    </div>
  </div>
</div>
<script>
''')

        f.write('var ALL_DATA = ')
        f.write(json_dados)
        f.write(';\n')
        f.write('var UPDATED_AT = "' + agora + '";\n')

        f.write('''
var filteredData = ALL_DATA.slice();
var selectedName = null;

document.getElementById("updated-label").textContent = "Atualizado em " + UPDATED_AT;

function barColor(pct) {
  if (pct === 100) return "#00c896";
  if (pct > 50)   return "#ffc857";
  if (pct > 0)    return "#ff8c42";
  return "#7b8299";
}
function esc(s) {
  return String(s || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

// Populate company dropdown
var companies = [];
ALL_DATA.forEach(function(d) { if (d.company && companies.indexOf(d.company) < 0) companies.push(d.company); });
companies.sort();
var sel = document.getElementById("filter-company");
companies.forEach(function(c) {
  var opt = document.createElement("option");
  opt.value = c; opt.textContent = c;
  sel.appendChild(opt);
});

function applyFilters() {
  var q   = document.getElementById("search-name").value.toLowerCase();
  var co  = document.getElementById("filter-company").value;
  var pr  = document.getElementById("filter-progress").value;
  filteredData = ALL_DATA.filter(function(d) {
    if (q  && d.name.toLowerCase().indexOf(q) < 0) return false;
    if (co && d.company !== co) return false;
    if (pr === "not_started" && d.pdi_pct !== 0) return false;
    if (pr === "in_progress" && (d.pdi_pct === 0 || d.pdi_pct === 100)) return false;
    if (pr === "completed"   && d.pdi_pct !== 100) return false;
    return true;
  });
  document.getElementById("filter-count").textContent = filteredData.length + " pessoas";
  renderList();
}

function resetFilters() {
  document.getElementById("search-name").value = "";
  document.getElementById("filter-company").value = "";
  document.getElementById("filter-progress").value = "";
  applyFilters();
}

function renderList() {
  var c = document.getElementById("person-list");
  c.innerHTML = "";
  if (!filteredData.length) {
    c.innerHTML = "<div class=\\"no-data\\">Nenhum resultado</div>";
    return;
  }
  filteredData.forEach(function(p) {
    var div = document.createElement("div");
    div.className = "person-item" + (selectedName === p.name ? " active" : "");
    var col = barColor(p.pdi_pct);
    div.innerHTML =
      "<div class=\\"person-name\\">" + esc(p.name) + "</div>" +
      "<div class=\\"person-meta\\">" +
        "<span style=\\"max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap\\">" + esc(p.company || "—") + "</span>" +
        "<div class=\\"pdi-mini\\">" +
          "<div class=\\"pdi-mini-bar\\"><div class=\\"pdi-mini-fill\\" style=\\"width:" + p.pdi_pct + "%;background:" + col + "\\"></div></div>" +
          "<span style=\\"font-size:11px;font-weight:600;color:" + col + "\\">" + p.pdi_pct + "%</span>" +
        "</div>" +
      "</div>";
    div.addEventListener("click", function() { selectPerson(p); });
    c.appendChild(div);
  });
}

function selectPerson(p) {
  selectedName = p.name;
  renderList();
  renderDetail(p);
}

function renderDetail(p) {
  var panel = document.getElementById("detail-panel");
  var col   = barColor(p.pdi_pct);

  var stats =
    "<div class=\\"stats-row\\">" +
      "<div class=\\"stat-card\\"><div class=\\"stat-label\\">Aulas assistidas (total)</div><div class=\\"stat-value\\" style=\\"color:var(--accent2)\\">" + p.total_consumed + "</div></div>" +
      "<div class=\\"stat-card\\"><div class=\\"stat-label\\">PDI concluido</div><div class=\\"stat-value\\" style=\\"color:var(--green)\\">" + p.pdi_done + " / " + p.pdi_total + "</div></div>" +
      "<div class=\\"stat-card\\"><div class=\\"stat-label\\">Evolucao PDI</div><div class=\\"stat-value\\" style=\\"color:" + col + "\\">" + p.pdi_pct + "%</div></div>" +
      "<div class=\\"stat-card\\"><div class=\\"stat-label\\">Aulas pendentes</div><div class=\\"stat-value\\" style=\\"color:var(--yellow)\\">" + (p.pdi_total - p.pdi_done) + "</div></div>" +
    "</div>";

  var prog =
    "<div class=\\"section-title\\">Progresso do PDI</div>" +
    "<div class=\\"pdi-progress-wrap\\">" +
      "<div>" +
        "<div class=\\"pdi-big-pct\\" style=\\"color:" + col + "\\">" + p.pdi_pct + "%</div>" +
        "<div class=\\"pdi-big-label\\">" + p.pdi_done + " de " + p.pdi_total + " aulas concluidas</div>" +
      "</div>" +
      "<div class=\\"pdi-bar-outer\\"><div class=\\"pdi-bar-inner\\" style=\\"width:" + p.pdi_pct + "%;background:" + col + "\\"></div></div>" +
    "</div>";

  var months = Object.keys(p.consumed_by_month || {});
  var chart  = "<div class=\\"section-title\\">Consumo mensal de aulas</div><div class=\\"chart-wrap\\">";
  if (months.length) {
    var mx = 1;
    months.forEach(function(m) { if (p.consumed_by_month[m] > mx) mx = p.consumed_by_month[m]; });
    chart += "<div class=\\"chart-bars\\">";
    months.forEach(function(m) {
      var v = p.consumed_by_month[m];
      var h = Math.max(4, Math.round(v / mx * 100));
      chart +=
        "<div class=\\"bar-group\\">" +
          "<div class=\\"bar-fill\\" style=\\"height:" + h + "px\\">" +
            "<div class=\\"bar-tip\\">" + v + " aulas</div>" +
          "</div>" +
          "<div class=\\"bar-label\\">" + m.replace(/-/g, "/") + "</div>" +
        "</div>";
    });
    chart += "</div>";
  } else {
    chart += "<div class=\\"no-data\\">Sem dados de consumo registrados</div>";
  }
  chart += "</div>";

  var lessons = "<div class=\\"section-title\\">Aulas do PDI</div><div class=\\"lessons-grid\\">";
  p.pdi_lessons.forEach(function(l) {
    var done = l.consumed;
    var link = l.link ? "<a class=\\"lesson-link\\" href=\\"" + esc(l.link) + "\\" target=\\"_blank\\">Acessar aula &rarr;</a>" : "";
    lessons +=
      "<div class=\\"lesson-card " + (done ? "lesson-done" : "lesson-pending") + "\\">" +
        "<div class=\\"lesson-icon " + (done ? "icon-done" : "icon-pending") + "\\">" + (done ? "&#10003;" : "&#9679;") + "</div>" +
        "<div>" +
          "<div class=\\"lesson-title-text\\">" + esc(l.title) + "</div>" +
          "<div class=\\"lesson-status " + (done ? "status-done" : "status-pending") + "\\">" + (done ? "Concluida" : "Pendente") + "</div>" +
          link +
        "</div>" +
      "</div>";
  });
  lessons += "</div>";

  var badges = "<div class=\\"detail-badges\\">";
  if (p.company) badges += "<span class=\\"badge badge-company\\">&#127970; " + esc(p.company) + "</span>";
  if (p.cargo)   badges += "<span class=\\"badge\\">" + esc(p.cargo) + "</span>";
  if (p.email_cs) badges += "<span class=\\"badge\\">CS: " + esc(p.email_cs) + "</span>";
  badges += "</div>";

  panel.innerHTML =
    "<div class=\\"detail-name\\">" + esc(p.name) + "</div>" +
    badges + stats + prog + chart + lessons;
}

document.getElementById("search-name").addEventListener("input", applyFilters);
document.getElementById("filter-company").addEventListener("change", applyFilters);
document.getElementById("filter-progress").addEventListener("change", applyFilters);
document.getElementById("reset-btn").addEventListener("click", resetFilters);
document.getElementById("print-btn").addEventListener("click", function() { window.print(); });

applyFilters();
''')
        f.write('</script>\n</body>\n</html>\n')

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
