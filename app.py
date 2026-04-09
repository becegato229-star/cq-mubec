"""
App Flask - Sistema de Certificado de Qualidade MUBEC
"""
import os, io, subprocess, tempfile
import pandas as pd
from flask import Flask, request, jsonify, send_file, render_template_string
from gerar_certificado import gerar_pdf, carregar_dados_base

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ERP_PATH = os.path.join(BASE_DIR, 'data', 'erp_atual.xlsx')

_dados_base = None
def get_dados_base():
    global _dados_base
    if _dados_base is None:
        _dados_base = carregar_dados_base()
    return _dados_base


def ler_excel_bytes(raw_bytes, filename=''):
    """Lê .xls ou .xlsx a partir de bytes."""
    if filename.lower().endswith('.xlsx'):
        return pd.read_excel(io.BytesIO(raw_bytes), engine='openpyxl')
    try:
        return pd.read_excel(io.BytesIO(raw_bytes), engine='xlrd')
    except Exception:
        pass
    # Converte com LibreOffice
    with tempfile.TemporaryDirectory() as tmpdir:
        xls_path = os.path.join(tmpdir, 'input.xls')
        with open(xls_path, 'wb') as f:
            f.write(raw_bytes)
        subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'xlsx', xls_path, '--outdir', tmpdir],
            check=True, capture_output=True
        )
        return pd.read_excel(os.path.join(tmpdir, 'input.xlsx'), engine='openpyxl')


def ler_excel_file(file_storage):
    return ler_excel_bytes(file_storage.read(), file_storage.filename or '')


def parse_erp_df(df):
    """Parseia DataFrame do ERP e retorna lista de notas."""
    df.columns = [str(c).strip() for c in df.columns]

    COL_NF    = 'Nº Nota Fiscal'
    COL_DATA  = 'Data de Emissão'
    COL_CLI_C = 'Cód. Cliente'
    COL_CLI_N = 'Nome do Cliente'
    COL_ITEM  = 'Item'
    COL_ALT   = 'Alternativo do Item'
    COL_DESC  = 'Descrição do Item'
    COL_UNID  = 'Unidade'
    COL_QTD   = 'Quantidade'

    clientes, produtos, _, _, _, bitola_specs = get_dados_base()
    notas = {}

    for _, row in df.iterrows():
        nf = str(row.get(COL_NF, '')).strip().split('.')[0]
        if not nf or nf in ('nan', '0', ''): continue

        if nf not in notas:
            data_raw = row.get(COL_DATA, '')
            if hasattr(data_raw, 'strftime'):
                data_fmt = data_raw.strftime('%d/%m/%Y')
            else:
                try:
                    data_fmt = pd.to_datetime(str(data_raw)).strftime('%d/%m/%Y')
                except:
                    data_fmt = str(data_raw)[:10]

            cli_cod  = str(int(float(row.get(COL_CLI_C, 0) or 0)))
            cli_nome = str(row.get(COL_CLI_N, '') or '').strip()
            cli_cnpj = ''
            cli_tel  = ''
            if cli_cod in clientes:
                cli_cnpj = clientes[cli_cod].get('cnpj', '')
                cli_tel  = clientes[cli_cod].get('telefone', '')

            notas[nf] = {
                'numero_nf':        nf,
                'data_emissao':     data_fmt,
                'cod_cliente':      cli_cod,
                'nome_cliente':     cli_nome,
                'cnpj_cliente':     cli_cnpj,
                'telefone_cliente': cli_tel,
                'itens': [],
                'tem_galvanizacao': True,
                'fornecedor_galv':  'JJ LESTE GALVANIZACAO LTDA',
                'cnpj_galv':        '26.412.069/0001-16',
                'passivacao':       'AMARELO',
                'camada':           '16 MICRA',
            }

        prod_cod = ''
        if COL_ALT in df.columns:
            v = str(row.get(COL_ALT, '') or '').strip().split('.')[0]
            if v and v != 'nan': prod_cod = v
        if not prod_cod:
            v = str(row.get(COL_ITEM, '') or '').strip().split('.')[0]
            if v and v != 'nan': prod_cod = v

        prod_desc = str(row.get(COL_DESC, '') or '').strip()
        qtd       = row.get(COL_QTD, '')
        unid      = str(row.get(COL_UNID, 'PC') or 'PC').strip()

        fm, fpp, carga = '', '', ''
        if prod_cod in produtos:
            p = produtos[prod_cod]
            fm    = p.get('composicao', '')
            fpp   = p.get('fpp', '')
            carga = p.get('kgf', '')
            if not prod_desc or prod_desc == 'nan':
                prod_desc = p.get('desc', '')
        # Se fpp ou carga vazios, busca na tabela de bitolas da aba SUPORTE pelo Ø (fm)
        if fm:
            try:
                spec = bitola_specs.get(float(fm))
                if spec:
                    if not fpp or str(fpp) in ('', 'nan', '0'):
                        fpp = spec['fpp']
                    if not carga or str(carga) in ('', 'nan', '0'):
                        carga = spec['kgf']
            except: pass

        notas[nf]['itens'].append({
            'item':      len(notas[nf]['itens']) + 1,
            'codigo':    prod_cod,
            'qtd':       qtd,
            'unid':      unid if unid != 'nan' else 'PC',
            'descricao': prod_desc if prod_desc != 'nan' else '',
            'fm':        fm,
            'fpp':       fpp,
            'carga':     carga,
        })

    return list(notas.values())


def carregar_erp_salvo():
    """Carrega o ERP que está salvo em data/erp_atual.xlsx (atualizado via GitHub)."""
    if os.path.exists(ERP_PATH):
        df = pd.read_excel(ERP_PATH, engine='openpyxl')
        return parse_erp_df(df)
    return []


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Certificado de Qualidade · MUBEC</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --verde: #3DDC84; --verde-dim: rgba(61,220,132,0.12);
    --preto: #0F0F0F; --900: #1A1A1A; --400: #888;
    --200: #E8E8E8; --100: #F5F5F5;
    --r: 10px; --shadow: 0 4px 24px rgba(0,0,0,0.08);
  }
  body { font-family:'DM Sans',sans-serif; background:#F0F2F5; color:var(--900); min-height:100vh; }
  .topbar { background:var(--preto); height:56px; display:flex; align-items:center; padding:0 32px; gap:16px; position:sticky; top:0; z-index:100; box-shadow:0 2px 12px rgba(0,0,0,.3); }
  .topbar img { height:22px; }
  .topbar-title { color:#fff; font-size:13px; font-weight:500; opacity:.6; }
  .sep { width:1px; height:20px; background:rgba(255,255,255,.15); }
  .badge { background:var(--verde-dim); color:var(--verde); font-size:11px; font-weight:700; padding:2px 10px; border-radius:20px; border:1px solid rgba(61,220,132,.3); margin-left:auto; }
  .layout { display:grid; grid-template-columns:320px 1fr; min-height:calc(100vh - 56px); }
  .sidebar { background:#fff; border-right:1px solid var(--200); display:flex; flex-direction:column; height:calc(100vh - 56px); overflow-y:auto; position:sticky; top:56px; }
  .s-sec { padding:16px; border-bottom:1px solid var(--200); }
  .s-sec:last-child { border:none; flex:1; }
  .s-label { font-size:10px; font-weight:700; letter-spacing:.1em; text-transform:uppercase; color:var(--400); margin-bottom:10px; }

  /* Tabs */
  .tabs { display:flex; gap:0; margin-bottom:12px; border:1.5px solid var(--200); border-radius:8px; overflow:hidden; }
  .tab { flex:1; padding:7px; font-size:11px; font-weight:700; text-align:center; cursor:pointer; background:#fff; color:var(--400); border:none; transition:.15s; }
  .tab.active { background:var(--preto); color:#fff; }

  .upload-zone { border:2px dashed var(--200); border-radius:var(--r); padding:18px; text-align:center; cursor:pointer; transition:.2s; background:var(--100); }
  .upload-zone:hover,.upload-zone.drag { border-color:var(--verde); background:var(--verde-dim); }
  .upload-icon { font-size:24px; margin-bottom:5px; }
  .upload-text { font-size:12px; color:var(--400); line-height:1.5; }
  .upload-text strong { color:var(--900); }
  #file-input { display:none; }

  .github-box { background:var(--100); border:1.5px solid var(--200); border-radius:var(--r); padding:12px; font-size:12px; }
  .github-box p { color:var(--400); line-height:1.6; margin-bottom:8px; }
  .github-box code { background:var(--200); padding:2px 6px; border-radius:4px; font-size:11px; font-family:'DM Mono',monospace; }
  .btn-load { width:100%; padding:9px; border-radius:8px; background:var(--verde); color:var(--preto); font-family:'DM Sans',sans-serif; font-size:13px; font-weight:700; border:none; cursor:pointer; margin-top:4px; transition:.15s; }
  .btn-load:hover { background:#2ec870; }
  .btn-load:disabled { opacity:.5; cursor:not-allowed; }

  .notes-list { display:flex; flex-direction:column; gap:6px; }
  .note-card { border:1.5px solid var(--200); border-radius:var(--r); padding:10px 12px; cursor:pointer; transition:.15s; }
  .note-card:hover,.note-card.active { border-color:var(--verde); background:var(--verde-dim); }
  .note-nf { font-size:12px; font-weight:700; font-family:'DM Mono',monospace; }
  .note-cli { font-size:11px; color:var(--400); margin-top:2px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
  .note-meta { display:flex; justify-content:space-between; margin-top:4px; font-size:10px; color:var(--400); font-family:'DM Mono',monospace; }
  .dot { width:6px; height:6px; border-radius:50%; background:var(--verde); display:inline-block; margin-right:4px; }

  .search-box { width:100%; padding:8px 10px; border:1.5px solid var(--200); border-radius:8px; font-family:'DM Sans',sans-serif; font-size:12px; outline:none; margin-bottom:10px; }
  .search-box:focus { border-color:var(--verde); }

  .main { padding:28px; overflow-y:auto; display:flex; flex-direction:column; gap:16px; }
  .empty { flex:1; display:flex; flex-direction:column; align-items:center; justify-content:center; gap:10px; color:var(--400); padding:80px 0; }
  .empty-icon { font-size:44px; opacity:.25; }
  .card { background:#fff; border-radius:var(--r); box-shadow:var(--shadow); overflow:hidden; }
  .card-hdr { padding:14px 18px; border-bottom:1px solid var(--200); display:flex; justify-content:space-between; align-items:center; }
  .card-title { font-size:13px; font-weight:700; }
  .card-body { padding:18px; }
  .grid2 { display:grid; grid-template-columns:1fr 1fr; gap:14px; }
  .grid3 { display:grid; grid-template-columns:1fr 1fr 1fr; gap:14px; }
  .full { grid-column:1/-1; }
  .field { display:flex; flex-direction:column; gap:4px; }
  label { font-size:10px; font-weight:700; color:var(--400); letter-spacing:.04em; text-transform:uppercase; }
  input,select { border:1.5px solid var(--200); border-radius:7px; padding:8px 10px; font-family:'DM Sans',sans-serif; font-size:13px; color:var(--900); background:#fff; outline:none; transition:.15s; width:100%; }
  input:focus,select:focus { border-color:var(--verde); }
  .toggle-row { display:flex; align-items:center; gap:10px; }
  .toggle { position:relative; width:42px; height:22px; }
  .toggle input { opacity:0; width:0; height:0; }
  .slider { position:absolute; inset:0; background:var(--200); border-radius:22px; cursor:pointer; transition:.2s; }
  .slider::before { content:''; position:absolute; height:16px; width:16px; left:3px; top:3px; background:#fff; border-radius:50%; transition:.2s; box-shadow:0 1px 4px rgba(0,0,0,.2); }
  .toggle input:checked+.slider { background:var(--verde); }
  .toggle input:checked+.slider::before { transform:translateX(20px); }
  .toggle-lbl { font-size:13px; font-weight:600; }
  .items-table { width:100%; border-collapse:collapse; font-size:12px; }
  .items-table th { background:var(--900); color:#fff; padding:7px 8px; text-align:left; font-size:10px; letter-spacing:.05em; font-weight:700; text-transform:uppercase; }
  .items-table td { padding:7px 8px; border-bottom:1px solid var(--200); vertical-align:middle; }
  .items-table tr:last-child td { border:none; }
  .items-table tr:hover td { background:var(--100); }
  .items-table td input { border:none; background:transparent; padding:0; font-size:12px; width:100%; outline:none; }
  .items-table td input:focus { border-bottom:1.5px solid var(--verde); }
  .tag { font-family:'DM Mono',monospace; font-size:11px; color:var(--verde); background:var(--verde-dim); padding:2px 8px; border-radius:4px; }
  .actions { display:flex; gap:10px; justify-content:flex-end; }
  .btn { padding:9px 20px; border-radius:8px; font-family:'DM Sans',sans-serif; font-size:13px; font-weight:700; cursor:pointer; border:none; transition:.15s; display:flex; align-items:center; gap:7px; }
  .btn-primary { background:var(--verde); color:var(--preto); }
  .btn-primary:hover { background:#2ec870; transform:translateY(-1px); box-shadow:0 4px 16px rgba(61,220,132,.4); }
  .btn-primary:disabled { opacity:.5; cursor:not-allowed; transform:none; box-shadow:none; }
  .btn-ghost { background:transparent; color:var(--400); border:1.5px solid var(--200); }
  .btn-ghost:hover { border-color:var(--400); color:var(--900); }
  .toast { position:fixed; bottom:20px; right:20px; padding:11px 18px; border-radius:10px; font-size:13px; font-weight:600; display:none; align-items:center; gap:8px; box-shadow:0 8px 32px rgba(0,0,0,.2); z-index:9999; }
  .toast.ok { background:var(--900); color:var(--verde); display:flex; }
  .toast.err { background:#FF4444; color:#fff; display:flex; }
  .spinner { width:15px; height:15px; border:2px solid rgba(0,0,0,.2); border-top-color:var(--preto); border-radius:50%; animation:spin .7s linear infinite; }
  @keyframes spin { to { transform:rotate(360deg); } }
</style>
</head>
<body>
<div class="topbar">
  <img src="/static/logo.png" alt="MUBEC">
  <div class="sep"></div>
  <span class="topbar-title">Certificado de Qualidade</span>
  <span class="badge">v1.0</span>
</div>
<div class="layout">
  <aside class="sidebar">
    <div class="s-sec">
      <div class="s-label">Carregar Notas</div>
      <div class="tabs">
        <button class="tab active" onclick="setTab('github')">📁 ERP do Dia</button>
        <button class="tab" onclick="setTab('upload')">⬆ Upload Manual</button>
      </div>

      <!-- Tab GitHub (ERP salvo no repo) -->
      <div id="tab-github">
        <div class="github-box">
          <p>O ERP do dia é carregado automaticamente do arquivo <code>data/erp_atual.xlsx</code> no GitHub.<br><br>
          Para atualizar: substitua esse arquivo no repositório e clique em <strong>Carregar</strong>.</p>
          <button class="btn-load" id="btn-carregar" onclick="carregarERP()">🔄 Carregar ERP do Dia</button>
        </div>
      </div>

      <!-- Tab Upload Manual -->
      <div id="tab-upload" style="display:none">
        <div class="upload-zone" id="zone" onclick="document.getElementById('file-input').click()">
          <div class="upload-icon">📊</div>
          <div class="upload-text"><strong>Clique ou arraste</strong> o Excel do ERP<br><span style="font-size:11px">.xls · .xlsx · até 32 MB</span></div>
        </div>
        <input type="file" id="file-input" accept=".xls,.xlsx" onchange="uploadERP(this)">
      </div>
    </div>

    <div class="s-sec">
      <div class="s-label" id="notes-count">Notas (0)</div>
      <input class="search-box" placeholder="🔍 Buscar por NF ou cliente..." oninput="filtrarNotas(this.value)">
      <div class="notes-list" id="notes-list">
        <div style="color:var(--400);font-size:12px;text-align:center;padding:16px 0">Carregue o ERP para listar as notas</div>
      </div>
    </div>
  </aside>

  <main class="main" id="main">
    <div class="empty"><div class="empty-icon">📋</div><div>Selecione uma nota para gerar o certificado</div></div>
  </main>
</div>
<div class="toast" id="toast"></div>

<script>
const FORNECEDORES = [
  { nome: 'JJ LESTE GALVANIZACAO LTDA',        cnpj: '26.412.069/0001-16' },
  { nome: 'GALVALLE INDUSTRIA E COMERCIO LTDA', cnpj: '12.882.845/0001-37' },
  { nome: 'ZINCAGEM MORIAH LTDA',               cnpj: '52.335.329/0001-07' },
];
let notas = [], notaAtual = null, notasFiltradas = [];

// ── Tabs ──────────────────────────────────────────────────────────────────

function setTab(tab) {
  document.querySelectorAll('.tab').forEach((t,i) => t.classList.toggle('active', (i===0&&tab==='github')||(i===1&&tab==='upload')));
  document.getElementById('tab-github').style.display = tab==='github' ? 'block' : 'none';
  document.getElementById('tab-upload').style.display = tab==='upload' ? 'block' : 'none';
}

// ── Carrega ERP do servidor (arquivo salvo no GitHub) ─────────────────────
async function carregarERP() {
  const btn = document.getElementById('btn-carregar');
  btn.disabled = true; btn.textContent = '⏳ Carregando...';
  try {
    const res = await fetch('/carregar-erp-salvo');
    const data = await res.json();
    if(data.error) toast(data.error, 'err');
    else { notas = data.notas; notasFiltradas = notas; renderNotes(); toast(`${notas.length} nota(s) carregada(s)!`, 'ok'); }
  } catch(e) { toast('Erro de conexão', 'err'); }
  btn.disabled = false; btn.textContent = '🔄 Carregar ERP do Dia';
}

// ── Upload manual ─────────────────────────────────────────────────────────
const zone = document.getElementById('zone');
zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag'); });
zone.addEventListener('dragleave', () => zone.classList.remove('drag'));
zone.addEventListener('drop', e => { e.preventDefault(); zone.classList.remove('drag'); if(e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]); });

function uploadERP(input) { if(input.files[0]) processFile(input.files[0]); }

async function processFile(file) {
  zone.innerHTML = '<div class="upload-icon">⏳</div><div class="upload-text">Processando...</div>';
  const form = new FormData(); form.append('file', file);
  try {
    const res = await fetch('/upload-erp', { method:'POST', body:form });
    const data = await res.json();
    if(data.error) toast(data.error, 'err');
    else { notas = data.notas; notasFiltradas = notas; renderNotes(); toast(`${notas.length} nota(s) carregada(s)!`, 'ok'); }
  } catch(e) { toast('Erro de conexão', 'err'); }
  zone.innerHTML = '<div class="upload-icon">📊</div><div class="upload-text"><strong>Clique ou arraste</strong> o Excel do ERP<br><span style="font-size:11px">.xls · .xlsx · até 32 MB</span></div>';
}

// ── Lista de notas ─────────────────────────────────────────────────────────
function filtrarNotas(q) {
  const ql = q.toLowerCase();
  notasFiltradas = notas.filter(n =>
    n.numero_nf.includes(q) || n.nome_cliente.toLowerCase().includes(ql)
  );
  renderNotes();
}

function renderNotes() {
  document.getElementById('notes-count').textContent = `Notas (${notasFiltradas.length}${notasFiltradas.length!==notas.length?' filtradas':''})`;
  document.getElementById('notes-list').innerHTML = notasFiltradas.map((n,i) => `
    <div class="note-card" id="card-${i}" onclick="selectNote('${n.numero_nf}')">
      <div class="note-nf"><span class="dot"></span>NF ${n.numero_nf}</div>
      <div class="note-cli">${n.nome_cliente||'—'}</div>
      <div class="note-meta"><span>${n.data_emissao||''}</span><span>${n.itens.length} item(s)</span></div>
    </div>`).join('');
}

function selectNote(nf) {
  document.querySelectorAll('.note-card').forEach(c=>c.classList.remove('active'));
  const idx = notasFiltradas.findIndex(n=>n.numero_nf===nf);
  document.getElementById(`card-${idx}`)?.classList.add('active');
  notaAtual = JSON.parse(JSON.stringify(notasFiltradas[idx]));
  renderEditor();
}

// ── Editor ────────────────────────────────────────────────────────────────
function renderEditor() {
  const n = notaAtual;
  const itemRows = n.itens.map((it,i) => `
    <tr>
      <td><span style="color:var(--400);font-size:11px">${it.item}</span></td>
      <td><input value="${it.codigo||''}" onchange="notaAtual.itens[${i}].codigo=this.value" style="font-family:'DM Mono',monospace;font-size:11px;color:var(--verde)"></td>
      <td><input value="${it.qtd||''}" onchange="notaAtual.itens[${i}].qtd=this.value" style="width:55px"></td>
      <td><input value="${it.unid||''}" onchange="notaAtual.itens[${i}].unid=this.value" style="width:38px"></td>
      <td><input value="${(it.descricao||'').replace(/"/g,'&quot;')}" onchange="notaAtual.itens[${i}].descricao=this.value"></td>
      <td><input value="${it.fm||''}" onchange="notaAtual.itens[${i}].fm=this.value" style="width:50px"></td>
      <td><input value="${it.fpp||''}" onchange="notaAtual.itens[${i}].fpp=this.value" style="width:45px"></td>
      <td><input value="${it.carga||''}" onchange="notaAtual.itens[${i}].carga=this.value" style="width:60px"></td>
    </tr>`).join('');

  document.getElementById('main').innerHTML = `
    <div class="card">
      <div class="card-hdr"><div class="card-title">Dados da Nota Fiscal</div><span class="tag">NF ${n.numero_nf}</span></div>
      <div class="card-body">
        <div class="grid3">
          <div class="field"><label>Nota Fiscal Nº</label><input value="${n.numero_nf}" onchange="notaAtual.numero_nf=this.value"></div>
          <div class="field"><label>Data de Emissão</label><input value="${n.data_emissao}" onchange="notaAtual.data_emissao=this.value"></div>
          <div class="field"><label>Telefone</label><input value="${n.telefone_cliente||''}" onchange="notaAtual.telefone_cliente=this.value"></div>
          <div class="field full"><label>Nome do Cliente</label><input value="${(n.nome_cliente||'').replace(/"/g,'&quot;')}" onchange="notaAtual.nome_cliente=this.value"></div>
          <div class="field"><label>CNPJ do Cliente</label><input value="${n.cnpj_cliente||''}" onchange="notaAtual.cnpj_cliente=this.value"></div>
        </div>
      </div>
    </div>
    <div class="card">
      <div class="card-hdr"><div class="card-title">Itens da Nota</div><span style="font-size:12px;color:var(--400)">${n.itens.length} item(s)</span></div>
      <div style="overflow-x:auto">
        <table class="items-table">
          <thead><tr><th>#</th><th>Código</th><th>Qtd</th><th>Und</th><th>Descrição</th><th>Ø mm</th><th>Passo</th><th>Carga KGF</th></tr></thead>
          <tbody>${itemRows}</tbody>
        </table>
      </div>
    </div>
    <div class="card">
      <div class="card-hdr"><div class="card-title">Tratamento de Superfície</div></div>
      <div class="card-body">
        <div style="margin-bottom:14px">
          <div class="toggle-row">
            <label class="toggle"><input type="checkbox" ${n.tem_galvanizacao?'checked':''} onchange="toggleGalv(this)"><span class="slider"></span></label>
            <span class="toggle-lbl">Galvanização</span>
          </div>
        </div>
        <div id="galv-fields" style="display:${n.tem_galvanizacao?'block':'none'}">
          <div id="banhos-lista">
            ${renderBanhos(n.banhos||[{fornecedor_galv:n.fornecedor_galv||'',cnpj_galv:n.cnpj_galv||'',passivacao:n.passivacao||'AMARELO',camada:n.camada||'16 MICRA'}])}
          </div>
          <button class="btn btn-ghost" style="margin-top:10px;font-size:12px" onclick="adicionarBanho()">+ Adicionar banho</button>
        </div>
      </div>
    </div>
    <div class="actions">
      <button class="btn btn-ghost" onclick="resetNote()">↩ Resetar</button>
      <button class="btn btn-primary" id="btn-gerar" onclick="gerarPDF()">📄 Gerar Certificado PDF</button>
    </div>`;
}

function renderBanhos(banhos) {
  // Garante que notaAtual.banhos está sincronizado
  notaAtual.banhos = banhos;
  return banhos.map((b, i) => `
    <div class="banho-item" style="border:1.5px solid var(--200);border-radius:8px;padding:12px;margin-bottom:8px;position:relative;">
      ${banhos.length > 1 ? `<button onclick="removerBanho(${i})" style="position:absolute;top:8px;right:8px;background:none;border:none;cursor:pointer;color:var(--400);font-size:16px;" title="Remover">×</button>` : ''}
      ${banhos.length > 1 ? `<div style="font-size:10px;font-weight:700;color:var(--400);letter-spacing:.08em;margin-bottom:8px;">BANHO ${i+1}</div>` : ''}
      <div class="grid2">
        <div class="field"><label>Fornecedor</label>
          <select onchange="atualizarBanho(${i},'fornecedor',this.value)">
            ${FORNECEDORES.map(f=>`<option value="${f.nome}" ${b.fornecedor_galv===f.nome?'selected':''}>${f.nome}</option>`).join('')}
          </select>
        </div>
        <div class="field"><label>CNPJ Fornecedor</label>
          <input id="cnpj-forn-${i}" value="${b.cnpj_galv||''}" readonly style="background:var(--100);color:var(--400)">
        </div>
        <div class="field"><label>Passivação</label>
          <select onchange="atualizarBanho(${i},'passivacao',this.value)">
            ${['AZUL','AMARELO','GALVANIZAÇÃO À FOGO'].map(p=>`<option ${b.passivacao===p?'selected':''}>${p}</option>`).join('')}
          </select>
        </div>
        <div class="field"><label>Camada</label>
          <select onchange="atualizarBanho(${i},'camada',this.value)">
            ${['8 MICRA','13 MICRA','16 MICRA','NBR 6313'].map(c=>`<option ${b.camada===c?'selected':''}>${c}</option>`).join('')}
          </select>
        </div>
      </div>
    </div>`).join('');
}

function atualizarBanho(idx, campo, valor) {
  if (!notaAtual.banhos) notaAtual.banhos = [];
  if (!notaAtual.banhos[idx]) notaAtual.banhos[idx] = {};
  if (campo === 'fornecedor') {
    const f = FORNECEDORES.find(x => x.nome === valor);
    notaAtual.banhos[idx].fornecedor_galv = valor;
    notaAtual.banhos[idx].cnpj_galv = f ? f.cnpj : '';
    const inp = document.getElementById(`cnpj-forn-${idx}`);
    if (inp) inp.value = notaAtual.banhos[idx].cnpj_galv;
  } else if (campo === 'passivacao') {
    notaAtual.banhos[idx].passivacao = valor;
  } else if (campo === 'camada') {
    notaAtual.banhos[idx].camada = valor;
  }
}

function adicionarBanho() {
  if (!notaAtual.banhos) notaAtual.banhos = [];
  notaAtual.banhos.push({
    fornecedor_galv: 'JJ LESTE GALVANIZACAO LTDA',
    cnpj_galv: '26.412.069/0001-16',
    passivacao: 'AMARELO',
    camada: '16 MICRA',
  });
  document.getElementById('banhos-lista').innerHTML = renderBanhos(notaAtual.banhos);
}

function removerBanho(idx) {
  notaAtual.banhos.splice(idx, 1);
  document.getElementById('banhos-lista').innerHTML = renderBanhos(notaAtual.banhos);
}

function toggleGalv(el) {
  notaAtual.tem_galvanizacao = el.checked;
  document.getElementById('galv-fields').style.display = el.checked?'block':'none';
}
function resetNote() {
  const orig = notas.find(n=>n.numero_nf===notaAtual.numero_nf);
  if(orig) { notaAtual=JSON.parse(JSON.stringify(orig)); renderEditor(); }
}
async function gerarPDF() {
  const btn = document.getElementById('btn-gerar');
  btn.disabled=true; btn.innerHTML='<div class="spinner"></div> Gerando...';
  try {
    const res = await fetch('/gerar-pdf',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(notaAtual)});
    if(!res.ok){const e=await res.json();toast(e.error||'Erro ao gerar PDF','err');}
    else {
      const blob=await res.blob();
      const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
      a.download=`CQ_${notaAtual.numero_nf}.pdf`; a.click();
      toast('PDF gerado!','ok');
    }
  } catch(e){toast('Erro de conexão','err');}
  btn.disabled=false; btn.innerHTML='📄 Gerar Certificado PDF';
}
function toast(msg,type){
  const t=document.getElementById('toast');
  t.textContent=(type==='ok'?'✓ ':'✗ ')+msg;
  t.className='toast '+type;
  setTimeout(()=>{t.style.display='none';t.className='toast';},3500);
}

// Carrega automaticamente ao abrir se tiver ERP salvo
window.onload = () => carregarERP();
</script>
</body>
</html>"""


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/static/<path:filename>')
def static_files(filename):
    return send_file(os.path.join(BASE_DIR, 'static', filename))

@app.route('/carregar-erp-salvo')
def carregar_erp_salvo_route():
    try:
        erp_path = ERP_PATH
        existe = os.path.exists(erp_path)
        tamanho = os.path.getsize(erp_path) if existe else 0
        print(f'[ERP] path={erp_path} existe={existe} tamanho={tamanho}')
        if not existe or tamanho == 0:
            return jsonify({'error': f'Arquivo erp_atual.xlsx não encontrado em {erp_path}. Suba o arquivo no GitHub em data/erp_atual.xlsx'}), 404
        notas = carregar_erp_salvo()
        if not notas:
            return jsonify({'error': 'Arquivo encontrado mas sem notas. Verifique o formato do Excel.'}), 400
        return jsonify({'notas': notas})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/upload-erp', methods=['POST'])
def upload_erp():
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400
    f = request.files['file']
    fname = f.filename or ''
    raw = f.read()
    print(f'[UPLOAD] arquivo={fname} tamanho={len(raw)}')
    try:
        # Tenta xlsx direto
        if fname.lower().endswith('.xlsx'):
            df = pd.read_excel(io.BytesIO(raw), engine='openpyxl')
        else:
            # Tenta xlrd para .xls
            try:
                df = pd.read_excel(io.BytesIO(raw), engine='xlrd')
            except Exception as e1:
                print(f'[UPLOAD] xlrd falhou: {e1}')
                # Tenta LibreOffice
                try:
                    with tempfile.TemporaryDirectory() as tmpdir:
                        xls_path = os.path.join(tmpdir, 'input.xls')
                        with open(xls_path, 'wb') as tmp:
                            tmp.write(raw)
                        subprocess.run(
                            ['libreoffice', '--headless', '--convert-to', 'xlsx',
                             xls_path, '--outdir', tmpdir],
                            check=True, capture_output=True, timeout=60
                        )
                        df = pd.read_excel(os.path.join(tmpdir, 'input.xlsx'), engine='openpyxl')
                except Exception as e2:
                    print(f'[UPLOAD] LibreOffice falhou: {e2}')
                    return jsonify({'error': f'Não foi possível ler o arquivo .xls. Converta para .xlsx no Excel e faça upload novamente.'}), 400
        notas = parse_erp_df(df)
        if not notas:
            return jsonify({'error': 'Nenhuma nota encontrada no arquivo.'}), 400
        print(f'[UPLOAD] {len(notas)} notas carregadas')
        return jsonify({'notas': notas})
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': f'Erro: {str(e)}'}), 400

@app.route('/gerar-pdf', methods=['POST'])
def gerar_pdf_route():
    dados = request.get_json()
    if not dados:
        return jsonify({'error': 'Dados inválidos'}), 400
    try:
        pdf_bytes = gerar_pdf(dados)
        return send_file(
            io.BytesIO(pdf_bytes), mimetype='application/pdf',
            as_attachment=True,
            download_name=f'CQ_{dados.get("numero_nf","cert")}.pdf'
        )
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5050))
    app.run(host='0.0.0.0', port=port, debug=False)
