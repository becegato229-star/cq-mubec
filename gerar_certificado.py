"""
Gerador de Certificado de Qualidade - MUBEC
"""

import io, os
import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, Image, HRFlowable, KeepTogether
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH     = os.path.join(BASE_DIR, 'static', 'logo.png')
ASSIN_PATH    = os.path.join(BASE_DIR, 'static', 'assinatura.png')
PLANILHA_PATH = os.path.join(BASE_DIR, 'data', 'planilha_base.xlsx')

VERDE_MUBEC = colors.HexColor('#3DDC84')
PRETO       = colors.HexColor('#1A1A1A')
CINZA_ESCURO= colors.HexColor('#2D2D2D')
CINZA_MEDIO = colors.HexColor('#555555')
CINZA_CLARO = colors.HexColor('#F5F5F5')
CINZA_BORDA = colors.HexColor('#DDDDDD')
BRANCO      = colors.white

# Contador global para garantir nomes únicos de ParagraphStyle
_style_counter = 0
def _ps(**kw):
    global _style_counter
    _style_counter += 1
    d = dict(fontName='Helvetica', fontSize=9, leading=11, textColor=colors.HexColor('#2D2D2D'))
    d.update(kw)
    return ParagraphStyle(f'_s{_style_counter}', **d)


def carregar_dados_base():
    xls = pd.read_excel(PLANILHA_PATH, sheet_name=None, engine="openpyxl")
    # Sheet2 tem: Código, Nome, E-mail, Tipo, Estado, Município, CNPJ, ..., Telefone
    df_cli = xls.get('Sheet2', list(xls.values())[0])
    clientes = {}
    for _, row in df_cli.iterrows():
        cod_raw = row.get('Código', row.iloc[0])
        if pd.isna(cod_raw): continue
        cod = str(int(float(cod_raw)))
        cnpj = str(row.get('CNPJ', '') or '').strip()
        if cnpj in ('nan','0','') or (cnpj.isdigit() and len(cnpj) < 10):
            cnpj = ''
        clientes[cod] = {
            'nome':     str(row.get('Nome', '') or '').strip(),
            'cnpj':     cnpj,
            'email':    str(row.get('E-mail', '') or '').strip(),
            'telefone': str(row.get('Telefone', '') or '').strip(),
        }
    df_comp = xls.get('Composição', xls.get('Composicao', pd.DataFrame()))
    produtos = {}
    for _, row in df_comp.iterrows():
        cod = str(row.get('COD', '')).strip()
        if cod and cod != 'nan':
            produtos[cod] = {
                'desc':       str(row.get('Desc', '')).strip(),
                'composicao': row.get('Composição', row.get('Composicao', '')),
                'fpp':        row.get('FPP', ''),
                'kgf':        row.get('KGF', ''),
            }

    # Mapa: código alternativo (ERP) -> código real (Item) -> specs da Composição
    # Extraído da Sheet1 histórica
    df_hist = xls.get('Sheet1', pd.DataFrame())
    alt_para_item = {}
    for _, row in df_hist.iterrows():
        try:
            alt  = str(int(float(str(row.get('Alternativo do Item','')).split('.')[0])))
            item = str(int(float(str(row.get('Item','')).split('.')[0])))
            if alt and item and alt != 'nan' and item != 'nan':
                alt_para_item[alt] = item
        except: pass
    # Expande produtos com os aliases: alt_cod -> mesmos dados do item real
    for alt, item in alt_para_item.items():
        if item in produtos and alt not in produtos:
            produtos[alt] = produtos[item].copy()
    df_sup = xls.get('SUPORTE', pd.DataFrame())
    forns = []
    # bitola_specs: {Ø_primitivo_mm: {'fpp': X, 'kgf': Y}}
    # Lê os dois lados da tabela da aba SUPORTE (linhas 14+)
    # Lado esquerdo:  col0=bitola, col1=carga, col2=FPP,   col4=primitivo_mm
    # Lado direito:   col6=bitola, col7=carga, col8=passo,  col10=primitivo_mm
    bitola_specs = {}
    for _, row in df_sup.iterrows():
        n = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
        c = str(row.iloc[6]).strip() if pd.notna(row.iloc[6]) else ''
        if n and n != 'nan' and len(n) > 5 and '.' in c:
            forns.append({'nome': n, 'cnpj': c})
        # Lado esquerdo: UNC/WW
        try:
            prim = float(row.iloc[4])
            fpp  = float(row.iloc[2])
            kgf  = float(row.iloc[1])
            if prim > 0 and fpp > 0:
                # Guarda o maior FPP para esse diâmetro (UNC tem prioridade)
                if prim not in bitola_specs or fpp > bitola_specs[prim]['fpp']:
                    bitola_specs[prim] = {'fpp': fpp, 'kgf': round(kgf, 2)}
        except: pass
        # Lado direito: métrico (passo rosca no lugar de FPP)
        try:
            prim_m = float(row.iloc[10])
            passo  = float(row.iloc[8])
            kgf_m  = float(row.iloc[7])
            if prim_m > 0 and passo > 0 and prim_m not in bitola_specs:
                bitola_specs[prim_m] = {'fpp': passo, 'kgf': round(kgf_m, 2)}
        except: pass
    return clientes, produtos, forns[:3], \
           ['8 MICRA','13 MICRA','16 MICRA'], \
           ['AZUL','AMARELO','GALVANIZAÇÃO À FOGO'], \
           bitola_specs


def fmt(v):
    if v == '' or v is None: return '-'
    try:
        f = float(v)
        return f'{f:g}' if f != int(f) else str(int(f))
    except:
        return str(v)


def gerar_pdf(dados_nota: dict) -> bytes:
    buf = io.BytesIO()
    W = A4[0] - 30*mm - 20*mm   # 160 mm de largura útil

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=30*mm, rightMargin=20*mm,
        topMargin=30*mm,  bottomMargin=20*mm,
    )

    # ── Estilos base ─────────────────────────────────────────────────────────
    ROW_H = 7*mm   # altura fixa de todas as linhas de dados

    def lbl(txt):
        """Label de campo: negrito pequeno cinza."""
        return Paragraph(f'<b>{txt}</b>', _ps(fontSize=7, textColor=CINZA_MEDIO,
                                               fontName='Helvetica-Bold'))

    def val(txt, size=8, bold=False, align=TA_LEFT, color=None):
        """Valor de campo."""
        kw = dict(fontSize=size, textColor=color or PRETO,
                  alignment=align,
                  fontName='Helvetica-Bold' if bold else 'Helvetica')
        return Paragraph(str(txt), _ps(**kw))

    def hdr(txt):
        """Cabeçalho de coluna (fundo escuro, texto branco)."""
        return Paragraph(f'<b>{txt}</b>',
                         _ps(fontSize=7.5, fontName='Helvetica-Bold',
                             textColor=BRANCO, alignment=TA_CENTER))

    def cell(txt, size=8, align=TA_CENTER, bold=False):
        """Célula de dado na tabela de itens."""
        return Paragraph(str(txt),
                         _ps(fontSize=size, alignment=align,
                             fontName='Helvetica-Bold' if bold else 'Helvetica',
                             textColor=PRETO))

    def sec_title(txt):
        """Título de seção (fundo escuro, texto branco, bold)."""
        return Paragraph(f'<b>{txt}</b>',
                         _ps(fontSize=8, fontName='Helvetica-Bold',
                             textColor=BRANCO, alignment=TA_CENTER))

    story = []

    # ── Cabeçalho da empresa ──────────────────────────────────────────────────
    logo_w = 52*mm
    cert_w = 50*mm
    info_w = W - logo_w - cert_w

    logo = Image(LOGO_PATH, width=logo_w, height=logo_w*(533/2000))

    info_rows = [
        [Paragraph('<b>MUBEC IND. E COM. LTDA.</b>',
                   _ps(fontSize=9, fontName='Helvetica-Bold', textColor=PRETO))],
        [Paragraph('CNPJ: 00.604.905/0001-70  |  IE: 114.407.465.110',
                   _ps(fontSize=7, textColor=CINZA_MEDIO, fontName='Helvetica-Bold'))],
        [Paragraph('R. Murta do Campo, 705 – Vila Alpina – São Paulo/SP  03.210-010',
                   _ps(fontSize=7, textColor=CINZA_MEDIO, fontName='Helvetica-Bold'))],
        [Paragraph('(11) 2271-2900  |  qualidade@mubec.com.br',
                   _ps(fontSize=7, textColor=CINZA_MEDIO, fontName='Helvetica-Bold'))],
    ]
    tbl_info = Table(info_rows, colWidths=[info_w])
    tbl_info.setStyle(TableStyle([
        ('LEFTPADDING',   (0,0),(-1,-1), 6),
        ('RIGHTPADDING',  (0,0),(-1,-1), 0),
        ('TOPPADDING',    (0,0),(-1,-1), 1),
        ('BOTTOMPADDING', (0,0),(-1,-1), 1),
    ]))

    tbl_header = Table(
        [[logo, tbl_info,
          Paragraph('<b>CERTIFICADO<br/>DE QUALIDADE</b>',
                    _ps(fontSize=11, fontName='Helvetica-Bold',
                        textColor=VERDE_MUBEC, alignment=TA_RIGHT))]],
        colWidths=[logo_w, info_w, cert_w]
    )
    tbl_header.setStyle(TableStyle([
        ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
        ('LEFTPADDING',   (0,0),(-1,-1), 0),
        ('RIGHTPADDING',  (0,0),(-1,-1), 0),
        ('TOPPADDING',    (0,0),(-1,-1), 0),
        ('BOTTOMPADDING', (0,0),(-1,-1), 0),
    ]))
    story.append(tbl_header)
    story.append(Spacer(1, 3*mm))
    story.append(HRFlowable(width='100%', thickness=2.5, color=VERDE_MUBEC))
    story.append(Spacer(1, 3*mm))

    # ── Bloco Cliente ─────────────────────────────────────────────────────────
    nome_cli = dados_nota.get('nome_cliente', '')
    cnpj_cli = dados_nota.get('cnpj_cliente', '')
    tel_cli  = dados_nota.get('telefone_cliente', '')
    nf_num   = dados_nota.get('numero_nf', '')
    nf_data  = dados_nota.get('data_emissao', '')

    # colWidths somam W = 160 mm
    cw_cli = [W*0.50, W*0.30, W*0.20]

    bloco_cli = Table(
        [[lbl('CLIENTE'),        lbl('CNPJ'),     lbl('TELEFONE')],
         [val(nome_cli, size=8), val(cnpj_cli),   val(tel_cli)]],
        colWidths=cw_cli,
        rowHeights=[5*mm, ROW_H]
    )
    bloco_cli.setStyle(TableStyle([
        ('BOX',           (0,0),(-1,-1), 0.5, CINZA_BORDA),
        ('LINEBEFORE',    (1,0),(1,-1),  0.5, CINZA_BORDA),
        ('LINEBEFORE',    (2,0),(2,-1),  0.5, CINZA_BORDA),
        ('BACKGROUND',    (0,0),(-1,-1), CINZA_CLARO),
        ('LEFTPADDING',   (0,0),(-1,-1), 5),
        ('RIGHTPADDING',  (0,0),(-1,-1), 3),
        ('TOPPADDING',    (0,0),(-1,-1), 2),
        ('BOTTOMPADDING', (0,0),(-1,-1), 2),
        ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
    ]))
    story.append(KeepTogether(bloco_cli))
    story.append(Spacer(1, 1.5*mm))

    # ── Bloco NF ──────────────────────────────────────────────────────────────
    cw_nf = [W*0.30, W*0.30, W*0.40]

    bloco_nf = Table(
        [[lbl('NOTA FISCAL Nº'),   lbl('DATA DE EMISSÃO'), Paragraph('', _ps())],
         [val(nf_num, bold=True),  val(nf_data, bold=True),Paragraph('', _ps())]],
        colWidths=cw_nf,
        rowHeights=[5*mm, ROW_H]
    )
    bloco_nf.setStyle(TableStyle([
        ('BOX',           (0,0),(1,-1),  0.5, CINZA_BORDA),
        ('LINEBEFORE',    (1,0),(1,-1),  0.5, CINZA_BORDA),
        ('BACKGROUND',    (0,0),(1,-1),  CINZA_CLARO),
        ('LEFTPADDING',   (0,0),(-1,-1), 5),
        ('RIGHTPADDING',  (0,0),(-1,-1), 3),
        ('TOPPADDING',    (0,0),(-1,-1), 2),
        ('BOTTOMPADDING', (0,0),(-1,-1), 2),
        ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
    ]))
    story.append(KeepTogether(bloco_nf))
    story.append(Spacer(1, 3*mm))

    # ── Tabela de Itens ───────────────────────────────────────────────────────
    # colWidths somam W = 160 mm
    # 10+18+14+11+desc+20+18+22 = 113 => desc = W-113 = 47
    cw_it = [10*mm, 18*mm, 14*mm, 11*mm, W-113*mm, 20*mm, 18*mm, 22*mm]

    header_it = [hdr('ITEM'), hdr('CÓDIGO'), hdr('QUANT.'), hdr('UNID.'),
                 hdr('DESCRIÇÃO'), hdr('Ø (mm)'), hdr('PASSO/FPP'), hdr('CARGA (KGF)')]

    itens = dados_nota.get('itens', [])
    data_it = [header_it]
    for it in itens:
        data_it.append([
            cell(it.get('item', '')),
            cell(it.get('codigo', '')),
            cell(fmt(it.get('qtd', ''))),
            cell(it.get('unid', 'PC')),
            cell(it.get('descricao', ''), size=7.5, align=TA_LEFT),
            cell(fmt(it.get('fm', ''))),
            cell(fmt(it.get('fpp', ''))),
            cell(fmt(it.get('carga', ''))),
        ])

    # Linhas vazias com altura fixa
    for _ in range(15 - len(itens)):
        data_it.append([cell('')] * 8)

    # rowHeights: header fixo + todas as linhas iguais
    row_heights_it = [8*mm] + [ROW_H] * 15

    row_bg = [('BACKGROUND', (0,i),(-1,i), CINZA_CLARO if i%2==0 else BRANCO)
              for i in range(1, 16)]

    tbl_itens = Table(data_it, colWidths=cw_it, rowHeights=row_heights_it, repeatRows=1)
    tbl_itens.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,0),  PRETO),
        ('TEXTCOLOR',     (0,0),(-1,0),  BRANCO),
        ('BOX',           (0,0),(-1,-1), 0.5, CINZA_BORDA),
        ('INNERGRID',     (0,0),(-1,-1), 0.3, CINZA_BORDA),
        ('TOPPADDING',    (0,0),(-1,-1), 2),
        ('BOTTOMPADDING', (0,0),(-1,-1), 2),
        ('LEFTPADDING',   (0,0),(-1,-1), 3),
        ('RIGHTPADDING',  (0,0),(-1,-1), 3),
        ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
        *row_bg,
    ]))

    # Wrapper com título "ITENS DA NOTA FISCAL"
    tbl_it_wrap = Table(
        [[sec_title('ITENS DA NOTA FISCAL')],
         [tbl_itens]],
        colWidths=[W]
    )
    tbl_it_wrap.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(0,0), CINZA_ESCURO),
        ('TOPPADDING',    (0,0),(0,0), 4),
        ('BOTTOMPADDING', (0,0),(0,0), 4),
        ('TOPPADDING',    (0,1),(0,1), 0),
        ('BOTTOMPADDING', (0,1),(0,1), 0),
        ('LEFTPADDING',   (0,0),(-1,-1), 0),
        ('RIGHTPADDING',  (0,0),(-1,-1), 0),
    ]))
    story.append(KeepTogether(tbl_it_wrap))
    story.append(Spacer(1, 3*mm))

    # ── Composição Química ────────────────────────────────────────────────────
    maior_bitola = 0.0
    for it in itens:
        try:
            v = float(it.get('fm', 0) or 0)
            if v > maior_bitola: maior_bitola = v
        except: pass

    # Dados fixos das duas composições
    COMP_ATE23 = {
        'cols': ['C %','Si %','Mn %','S %','P %','Alt %','B %',
                 'Ca %','Cu %','Cr %','Mo %','N ppm','Nb %','Ni %','Sn %'],
        'vals': ['0,0500','0,09','0,40','0,021','0,022','0,002',
                 '0,0000','0,0014','0,01','0,01','0,002','26',
                 '0,000','0,01','0,001'],
        'nota': 'Composição referente a bitolas até 23,00 mm',
    }
    COMP_ACIMA23 = {
        'cols': ['%C','%Mn','%Si','%P','%S','%Nb',
                 '%Cu','%Cr','%Ni','%Sn','%Mo','%Al','%V'],
        'vals': ['0,21','0,59','0,15','0,019','0,027','0,000',
                 '0,27','0,10','0,08','0,016','0,018','0,009','0,000'],
        'nota': 'Composição referente a bitolas acima de 23,00 mm',
    }

    def build_comp_table(comp):
        n_c = len(comp['cols'])
        cw  = [W / n_c] * n_c
        tbl = Table(
            [[Paragraph(f'<b>{c}</b>', _ps(fontSize=7, fontName='Helvetica-Bold',
                        textColor=BRANCO, alignment=TA_CENTER)) for c in comp['cols']],
             [Paragraph(f'<b>{v}</b>', _ps(fontSize=8, fontName='Helvetica-Bold',
                        textColor=PRETO,  alignment=TA_CENTER)) for v in comp['vals']]],
            colWidths=cw,
            rowHeights=[6.5*mm, ROW_H]
        )
        tbl.setStyle(TableStyle([
            ('BACKGROUND',    (0,0),(-1,0),  CINZA_ESCURO),
            ('BACKGROUND',    (0,1),(-1,1),  CINZA_CLARO),
            ('BOX',           (0,0),(-1,-1), 0.5, CINZA_BORDA),
            ('INNERGRID',     (0,0),(-1,-1), 0.3, CINZA_BORDA),
            ('TOPPADDING',    (0,0),(-1,-1), 2),
            ('BOTTOMPADDING', (0,0),(-1,-1), 2),
            ('LEFTPADDING',   (0,0),(-1,-1), 1),
            ('RIGHTPADDING',  (0,0),(-1,-1), 1),
            ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
            ('LINEABOVE',     (0,0),(-1,0),  1, VERDE_MUBEC),
        ]))
        return tbl

    def build_comp_wrap(titulo, tbl, nota):
        wrap = Table(
            [[sec_title(titulo)],
             [tbl],
             [Paragraph(nota, _ps(fontSize=6.5, textColor=CINZA_MEDIO,
                                  alignment=TA_RIGHT, fontName='Helvetica-Oblique'))]],
            colWidths=[W]
        )
        wrap.setStyle(TableStyle([
            ('BACKGROUND',    (0,0),(0,0), CINZA_ESCURO),
            ('TOPPADDING',    (0,0),(0,0), 4),
            ('BOTTOMPADDING', (0,0),(0,0), 4),
            ('TOPPADDING',    (0,1),(0,1), 0),
            ('BOTTOMPADDING', (0,1),(0,1), 0),
            ('TOPPADDING',    (0,2),(0,2), 2),
            ('BOTTOMPADDING', (0,2),(0,2), 0),
            ('LEFTPADDING',   (0,0),(-1,-1), 0),
            ('RIGHTPADDING',  (0,0),(-1,-1), 0),
        ]))
        return wrap

    if maior_bitola > 0:
        if maior_bitola <= 23.0:
            story.append(KeepTogether(build_comp_wrap(
                'COMPOSIÇÃO QUÍMICA DA MATÉRIA-PRIMA',
                build_comp_table(COMP_ATE23),
                COMP_ATE23['nota']
            )))
        else:
            story.append(KeepTogether(build_comp_wrap(
                'COMPOSIÇÃO QUÍMICA DA MATÉRIA-PRIMA — ATÉ 23,00 mm',
                build_comp_table(COMP_ATE23),
                COMP_ATE23['nota']
            )))
            story.append(Spacer(1, 2*mm))
            story.append(KeepTogether(build_comp_wrap(
                'COMPOSIÇÃO QUÍMICA DA MATÉRIA-PRIMA — ACIMA DE 23,00 mm',
                build_comp_table(COMP_ACIMA23),
                COMP_ACIMA23['nota']
            )))
        story.append(Spacer(1, 3*mm))

    # ── Tratamento de Superfície ──────────────────────────────────────────────
    tem_galv   = dados_nota.get('tem_galvanizacao', False)
    galv_txt   = 'SIM' if tem_galv else 'NÃO'
    galv_cor   = VERDE_MUBEC if tem_galv else CINZA_ESCURO

    # Suporta múltiplos banhos: banhos = lista de {fornecedor, cnpj, passivacao, camada}
    # Para compatibilidade, se só tiver os campos simples monta lista de 1
    banhos = dados_nota.get('banhos', [])
    if not banhos and tem_galv:
        banhos = [{
            'fornecedor_galv': dados_nota.get('fornecedor_galv', ''),
            'cnpj_galv':       dados_nota.get('cnpj_galv', ''),
            'passivacao':      dados_nota.get('passivacao', ''),
            'camada':          dados_nota.get('camada', ''),
        }]

    col_a = 50*mm
    col_b = W - 90*mm
    col_c = 40*mm

    def linha_banho(banho, primeiro=False):
        """Monta uma linha da tabela para um banho."""
        forn  = banho.get('fornecedor_galv', '-') if tem_galv else '-'
        cnpj  = banho.get('cnpj_galv', '-')       if tem_galv else '-'
        passy = banho.get('passivacao', '-')       if tem_galv else '-'
        cam   = banho.get('camada', '-')           if tem_galv else '-'

        return [
            # Col A: GALVANIZAÇÃO / PASSIVAÇÃO
            Table([
                [Paragraph('<b>GALVANIZAÇÃO</b>' if primeiro else '',
                           _ps(fontSize=7, fontName='Helvetica-Bold', textColor=CINZA_MEDIO)),
                 Paragraph('<b>PASSIVAÇÃO</b>',
                           _ps(fontSize=7, fontName='Helvetica-Bold', textColor=CINZA_MEDIO))],
                [val(galv_txt if primeiro else '', bold=True, color=galv_cor if primeiro else None),
                 val(passy, bold=True)],
            ], colWidths=[col_a*0.5, col_a*0.5]),

            # Col B: FORNECEDOR / CNPJ
            Table([
                [Paragraph('<b>FORNECEDOR</b>',
                           _ps(fontSize=7, fontName='Helvetica-Bold', textColor=CINZA_MEDIO))],
                [val(forn, size=7.5)],
                [Paragraph('<b>CNPJ</b>',
                           _ps(fontSize=7, fontName='Helvetica-Bold', textColor=CINZA_MEDIO))],
                [val(cnpj, size=7.5)],
            ], colWidths=[col_b]),

            # Col C: CAMADA
            Table([
                [Paragraph('<b>CAMADA</b>',
                           _ps(fontSize=7, fontName='Helvetica-Bold', textColor=CINZA_MEDIO))],
                [val(cam, bold=True, align=TA_CENTER)],
            ], colWidths=[col_c]),
        ]

    if tem_galv and banhos:
        ts_rows = [linha_banho(b, primeiro=(i==0)) for i, b in enumerate(banhos)]
    else:
        ts_rows = [linha_banho({}, primeiro=True)]

    ts_style = [
        ('BOX',           (0,0),(-1,-1), 0.5, CINZA_BORDA),
        ('LINEBEFORE',    (1,0),(1,-1),  0.5, CINZA_BORDA),
        ('LINEBEFORE',    (2,0),(2,-1),  0.5, CINZA_BORDA),
        ('BACKGROUND',    (0,0),(-1,-1), CINZA_CLARO),
        ('TOPPADDING',    (0,0),(-1,-1), 5),
        ('BOTTOMPADDING', (0,0),(-1,-1), 5),
        ('LEFTPADDING',   (0,0),(-1,-1), 6),
        ('RIGHTPADDING',  (0,0),(-1,-1), 4),
        ('VALIGN',        (0,0),(-1,-1), 'TOP'),
        ('LINEABOVE',     (0,0),(-1,0),  1, VERDE_MUBEC),
    ]
    # Linha separadora entre banhos
    for i in range(1, len(ts_rows)):
        ts_style.append(('LINEABOVE', (0,i), (-1,i), 0.5, CINZA_BORDA))

    tbl_ts = Table(ts_rows, colWidths=[col_a, col_b, col_c])
    tbl_ts.setStyle(TableStyle(ts_style))

    tbl_ts_wrap = Table(
        [[sec_title('TRATAMENTO DE SUPERFÍCIE')],
         [tbl_ts]],
        colWidths=[W]
    )
    tbl_ts_wrap.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(0,0), CINZA_ESCURO),
        ('TOPPADDING',    (0,0),(0,0), 4),
        ('BOTTOMPADDING', (0,0),(0,0), 4),
        ('TOPPADDING',    (0,1),(0,1), 0),
        ('BOTTOMPADDING', (0,1),(0,1), 0),
        ('LEFTPADDING',   (0,0),(-1,-1), 0),
        ('RIGHTPADDING',  (0,0),(-1,-1), 0),
    ]))
    story.append(KeepTogether(tbl_ts_wrap))
    story.append(Spacer(1, 5*mm))

    # ── Declaração ────────────────────────────────────────────────────────────
    story.append(Paragraph(
        'Certifico o envio do produto acima, através da Nota Fiscal em referência. '
        'Material produzido e inspecionado de acordo com todas as exigências técnicas '
        'e especificações da norma NBR 6313.',
        _ps(fontSize=8, textColor=CINZA_ESCURO,
            alignment=TA_CENTER, fontName='Helvetica-Oblique')
    ))
    story.append(Spacer(1, 4*mm))

    # ── Assinatura centralizada ───────────────────────────────────────────────
    assin_w = 70*mm
    assin_h = assin_w * (261/1858)
    assin_img = Image(ASSIN_PATH, width=assin_w, height=assin_h)
    tbl_assin = Table([[assin_img]], colWidths=[W])
    tbl_assin.setStyle(TableStyle([
        ('ALIGN',         (0,0),(0,0), 'CENTER'),
        ('LEFTPADDING',   (0,0),(-1,-1), 0),
        ('RIGHTPADDING',  (0,0),(-1,-1), 0),
        ('TOPPADDING',    (0,0),(-1,-1), 0),
        ('BOTTOMPADDING', (0,0),(-1,-1), 0),
    ]))
    story.append(tbl_assin)

    # ── Rodapé ────────────────────────────────────────────────────────────────
    story.append(Spacer(1, 6*mm))
    story.append(HRFlowable(width='100%', thickness=1, color=CINZA_BORDA))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        f'Documento gerado em {datetime.now().strftime("%d/%m/%Y às %H:%M")}  |  '
        'MUBEC IND. E COM. LTDA. — qualidade@mubec.com.br — (11) 2271-2900',
        _ps(fontSize=7.5, textColor=CINZA_MEDIO,
            alignment=TA_CENTER, fontName='Helvetica-Oblique')
    ))

    doc.build(story)
    buf.seek(0)
    return buf.read()
