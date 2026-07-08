import io, os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, Image, HRFlowable, KeepTogether)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

LOGO_PATH  = '/home/claude/logo.png'
ASSIN_PATH = '/home/claude/assinatura.png'
VERDE=colors.HexColor('#3DDC84'); PRETO=colors.HexColor('#1A1A1A')
C_ESC=colors.HexColor('#2D2D2D'); C_MED=colors.HexColor('#555555')
C_CLA=colors.HexColor('#F5F5F5'); C_BOR=colors.HexColor('#DDDDDD'); BRAN=colors.white

_c=[0]
def ps(**kw):
    _c[0]+=1; d=dict(fontName='Helvetica',fontSize=9,leading=11,textColor=C_ESC); d.update(kw)
    return ParagraphStyle(f'_s{_c[0]}',**d)

def lbl(t): return Paragraph(f'<b>{t}</b>',ps(fontSize=7,textColor=C_MED,fontName='Helvetica-Bold'))
def val(t,bold=False,size=8,align=TA_LEFT,color=None):
    return Paragraph(str(t),ps(fontSize=size,textColor=color or PRETO,alignment=align,fontName='Helvetica-Bold' if bold else 'Helvetica'))
def hdr(t): return Paragraph(f'<b>{t}</b>',ps(fontSize=7.5,fontName='Helvetica-Bold',textColor=BRAN,alignment=TA_CENTER))
def cell(t,size=8,align=TA_CENTER,bold=False):
    return Paragraph(str(t),ps(fontSize=size,alignment=align,fontName='Helvetica-Bold' if bold else 'Helvetica',textColor=PRETO))
def sec(t): return Paragraph(f'<b>{t}</b>',ps(fontSize=8,fontName='Helvetica-Bold',textColor=BRAN,alignment=TA_CENTER))
def fmt(v):
    if v=='' or v is None: return ''
    try:
        if str(v)=='nan': return ''
        f=float(v); return f'{f:g}' if f!=int(f) else str(int(f))
    except: return str(v)

# ── Composições por fornecedor ────────────────────────────────────────────────
# GERDAU (Certificado_1) — bitolas até ~11mm primitivo
COMP_GERDAU = {
    'fornecedor': 'GERDAU',
    'cols': ['%C', '%Mn', '%Si', '%P', '%S', '%Cu'],
    'vals': ['0,06', '0,45', '0,12', '0,017', '0,008', '0,01'],
    'nota': 'Composição referente ao Certificado de Qualidade GERDAU (FIO MÁQUINA)',
}
# SIMEC (Certificado_2) — bitolas a partir de ~12,7mm primitivo
COMP_SIMEC = {
    'fornecedor': 'SIMEC',
    'cols': ['%C', '%Mn', '%Si', '%P', '%S', '%Nb', '%Cu', '%Cr', '%Ni'],
    'vals': ['0,21', '0,59', '0,15', '0,019', '0,027', '0,000', '0,27', '0,10', '0,08'],
    'nota': 'Composição referente ao Certificado de Qualidade SIMEC (FIO MÁQUINA)',
}

# Mapeamento primitivo_mm -> {cert, forn} carregado da planilha_base
import pandas as pd
_bitola_cache = None
def get_bitola_info():
    global _bitola_cache
    if _bitola_cache is not None:
        return _bitola_cache
    planilha = '/home/claude/planilha_base.xlsx'
    xls = pd.read_excel(planilha, sheet_name=None, engine='openpyxl')
    df_sup = xls.get('SUPORTE', pd.DataFrame())
    info = {}
    for _, row in df_sup.iloc[14:].iterrows():
        # Lado esquerdo
        try:
            p=float(row.iloc[4]); f=float(row.iloc[2]); k=float(row.iloc[1])
            c=str(row.iloc[5]).strip(); fn=str(row.iloc[6]).strip()
            if p>0 and c not in ('nan',''):
                info[p]={'fpp':f,'kgf':round(k,2),'cert':c,'fornecedor':fn}
        except: pass
        # Lado direito (métrico)
        try:
            p=float(row.iloc[12]); f=float(row.iloc[10]); k=float(row.iloc[9])
            c=str(row.iloc[13]).strip(); fn=str(row.iloc[14]).strip()
            if p>0 and c not in ('nan','') and p not in info:
                info[p]={'fpp':f,'kgf':round(k,2),'cert':c,'fornecedor':fn}
        except: pass
    _bitola_cache = info
    return info

def get_comp_por_bitola(fm):
    """Retorna a composição correta baseada no primitivo (fm) do item."""
    try:
        prim = float(fm)
    except:
        return None
    info = get_bitola_info()
    bi = info.get(prim)
    if bi:
        cert = bi.get('cert','')
        if 'gerdau' in bi.get('fornecedor','').lower() or cert == 'Certificado_1':
            return COMP_GERDAU
        else:
            return COMP_SIMEC
    # Fallback por threshold
    return COMP_GERDAU if prim <= 11.0 else COMP_SIMEC

def gerar(dados):
    buf=io.BytesIO(); W=A4[0]-50*mm
    doc=SimpleDocTemplate(buf,pagesize=A4,leftMargin=30*mm,rightMargin=20*mm,topMargin=30*mm,bottomMargin=20*mm)
    ROW_H=7*mm; story=[]
    logo_w=52*mm; cert_w=50*mm; info_w=W-logo_w-cert_w
    logo=Image(LOGO_PATH,width=logo_w,height=logo_w*(533/2000))
    info_rows=[[Paragraph('<b>MUBEC IND. E COM. LTDA.</b>',ps(fontSize=9,fontName='Helvetica-Bold',textColor=PRETO))],
               [Paragraph('CNPJ: 00.604.905/0001-70  |  IE: 114.407.465.110',ps(fontSize=7,textColor=C_MED,fontName='Helvetica-Bold'))],
               [Paragraph('R. Murta do Campo, 705 – Vila Alpina – São Paulo/SP  03.210-010',ps(fontSize=7,textColor=C_MED,fontName='Helvetica-Bold'))],
               [Paragraph('(11) 2271-2900  |  qualidade@mubec.com.br',ps(fontSize=7,textColor=C_MED,fontName='Helvetica-Bold'))]]
    tbl_info=Table(info_rows,colWidths=[info_w])
    tbl_info.setStyle(TableStyle([('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),0),('TOPPADDING',(0,0),(-1,-1),1),('BOTTOMPADDING',(0,0),(-1,-1),1)]))
    tbl_hdr=Table([[logo,tbl_info,Paragraph('<b>CERTIFICADO<br/>DE QUALIDADE</b>',ps(fontSize=11,fontName='Helvetica-Bold',textColor=VERDE,alignment=TA_RIGHT))]],colWidths=[logo_w,info_w,cert_w])
    tbl_hdr.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),0)]))
    story.append(tbl_hdr); story.append(Spacer(1,3*mm))
    story.append(HRFlowable(width='100%',thickness=2.5,color=VERDE)); story.append(Spacer(1,3*mm))

    c1,c2,c3=W*0.50,W*0.30,W*0.20
    bloco_cli=Table([[lbl('CLIENTE'),lbl('CNPJ'),lbl('TELEFONE')],[val(dados['nome_cliente'],size=8),val(dados['cnpj_cliente']),val(dados['telefone_cliente'])]],colWidths=[c1,c2,c3],rowHeights=[5*mm,ROW_H])
    bloco_cli.setStyle(TableStyle([('BOX',(0,0),(-1,-1),0.5,C_BOR),('LINEBEFORE',(1,0),(1,-1),0.5,C_BOR),('LINEBEFORE',(2,0),(2,-1),0.5,C_BOR),('BACKGROUND',(0,0),(-1,-1),C_CLA),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),4),('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(KeepTogether(bloco_cli)); story.append(Spacer(1,1.5*mm))

    n1,n2,n3=W*0.30,W*0.30,W*0.40
    bloco_nf=Table([[lbl('NOTA FISCAL Nº'),lbl('DATA DE EMISSÃO'),Paragraph('',ps())],[val(dados['numero_nf'],bold=True),val(dados['data_emissao'],bold=True),Paragraph('',ps())]],colWidths=[n1,n2,n3],rowHeights=[5*mm,ROW_H])
    bloco_nf.setStyle(TableStyle([('BOX',(0,0),(1,-1),0.5,C_BOR),('LINEBEFORE',(1,0),(1,-1),0.5,C_BOR),('BACKGROUND',(0,0),(1,-1),C_CLA),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),4),('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(KeepTogether(bloco_nf)); story.append(Spacer(1,3*mm))

    cw_it=[10*mm,18*mm,14*mm,11*mm,W-113*mm,20*mm,18*mm,22*mm]
    itens=dados['itens']
    def build_items_wrap(titulo,subset):
        rows=[[hdr('ITEM'),hdr('CÓDIGO'),hdr('QUANT.'),hdr('UNID.'),hdr('DESCRIÇÃO'),hdr('Ø (mm)'),hdr('PASSO/FPP'),hdr('CARGA (KGF)')]]
        for it in subset:
            rows.append([cell(it['item']),cell(it['codigo']),cell(fmt(it['qtd'])),cell(it['unid']),cell(it['descricao'],size=7.5,align=TA_LEFT),cell(fmt(it['fm'])),cell(fmt(it['fpp'])),cell(fmt(it['carga']))])
        rh=[8*mm]+[ROW_H]*len(subset)
        rbg=[('BACKGROUND',(0,i),(-1,i),C_CLA if i%2==0 else BRAN) for i in range(1,len(subset)+1)]
        tbl=Table(rows,colWidths=cw_it,rowHeights=rh,repeatRows=1)
        tbl.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),PRETO),('TEXTCOLOR',(0,0),(-1,0),BRAN),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,0),7.5),('BOX',(0,0),(-1,-1),0.5,C_BOR),('INNERGRID',(0,0),(-1,-1),0.3,C_BOR),('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2),('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),('VALIGN',(0,0),(-1,-1),'MIDDLE'),*rbg]))
        wrap=Table([[sec(titulo)],[tbl]],colWidths=[W])
        wrap.setStyle(TableStyle([('BACKGROUND',(0,0),(0,0),C_ESC),('TOPPADDING',(0,0),(0,0),4),('BOTTOMPADDING',(0,0),(0,0),4),('TOPPADDING',(0,1),(0,1),0),('BOTTOMPADDING',(0,1),(0,1),0),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0)]))
        return wrap
    if len(itens)>15:
        meio=(len(itens)+1)//2
        story.append(KeepTogether(build_items_wrap('ITENS DA NOTA FISCAL',itens[:meio]))); story.append(Spacer(1,3*mm))
        story.append(KeepTogether(build_items_wrap('ITENS DA NOTA FISCAL (continuação)',itens[meio:])))
    else:
        story.append(KeepTogether(build_items_wrap('ITENS DA NOTA FISCAL',itens)))
    story.append(Spacer(1,3*mm))

    # ── Composição Química — por fornecedor/certificado ───────────────────────
    # Determina quais composições mostrar com base nos itens
    comps_usadas = {}  # cert -> comp
    for it in itens:
        fm = it.get('fm','')
        if not fm: continue
        comp = get_comp_por_bitola(fm)
        if comp:
            comps_usadas[comp['fornecedor']] = comp

    def build_comp(comp):
        n=len(comp['cols']); cw=[W/n]*n
        t=Table([[Paragraph(f'<b>{c}</b>',ps(fontSize=7,fontName='Helvetica-Bold',textColor=BRAN,alignment=TA_CENTER)) for c in comp['cols']],
                 [Paragraph(f'<b>{v}</b>',ps(fontSize=8,fontName='Helvetica-Bold',textColor=PRETO,alignment=TA_CENTER)) for v in comp['vals']]],
                colWidths=cw,rowHeights=[6.5*mm,ROW_H])
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),C_ESC),('BACKGROUND',(0,1),(-1,1),C_CLA),('BOX',(0,0),(-1,-1),0.5,C_BOR),('INNERGRID',(0,0),(-1,-1),0.3,C_BOR),('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2),('LEFTPADDING',(0,0),(-1,-1),1),('RIGHTPADDING',(0,0),(-1,-1),1),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LINEABOVE',(0,0),(-1,0),1,VERDE)]))
        return t

    def build_wrap_comp(titulo,tbl,nota):
        w=Table([[sec(titulo)],[tbl],[Paragraph(nota,ps(fontSize=6.5,textColor=C_MED,alignment=TA_RIGHT,fontName='Helvetica-Oblique'))]],colWidths=[W])
        w.setStyle(TableStyle([('BACKGROUND',(0,0),(0,0),C_ESC),('TOPPADDING',(0,0),(0,0),4),('BOTTOMPADDING',(0,0),(0,0),4),('TOPPADDING',(0,1),(0,1),0),('BOTTOMPADDING',(0,1),(0,1),0),('TOPPADDING',(0,2),(0,2),2),('BOTTOMPADDING',(0,2),(0,2),0),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0)]))
        return w

    # Ordem preferencial: GERDAU primeiro, SIMEC depois
    ordem = ['GERDAU','SIMEC']
    for forn in ordem:
        if forn in comps_usadas:
            comp = comps_usadas[forn]
            titulo = f'COMPOSIÇÃO QUÍMICA DA MATÉRIA-PRIMA — {forn}'
            story.append(KeepTogether(build_wrap_comp(titulo, build_comp(comp), comp['nota'])))
            story.append(Spacer(1,3*mm))

    # ── Tratamento de Superfície ──────────────────────────────────────────────
    tem_galv=dados.get('tem_galvanizacao',True)
    galv_cor=VERDE if tem_galv else C_ESC
    banhos=dados.get('banhos',[])
    col_a=50*mm; col_b=W-90*mm; col_c=40*mm

    def linha_banho(b,primeiro=False):
        forn=b.get('fornecedor_galv','-') if tem_galv else '-'
        cnpj=b.get('cnpj_galv','-') if tem_galv else '-'
        passy=b.get('passivacao','-') if tem_galv else '-'
        cam=b.get('camada','-') if tem_galv else '-'
        galv_txt='SIM' if tem_galv else 'NÃO'
        return [
            Table([[Paragraph('<b>GALVANIZAÇÃO</b>' if primeiro else '',ps(fontSize=7,fontName='Helvetica-Bold',textColor=C_MED)),Paragraph('<b>PASSIVAÇÃO</b>',ps(fontSize=7,fontName='Helvetica-Bold',textColor=C_MED))],[val(galv_txt if primeiro else '',bold=True,color=galv_cor if primeiro else None),val(passy,bold=True)]],colWidths=[col_a*0.5,col_a*0.5]),
            Table([[Paragraph('<b>FORNECEDOR</b>',ps(fontSize=7,fontName='Helvetica-Bold',textColor=C_MED))],[val(forn,size=7.5)],[Paragraph('<b>CNPJ</b>',ps(fontSize=7,fontName='Helvetica-Bold',textColor=C_MED))],[val(cnpj,size=7.5)]],colWidths=[col_b]),
            Table([[Paragraph('<b>CAMADA</b>',ps(fontSize=7,fontName='Helvetica-Bold',textColor=C_MED))],[val(cam,bold=True,align=TA_CENTER)]],colWidths=[col_c]),
        ]
    ts_rows=[linha_banho(b,primeiro=(i==0)) for i,b in enumerate(banhos)] if (tem_galv and banhos) else [linha_banho({},primeiro=True)]
    ts_style=[('BOX',(0,0),(-1,-1),0.5,C_BOR),('LINEBEFORE',(1,0),(1,-1),0.5,C_BOR),('LINEBEFORE',(2,0),(2,-1),0.5,C_BOR),('BACKGROUND',(0,0),(-1,-1),C_CLA),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),4),('VALIGN',(0,0),(-1,-1),'TOP'),('LINEABOVE',(0,0),(-1,0),1,VERDE)]
    for i in range(1,len(ts_rows)): ts_style.append(('LINEABOVE',(0,i),(-1,i),0.5,C_BOR))
    tbl_ts=Table(ts_rows,colWidths=[col_a,col_b,col_c])
    tbl_ts.setStyle(TableStyle(ts_style))
    wrap_ts=Table([[sec('TRATAMENTO DE SUPERFÍCIE')],[tbl_ts]],colWidths=[W])
    wrap_ts.setStyle(TableStyle([('BACKGROUND',(0,0),(0,0),C_ESC),('TOPPADDING',(0,0),(0,0),4),('BOTTOMPADDING',(0,0),(0,0),4),('TOPPADDING',(0,1),(0,1),0),('BOTTOMPADDING',(0,1),(0,1),0),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0)]))
    story.append(KeepTogether(wrap_ts)); story.append(Spacer(1,5*mm))

    story.append(Paragraph('Certifico o envio do produto acima, através da Nota Fiscal em referência. Material produzido e inspecionado de acordo com todas as exigências técnicas e especificações da norma NBR 6313.',ps(fontSize=8,textColor=C_ESC,alignment=TA_CENTER,fontName='Helvetica-Oblique')))
    story.append(Spacer(1,4*mm))
    assin_w=70*mm; assin_h=assin_w*(261/1858)
    assin=Image(ASSIN_PATH,width=assin_w,height=assin_h)
    tbl_a=Table([[assin]],colWidths=[W])
    tbl_a.setStyle(TableStyle([('ALIGN',(0,0),(0,0),'CENTER'),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),0)]))
    story.append(tbl_a)
    story.append(Spacer(1,6*mm))
    story.append(HRFlowable(width='100%',thickness=1,color=C_BOR)); story.append(Spacer(1,2*mm))
    story.append(Paragraph(f'Documento gerado em {datetime.now().strftime("%d/%m/%Y às %H:%M")}  |  MUBEC IND. E COM. LTDA. — qualidade@mubec.com.br — (11) 2271-2900',ps(fontSize=7.5,textColor=C_MED,alignment=TA_CENTER,fontName='Helvetica-Oblique')))
    doc.build(story); buf.seek(0)
    return buf.read()
