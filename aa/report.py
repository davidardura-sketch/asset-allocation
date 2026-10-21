"""Informe PDF con el contenido completo de señales y asignación."""
from datetime import datetime
from io import BytesIO
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                              PageBreak, KeepTogether)

NAVY=colors.HexColor('#17365D');TEAL=colors.HexColor('#246B70');INK=colors.HexColor('#162B40')
MUTED=colors.HexColor('#526579');PALE=colors.HexColor('#EAF0F5')


def num(v,d=2):
    return f'{v:,.{d}f}'.replace(',','_').replace('.',',').replace('_','.')


def pct(v,d=2):
    return num(v*100,d)+'%'


def export_pdf(r):
    stream=BytesIO();size=A4;width=size[0]-32*mm
    doc=SimpleDocTemplate(stream,pagesize=size,rightMargin=16*mm,leftMargin=16*mm,
                          topMargin=19*mm,bottomMargin=16*mm,title=f'Asset Allocation {r.month:%Y-%m}',
                          author='Asset Allocation',pageCompression=1)
    styles=getSampleStyleSheet()
    styles.add(ParagraphStyle(name='TitleAA',fontName='Helvetica-Bold',fontSize=21,leading=26,textColor=INK,spaceAfter=7))
    styles.add(ParagraphStyle(name='SubAA',fontName='Helvetica',fontSize=9,leading=13,textColor=MUTED,spaceAfter=11))
    styles.add(ParagraphStyle(name='HeadingAA',fontName='Helvetica-Bold',fontSize=11,leading=14,textColor=NAVY,spaceBefore=10,spaceAfter=6))
    styles.add(ParagraphStyle(name='CellAA',fontName='Helvetica',fontSize=8.2,leading=11,textColor=INK))
    styles.add(ParagraphStyle(name='HeadAA',fontName='Helvetica-Bold',fontSize=8,leading=10,textColor=colors.white))
    styles.add(ParagraphStyle(name='BodyAA',fontName='Helvetica',fontSize=9,leading=13,textColor=INK,spaceAfter=7))
    story=[]
    def para(t,style='BodyAA'):return Paragraph(escape(str(t)),styles[style])
    def title(t,sub):story.extend([para(t,'TitleAA'),para(sub,'SubAA')])
    def heading(t):story.append(para(t,'HeadingAA'))
    def table(headers,rows,widths=None):
        data=[[para(x,'HeadAA') for x in headers]]+[[para(v) for v in row] for row in rows]
        t=Table(data,colWidths=[width*x for x in widths] if widths else [width/len(headers)]*len(headers),repeatRows=1,hAlign='LEFT')
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),NAVY),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,PALE]),
                               ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),5),
                               ('BOTTOMPADDING',(0,0),(-1,-1),5),('LEFTPADDING',(0,0),(-1,-1),7),
                               ('LINEBELOW',(0,-1),(-1,-1),.5,colors.HexColor('#CFD9E3'))]))
        story.append(t)
    period=f'Asignación {r.month:%m/%Y} · cierre {r.cutoff:%d/%m/%Y} · moneda EUR · motor mixto v0.4'
    title('Asset Allocation',period)
    story.append(para(f'{r.position}   |   Score mixto {num(r.score,3)}   |   Duración RF {num(r.signals["Q48"])} años'))
    heading('Asignación por grandes bloques')
    table(['Bloque','Estratégico','Recomendado','Cambio (pp)'],
          [[b['name'],pct(b['strategic']),pct(b['recommended']),num(b['change']*100)] for b in r.blocks]+[['TOTAL','100,00%','100,00%','0,00']], [.4,.2,.2,.2])
    heading('Detalle de renta variable')
    table(['Región','Peso dentro RV','Peso cartera','Score mixto'],
          [[a['name'],pct(a['within']),pct(a['portfolio']),num(a['score'],3)] for a in r.regions],[.4,.2,.2,.2])
    heading('Detalle de renta fija y duración')
    table(['Segmento','Peso RF','Peso cartera','Score','Yield','OAS (bp)','OAD (años)','Contrib. duración cartera'],
          [[a['name'],pct(a['within']),pct(a['portfolio']),num(a['score'],3),num(a['yield_pct'])+'%',num(a['oas_bp']),num(a['duration']),num(a['duration_portfolio'])] for a in r.fixed_income],
          [.25,.1,.11,.09,.1,.1,.11,.14])
    story.append(Spacer(1,7));story.append(para('La duración procede de la composición de segmentos RF. No incorpora un overlay direccional de curva. El residual defensivo se reparte entre RF, liquidez y oro con las proporciones estratégicas.','SubAA'))
    story.append(PageBreak())
    s=r.signals
    title('Señales globales',period+' · scores acotados entre -1 y +1')
    heading('Tendencia de renta variable')
    table(['Métrica','Valor','Escala','Score','Peso','Contribución'],
          [[name,pct(s[f'B{i}']),pct(s[f'C{i}']),num(s[f'D{i}'],3),pct(s[f'E{i}'],0),num(s[f'F{i}'],3)]
           for name,i in [('Retorno 3 meses',6),('Retorno 6 meses',7),('Retorno 12 meses',8),('Amplitud sobre media 10m',9)]],[.3,.14,.14,.14,.14,.14])
    heading('Valoración de renta variable')
    table(['Earnings yield forward ponderado','EURIBOR 12m','ERP proxy','ERP neutral','Score valoración'],
          [[pct(s['B14']),pct(s['B15']),pct(s['B16']),pct(s['C16']),num(s['D16'],3)]])
    heading('Régimen de riesgo')
    table(['Métrica','Valor','Escala','Score','Peso','Contribución'],
          [[name,num(s[f'B{i}'],3) if i==20 else pct(s[f'B{i}']),num(s[f'C{i}'],3),num(s[f'D{i}'],3),pct(s[f'E{i}'],0),num(s[f'F{i}'],3)]
           for name,i in [('Volatilidad RV 6m / 36m',20),('Drawdown RV 12m',21),('HY / soberano, retorno relativo 6m',22)]],[.3,.14,.14,.14,.14,.14])
    heading('Combinación y referencias')
    table(['Señal','Score','Peso','Contribución'],[[name,num(s[f'D{i}'],3),pct(s[f'E{i}'],0),num(s[f'F{i}'],3)] for name,i in [('Tendencia',27),('Valoración',28),('Riesgo',29)]],[.4,.2,.2,.2])
    story.append(Spacer(1,8))
    table(['PER fwd pond.','P/B pond.','Vol. 90d pond.','Yield RF estr.','OAS RF estr.','Duración neutral','Duración táctica','Diferencia'],
          [[num(s['I6']),num(s['I7']),num(s['I8'])+'%',num(s['I11'])+'%',num(s['I12'])+' bp',num(s['I13'])+' a',num(s['I14'])+' a',num(s['I15'])+' a']])
    story.append(PageBreak())
    title('Señales regionales de renta variable',period)
    heading('Tendencia, valoración y riesgo')
    table(['Región','Ret. 6m','Ret. 12m','Trend','PER fwd','P/B','Valor','Vol. 90d','Riesgo'],
          [[a['name'],pct(a['return6']),pct(a['return12']),num(a['trend'],3),num(a['pe']),num(a['pb']),num(a['value'],3),pct(a['volatility']),num(a['risk'],3)] for a in r.regions],
          [.24,.1,.1,.09,.09,.09,.1,.1,.09])
    heading('Traducción del score a pesos')
    table(['Región','Peso estratégico RV','Score mixto','Factor tilt','Peso táctico RV','Peso cartera'],
          [[a['name'],pct(a['strategic']),num(a['score'],3),num(a['tilt'],4),pct(a['within']),pct(a['portfolio'])] for a in r.regions],[.27,.17,.14,.14,.14,.14])
    heading('Reglas de cálculo')
    for t in ['Trend regional: 40% de la señal a seis meses y 60% de la señal a doce meses.',
              'Valor regional: 60% del PER forward relativo y 40% del P/B relativo frente a sus medias estratégicas. Escala relativa: 30%.',
              'Riesgo regional: volatilidad 90d frente a la media ponderada estratégica, con escala relativa del 50%.',
              'Score regional: 45% tendencia, 40% valoración y 15% riesgo. El tilt máximo relativo es '+pct(r.params.region_tilt,0)+'. Los pesos se normalizan dentro de RV.',
              'Se mantienen las series de precio si la plantilla no incorpora versiones total return. La limitación se recoge en el control de datos.']:
        story.append(para(t))
    story.append(PageBreak())
    title('Señales de renta fija',period)
    heading('Tendencia, riesgo y analíticas fundamentales')
    table(['Segmento','Ret. 6m','Ret. 12m','Vol. 6m','Trend','Riesgo','Yield','OAS (bp)','OAD'],
          [[a['name'],pct(a['return6']),pct(a['return12']),pct(a['volatility']),num(a['trend'],3),num(a['risk'],3),num(a['yield_pct'])+'%',num(a['oas_bp']),num(a['duration'])] for a in r.fixed_income],
          [.24,.1,.1,.1,.09,.09,.1,.1,.08])
    heading('Fundamental, score y asignación')
    table(['Segmento','Peso estr. RF','Fundam.','Score mixto','Factor tilt','Peso táctico RF','Peso cartera'],
          [[a['name'],pct(a['strategic']),num(a['fundamental'],3),num(a['score'],3),num(a['tilt'],4),pct(a['within']),pct(a['portfolio'])] for a in r.fixed_income],[.25,.13,.12,.12,.12,.13,.13])
    heading('Contribuciones a la duración')
    table(['Segmento','Contribución duración RF (años)','Contribución duración cartera (años)'],
          [[a['name'],num(a['duration_rf'],4),num(a['duration_portfolio'],4)] for a in r.fixed_income]+[['TOTAL',num(s['Q48'],4),num(s['R48'],4)]],[.4,.3,.3])
    story.append(Spacer(1,9))
    story.append(para('Score RF: '+', '.join(pct(x,0) for x in r.params.bond_signal_weights)+' para tendencia, riesgo y fundamental, respectivamente. Fundamental: '+pct(r.params.bond_fundamental_weights[0],0)+' carry/Yield y '+pct(r.params.bond_fundamental_weights[1],0)+' OAS. Tilt máximo relativo: '+pct(r.params.bond_tilt,0)+'.'))
    story.append(PageBreak())
    title('Datos, controles y metodología',period)
    heading('Trazabilidad de la ejecución')
    story.append(para('Archivo: '+r.dataset.filename))
    story.append(para('SHA-256: '+r.dataset.digest,'SubAA'))
    story.append(para(f'Histórico RV: {r.dates[0]:%d/%m/%Y} a {r.dates[-1]:%d/%m/%Y}, {len(r.dates)} cierres. RF/caja comienza en el primer mes con datos propios.'))
    heading('Conciliaciones')
    table(['Control','Actual','Esperado','Diferencia','Estado'],[[c['control'],num(c['actual'],6),num(c['expected'],6),num(c['difference'],8),c['status']] for c in r.checks],[.4,.16,.16,.16,.12])
    heading('Limitaciones y convenciones aplicadas')
    for t in r.warnings:story.append(para(t))
    story.append(para(f'RV estratégica {pct(r.params.equity,0)}, limitada entre {pct(r.params.equity_min,0)} y {pct(r.params.equity_max,0)}. Una señal positiva usa el tramo hasta el máximo; una negativa, hasta el mínimo.'))
    story.append(para(f'FX: conversión a EUR sin cobertura. Liquidez: EURIBOR anual capitalizado con el tipo del mes anterior y días reales/{r.params.day_basis}.'))
    story.append(para('Se conserva la cesta de niveles rebajados a 100 y ponderados de v0.4. No se sustituye por un índice con rebalanceo mensual. No se estima rentabilidad esperada ni se realiza un backtest con el snapshot actual.'))
    def footer(canvas,doc):
        canvas.saveState();canvas.setStrokeColor(colors.HexColor('#D6DFE7'));canvas.line(16*mm,13*mm,size[0]-16*mm,13*mm)
        canvas.setFont('Helvetica',8);canvas.setFillColor(MUTED)
        canvas.drawString(16*mm,9*mm,'Asset Allocation · motor mixto v0.4')
        canvas.drawRightString(size[0]-16*mm,9*mm,f'{doc.page}');canvas.restoreState()
    doc.build(story,onFirstPage=footer,onLaterPages=footer)
    return stream.getvalue()
