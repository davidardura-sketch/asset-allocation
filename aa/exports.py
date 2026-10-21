"""Exportación portátil: no necesita Excel, Bloomberg ni Node en Streamlit."""
from __future__ import annotations

from dataclasses import asdict
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
import json
import math
import re

import xlsxwriter
from openpyxl.formula.translate import Translator

from .model import Result

TEMPLATE = Path(__file__).parent / 'templates' / 'v04.json'


def _parameter_cells(p):
    vals = { 'B8':p.equity, 'B9':p.equity_min, 'B10':p.equity_max, 'B11':p.bonds,
             'B12':p.cash, 'B13':p.gold, 'B25':p.erp_neutral, 'B26':p.valuation_scale,
             'B30':p.volatility_scale, 'B31':p.drawdown_scale, 'B32':p.credit_scale,
             'B33':p.region_tilt, 'B34':p.bond_tilt }
    for col,rows,values in [('E',range(6,11),p.equity_regions),('H',range(6,10),p.bond_segments),
                            ('B',range(16,19),p.signal_weights),('B',range(19,23),p.trend_weights),
                            ('B',range(27,30),p.trend_scales),('B',range(35,38),p.defensive),
                            ('E',range(21,24),p.bond_signal_weights),('E',range(24,26),p.bond_fundamental_weights),
                            ('E',range(26,28),p.bond_fundamental_scales)]:
        vals.update({f'{col}{r}':float(v) for r,v in zip(rows,values)})
    return vals


def export_excel(result: Result) -> bytes:
    """Libro autónomo con fórmulas estándar y valores calculados guardados."""
    r=result; p=r.params; n=len(r.dates); last=n+5
    template=json.loads(TEMPLATE.read_text(encoding='utf8'))
    buffer=BytesIO()
    book=xlsxwriter.Workbook(buffer, {'in_memory':True,'strings_to_formulas':False,'strings_to_urls':False})
    book.set_properties({'title':f'Asset Allocation {r.month:%Y-%m}', 'subject':'Motor mixto v0.4',
                         'comments':'Resultados calculados en Python. Las fórmulas Excel se recalculan al abrir.'})
    book.set_calc_mode('auto')
    fmt={
        'body':book.add_format({'font_name':'Arial','font_size':10,'valign':'vcenter'}),
        'title':book.add_format({'font_name':'Arial','font_size':15,'bold':True,'font_color':'162B40'}),
        'header':book.add_format({'font_name':'Arial','font_size':10,'bold':True,'bg_color':'17365D','font_color':'FFFFFF','text_wrap':True,'valign':'vcenter'}),
        'total':book.add_format({'font_name':'Arial','font_size':10,'bold':True,'bg_color':'E7EFF4','num_format':'0.00%;(0.00%);"-"'}),
        'pct':book.add_format({'font_name':'Arial','font_size':10,'num_format':'0.00%;(0.00%);"-"'}),
        'num':book.add_format({'font_name':'Arial','font_size':10,'num_format':'0.0000;(0.0000);"-"'}),
        'unitpct':book.add_format({'font_name':'Arial','font_size':10,'num_format':'0.00"%"'}),
        'date':book.add_format({'font_name':'Arial','font_size':10,'num_format':'dd/mm/yyyy'}),
        'input':book.add_format({'font_name':'Arial','font_size':10,'bg_color':'FFF2CC','font_color':'0000FF','num_format':'0.0000'}),
        'note':book.add_format({'font_name':'Arial','font_size':10,'font_color':'526579','text_wrap':True,'valign':'top'}),
    }
    names=['12_Asignacion','10_Parametros','11_Senales','06_Series_EUR','10_Series_Motor',
           '02_BBG_Snapshot','Datos','00_Control','Calidad']
    sheets={name:book.add_worksheet(name) for name in names}
    for name,s in sheets.items():
        s.hide_gridlines(2);s.set_default_row(20);s.set_column('A:A',32,fmt['body']);s.set_column('B:R',16,fmt['body'])
        s.set_landscape();s.fit_to_pages(1,0);s.set_paper(9)
        s.set_footer('&LAsset Allocation v0.4&R&P / &N')
        if name in ('Datos','06_Series_EUR','10_Series_Motor'):
            s.freeze_panes(5,1);s.set_column('A:A',15,fmt['date'])
        s.set_tab_color('17365D' if name=='12_Asignacion' else '527B94' if name=='10_Parametros' else 'DCE5EC')

    def write(s,a,v,style=None):
        if isinstance(v,(float,int)) and not isinstance(v,bool):
            if math.isfinite(v):s.write_number(a,float(v),style or fmt['num'])
            else:s.write_blank(a,None,style or fmt['num'])
        elif isinstance(v,date):s.write_datetime(a,datetime.combine(v,datetime.min.time()),fmt['date'])
        elif v is None:s.write_blank(a,None,style or fmt['body'])
        else:s.write_string(a,str(v),style or fmt['body'])

    def formula(s,a,f,v,style=None):
        if isinstance(v,float) and not math.isfinite(v):v=''
        if isinstance(v,date):v=(v-date(1899,12,30)).days
        s.write_formula(a,f,style or fmt['num'],v)

    # Se conserva la estructura de las dos pestañas de salida del motor v0.4.
    for name in ('10_Parametros','11_Senales','12_Asignacion'):
        s=sheets[name]
        header_rows=({'11_Senales':(4,5,12,13,18,19,25,26,33,34,42,43),
                      '12_Asignacion':(7,8,16,17,26,27,35,36),
                      '10_Parametros':(4,15,24)})[name]
        for rr in header_rows:
            for cc in range(18 if name=='11_Senales' else 8):s.write_blank(rr-1,cc,None,fmt['header'])
        for a,v in template[name].items():
            if not isinstance(v,str) or not v.startswith('='):
                rr=int(re.search(r'\d+',a)[0])
                write(s,a,v,fmt['header'] if rr in header_rows else None)
        s.set_row(0,27)
        write(s,'A1',template[name]['A1'],fmt['title'])
        s.merge_range('A2:R3',template[name].get('A2',''),fmt['note'])
        if name=='11_Senales':
            for rr in (4,5,12,13,18,19,25,26,33,34,42,43):
                s.set_row(rr-1,32 if rr in (34,43) else 23,fmt['header'])
            s.set_column('A:A',34);s.set_column('B:B',28)
        elif name=='12_Asignacion':
            for rr in (7,8,16,17,26,27,35,36):s.set_row(rr-1,30,fmt['header'])
            s.set_column('A:A',31);s.set_column('B:D',17);s.set_column('E:H',21)
            for a in ('A33','A49'):s.set_row(int(a[1:])-1,42)
        else:
            s.set_column('A:A',35);s.set_column('D:D',34);s.set_column('E:E',24)
            s.set_column('G:G',28);s.set_column('H:H',17)
            for rr in (4,15,24):s.set_row(rr-1,26,fmt['header'])
            for rr in (13,14,15,16,17,18,30,31,32):
                if rr in (15,24):continue
                if template[name].get(f'E{rr}') and isinstance(template[name][f'E{rr}'],str):
                    text=template[name][f'E{rr}']
                    s.merge_range(f'E{rr}:H{rr}',text,fmt['note']);s.set_row(rr-1,36)
    ps=sheets['10_Parametros']
    for a,v in _parameter_cells(p).items():write(ps,a,v,fmt['input'])
    formula(ps,'B6',"='00_Control'!B6",r.cutoff,fmt['date'])
    # Valores que en julio eran constantes pero ahora son configurables.
    ss=sheets['11_Senales']
    for row,weight in zip((20,21,22),p.risk_weights):write(ss,f'E{row}',weight,fmt['input'])
    write(ss,'A49','La duración procede de la composición RF. El modelo no incorpora un overlay de curva.',fmt['note'])
    for name,caches in [('11_Senales',r.signals),('12_Asignacion',r.allocation_cells)]:
        s=sheets[name]
        for a,f in template[name].items():
            if not isinstance(f,str) or not f.startswith('='):continue
            def shift_reference(m):
                shifted=re.sub(r'([A-Z]+)(\d+)',lambda c:f'{c[1]}{int(c[2])+last-324}',m[2])
                return m[1]+'!'+shifted
            f=re.sub(r"('06_Series_EUR'|'10_Series_Motor')!([A-Z]+\d+(?::[A-Z]+\d+)?)",shift_reference,f)
            if a=='J5' and name=='12_Asignacion':
                write(s,a,caches[a]);continue
            if name=='11_Senales' and a=='I12' and p.oas_unit=='bp':f=f.replace('100*(', '1*(')
            if name=='11_Senales' and a in ('J44','J45','J46','J47') and p.oas_unit=='bp':f=f.replace('*100','*1')
            if a not in caches:raise RuntimeError(f'Falta caché: {name}!{a}')
            style=fmt['pct'] if name=='12_Asignacion' and (re.match(r'[BC](9|10|11|12|13|18|19|20|21|22|23|28|29|30|31|32)$',a)) else fmt['num']
            if name=='11_Senales' and (a=='B31' or re.match(r'E(6|7|8|9|20|21|22|27|28|29)$',a)):
                style=fmt['pct']
            formula(s,a,f,caches[a],style)
        for a,v in caches.items():
            if a not in template[name] or not str(template[name][a]).startswith('='):
                write(s,a,v)
    a=sheets['12_Asignacion']
    write(a,'A1',f'Asset Allocation · {r.month:%m/%Y} · v0.4',fmt['title'])
    for addr,ref in [('B9','B12'),('B10','B11'),('B11','B8'),('B12','B13')]:formula(a,addr,f"='10_Parametros'!{ref}",r.allocation_cells[addr],fmt['pct'])
    formula(a,'B13','=SUM(B9:B12)',1.,fmt['total'])
    formula(a,'D13','=ROUND(C13-B13,10)',0.,fmt['total'])
    for rr in range(9,13):formula(a,f'D{rr}',f'=(C{rr}-B{rr})*100',r.allocation_cells[f'D{rr}']*100,fmt['num'])
    formula(a,'B5',"='00_Control'!B6",r.cutoff,fmt['date'])
    for addr in ('F5','B23','C23','B32','C32','C13'):formula(a,addr,template['12_Asignacion'].get(addr,'=1'),r.allocation_cells[addr],fmt['pct'])
    for rr in range(28,33):
        # Yield es un dato Bloomberg expresado en puntos porcentuales.
        f=template['12_Asignacion'][f'E{rr}'];formula(a,f'E{rr}',f,r.allocation_cells[f'E{rr}'],fmt['unitpct'])
    a.merge_range('A41:H42','Exportación mensual calculada por la aplicación. Cargar el próximo Excel en la aplicación para actualizar datos y PDF.',fmt['note'])
    a.merge_range('A44:H46','Limitaciones: '+' '.join(r.warnings),fmt['note'])
    a.set_print_scale(80);a.print_area('A1:N46')
    # Datos originales alineados por mes. Se conservan también las fechas originales RF y FX.
    ds=sheets['Datos'];write(ds,'A1','Datos Bloomberg de entrada',fmt['title'])
    ds.merge_range('A2:P3',f'Fuente: {r.dataset.filename}. SHA-256: {r.dataset.digest}',fmt['note'])
    headers=['Fecha RV','SPXT','SXXR','MXJP','MXPCJ','MXEF','XAU','Soberano','Cédulas','IG','HY','Tipo caja %','EURUSD','EURJPY','Fecha RF','Fecha FX']
    ds.write_row(4,0,headers,fmt['header'])
    for i,(dt,raw) in enumerate(zip(r.dates,r.raw),6):
        write(ds,f'A{i}',dt)
        for j,v in enumerate(raw,1):
            ds.write_number(i-1,j,float(v),fmt['num']) if math.isfinite(v) else ds.write_blank(i-1,j,None)
        month=date(dt.year,dt.month,1)
        # Indexación por año/mes evita desfases si Bloomberg usa un día hábil.
        rf=next((row[0] for m,row in r.dataset.histories['rf'].items() if (m.year,m.month)==(month.year,month.month)),None)
        fx=next(row[0] for m,row in r.dataset.histories['fx'].items() if (m.year,m.month)==(month.year,month.month))
        write(ds,f'O{i}',rf);write(ds,f'P{i}',fx)
    ds.autofilter(4,0,last-1,15)
    ns=sheets['06_Series_EUR'];write(ns,'A1','Series mensuales en EUR',fmt['title'])
    ns.write_row(4,0,['Fecha','EQ_US','EQ_EU','EQ_JP','EQ_APXJ','EQ_EM','FI_SOV_IG','FI_COV','FI_IG','FI_HY','CASH_EUR','ALT_GOLD','EURUSD','EURJPY','Tasa caja %'],fmt['header'])
    start_rf=next(i for i in range(n) if math.isfinite(r.raw[i,6]))+6
    for i in range(n):
        row=i+6;fms=[f"='Datos'!A{row}"]
        for col in 'BCDEF':
            fxcol=None if col=='C' else 'N' if col=='D' else 'M'
            fms.append(f"='Datos'!{col}{row}/'Datos'!${col}$6*100" if not fxcol else f"=('Datos'!{col}{row}/'Datos'!{fxcol}{row})/('Datos'!${col}$6/'Datos'!${fxcol}$6)*100")
        for target,source in zip('GHIJ','HIJK'):
            fms.append('=""' if row<start_rf else f"='Datos'!{source}{row}/'Datos'!${source}${start_rf}*100")
        fms.append('=""' if row<start_rf else '=100' if row==start_rf else f"=K{row-1}*(1+O{row-1}/100*(A{row}-A{row-1})/'00_Control'!$B$14)")
        fms.extend([f"=('Datos'!G{row}/'Datos'!M{row})/('Datos'!$G$6/'Datos'!$M$6)*100",f"='Datos'!M{row}",f"='Datos'!N{row}",'=""' if row<start_rf else f"='Datos'!L{row}"])
        for j,fm in enumerate(fms):formula(ns,f'{chr(65+j)}{row}',fm,float(r.levels[i,j]),fmt['date'] if j==0 else fmt['num'])
    ms=sheets['10_Series_Motor'];write(ms,'A1','Cálculo histórico de las señales',fmt['title'])
    for addr,v in template['10_Series_Motor'].items():
        if int(re.search(r'\d+',addr)[0])<=5:write(ms,addr,v,fmt['header'] if addr.endswith('5') else fmt['body'])
    for i in range(n):
        row=i+6
        for j in range(18):
            addr=f'{chr(65+j)}{row}';src=f'{chr(65+j)}324'
            v=float(r.motor[i,j])
            # Ventanas incompletas no deben incluir encabezados ni dar falsos ceros.
            fm=Translator(template['10_Series_Motor'][src],origin=src).translate_formula(addr) if math.isfinite(v) else '=""'
            formula(ms,addr,fm,v,fmt['date'] if j==0 else fmt['num'])
    sn=sheets['02_BBG_Snapshot'];write(sn,'A1','Snapshot recibido de Bloomberg',fmt['title'])
    sn.merge_range('A2:S3',f'Fuente: {r.dataset.filename}. OAS de entrada: {p.oas_unit}. Snapshot con lag máximo ±{p.snapshot_lag} días.',fmt['note'])
    sn.write_row(3,0,['ID','Bloque','Activo','Ticker operativo','Ticker analítica','Tipo serie','PX_LAST','CRNCY','LAST_UPDATE_DT','VOLATILITY_90D','PE_RATIO','BEST_PE_RATIO','PX_TO_BOOK_RATIO','Dividend yield (%)','Yield (%)','OAS original','OAD','Estado','Nota'],fmt['header'])
    sn.set_row(3,34)
    for i,row in enumerate(r.dataset.snapshots,5):
        for j,v in enumerate(row):write(sn,f'{xlsxwriter.utility.xl_col_to_name(j)}{i}',v)
    cs=sheets['00_Control'];write(cs,'A1','Configuración de ejecución',fmt['title'])
    for addr,v in {'A5':'Versión','B5':'0.4','A6':'Fecha de corte','B6':r.cutoff,'A7':'Mes de asignación','B7':r.month,
                   'A10':'Moneda','B10':'EUR','A12':'Lag snapshot (días)','B12':p.snapshot_lag,
                   'A14':'Base días caja','B14':p.day_basis,'A16':'OAS entrada','B16':p.oas_unit}.items():write(cs,addr,v)
    for i,text in enumerate(r.warnings,20):cs.merge_range(i-1,0,i-1,7,text,fmt['note']);cs.set_row(i-1,32)
    qs=sheets['Calidad'];write(qs,'A1','Conciliaciones del motor',fmt['title'])
    qs.write_row(4,0,['Control','Actual','Esperado','Diferencia','Estado'],fmt['header'])
    for i,c in enumerate(r.checks,6):
        write(qs,f'A{i}',c['control']);write(qs,f'B{i}',c['actual']);write(qs,f'C{i}',c['expected'])
        formula(qs,f'D{i}',f'=B{i}-C{i}',c['difference']);formula(qs,f'E{i}',f'=IF(ABS(D{i})<0.000000001,"OK","ERROR")',c['status'])
    ps.write('A40','Exportación de configuración',fmt['header'])
    ps.merge_range('A41:H42','La configuración completa se descarga en JSON desde la aplicación. Las fórmulas de este libro reproducen v0.4; las constantes de riesgo regional y RF permanecen como en el modelo original.',fmt['note'])
    book.close()
    return buffer.getvalue()


def export_parameters(result):
    return json.dumps({'version':'0.4','parameters':asdict(result.params)},ensure_ascii=False,indent=2).encode('utf8')
