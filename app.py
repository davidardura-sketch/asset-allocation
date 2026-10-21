"""Entrada para Streamlit Community Cloud: streamlit run app.py."""
from dataclasses import asdict, fields, replace
from datetime import date, timedelta
from hashlib import sha256
from html import escape
import json
from zoneinfo import ZoneInfo
from datetime import datetime

import pandas as pd
import streamlit as st

from aa.model import Parameters, InputError, load_excel, calculate, next_month
from aa.exports import export_excel, export_parameters
from aa.report import export_pdf, num, pct

st.set_page_config(page_title='Asset Allocation',page_icon='📊',layout='wide')
st.markdown('''<style>
 .block-container {max-width:1440px;padding-top:2.5rem;padding-bottom:3rem}
 h1 {font-weight:700;letter-spacing:-1.4px;font-size:2.8rem!important}
 h2 {letter-spacing:-.5px} h3 {font-size:1.2rem!important}
 [data-testid="stMetric"] {background:white;border:1px solid #DCE4EA;border-radius:12px;padding:20px}
 [data-testid="stMetricLabel"] {color:#526579}
 [data-testid="stMetricValue"] {font-size:2rem;letter-spacing:-.7px}
 [data-testid="stFileUploader"] {background:white;border:1px dashed #9BB3BC;border-radius:12px;padding:14px}
 .eyebrow {color:#247078;letter-spacing:2px;font-size:11px;font-weight:700;margin-bottom:10px}
 .step {font-size:12px;letter-spacing:1px;color:#526579;font-weight:600;margin:12px 0}
 .empty {background:#FFFFFF;border:1px solid #DCE4EA;border-radius:12px;padding:28px;height:100%;min-height:150px}
 .empty strong {color:#17365D;font-size:17px} .empty p {font-size:14px;color:#526579;margin-top:10px}
 .badge {display:inline-block;background:#E3F0EE;color:#205D61;border-radius:25px;padding:5px 12px;font-size:12px;font-weight:600}
 </style>''',unsafe_allow_html=True)


def dataframe(records,labels,formats=None):
    frame=pd.DataFrame(records)[list(labels)].rename(columns=labels)
    if formats:
        st.dataframe(frame.style.format({labels[k]:v for k,v in formats.items()}),hide_index=True,width='stretch')
    else:
        st.dataframe(frame,hide_index=True,width='stretch')


def reset_session():
    for key in ('source_data','result','exports','result_key','upload_digest'):
        st.session_state.pop(key,None)


st.markdown('<div class="eyebrow">INVESTMENT DESK / MOTOR MIXTO V0.4</div>',unsafe_allow_html=True)
head,meta=st.columns([4,1])
with head:
    st.title('Asset Allocation')
    st.write('Del Excel de Bloomberg a la propuesta mensual de cartera.')
with meta:
    st.markdown('<br><span class="badge">EUR · Rebalanceo mensual</span>',unsafe_allow_html=True)

with st.sidebar:
    st.markdown('### Configuración del motor')
    st.caption('Mismos pesos y reglas que la versión v0.4 de los meses anteriores.')
    with st.expander('Pesos estratégicos y límites',expanded=True):
        eq=st.number_input('Renta variable estratégica (%)',0.,99.,40.,step=1.)
        fi=st.number_input('Renta fija estratégica (%)',0.,100.,40.,step=1.)
        cash=st.number_input('Liquidez estratégica (%)',0.,100.,15.,step=1.)
        gold=st.number_input('Oro estratégico (%)',0.,100.,5.,step=1.)
        mn,mx=st.columns(2)
        emin=mn.number_input('RV mínima (%)',0.,100.,30.,step=1.)
        emax=mx.number_input('RV máxima (%)',0.,99.,60.,step=1.)
        st.caption(f'Total estratégico: {eq+fi+cash+gold:.1f}%')
    with st.expander('Señales y datos de entrada'):
        tw=st.number_input('Peso tendencia (%)',0.,100.,50.,step=5.)
        vw=st.number_input('Peso valoración (%)',0.,100.,25.,step=5.)
        rw=st.number_input('Peso riesgo (%)',0.,100.,25.,step=5.)
        lag=st.number_input('Lag snapshot permitido (días)',0,31,3)
        oas=st.selectbox('Unidad OAS en el archivo',options=['Porcentaje (v0.4)','Puntos básicos'])
        st.caption('La plantilla habitual usa INDEX_OAS_TSY en porcentaje. 0,487466% equivale a 48,7466 bp.')
    with st.expander('Configuración avanzada en JSON'):
        advanced=st.file_uploader('Recuperar parámetros guardados',type=['json'],key='params_upload')
        st.caption('Opcional. Un JSON guardado sustituye todos los parámetros de los controles anteriores.')
    st.divider()
    st.caption('El Excel se envía al servidor Streamlit para procesarlo en memoria. Esta aplicación no lo guarda en disco ni lo añade a GitHub. No utiliza cachés compartidas entre usuarios.')
    if st.button('Vaciar sesión',width='stretch'):
        reset_session();st.session_state['uploader_generation']=st.session_state.get('uploader_generation',0)+1
        st.rerun()

st.markdown('<div class="step">01 / CARGAR DATOS</div>',unsafe_allow_html=True)
upload=st.file_uploader('Excel actualizado en Bloomberg',type=['xlsx','xlsm'],
                         help='Usa la plantilla con Universo, Snapshot e históricos RV, RF y FX. Guarda los datos calculados antes de cargarla.',
                         key=f'bloomberg_{st.session_state.get("uploader_generation",0)}')
st.caption('Formatos .xlsx y .xlsm · Hasta 25 MB · No requiere conexión al Terminal · Las macros no se ejecutan')

try:
    p=Parameters(equity=eq/100,bonds=fi/100,cash=cash/100,gold=gold/100,equity_min=emin/100,
                 equity_max=emax/100,signal_weights=(tw/100,vw/100,rw/100),snapshot_lag=lag,
                 oas_unit='percent' if oas.startswith('Porcentaje') else 'bp')
    if advanced:
        if advanced.size>100_000:raise InputError('El JSON supera 100 KB.')
        config=json.loads(advanced.getvalue())
        if config.get('version')!='0.4':raise InputError('El JSON debe declarar version 0.4.')
        allowed={f.name for f in fields(Parameters)}
        if not isinstance(config.get('parameters'),dict) or set(config['parameters'])!=allowed:
            raise InputError('El JSON debe contener todos los parámetros v0.4 y ningún campo desconocido.')
        p=Parameters(**config['parameters'])
        st.info('Configuración JSON activa. Se aplican sus parámetros completos en lugar de los controles laterales.')
    p.validate()
except (InputError,json.JSONDecodeError,TypeError,AttributeError) as exc:
    st.error(f'Configuración no válida: {exc}');st.stop()

if upload is None:
    if 'upload_digest' in st.session_state:reset_session()
    st.markdown('<div class="step">02 / REVISAR Y DESCARGAR</div>',unsafe_allow_html=True)
    for col,title,body in zip(st.columns(3),['Señales del mercado','Propuesta de cartera','Excel y PDF'],
                              ['Tendencia, valoración y régimen de riesgo, con detalle regional y de renta fija.',
                               'Pesos por bloque, posiciones regionales y duración de la renta fija.',
                               'Libro con fórmulas y un informe que recoge señales, asignación y controles.']):
        col.markdown(f'<div class="empty"><strong>{title}</strong><p>{body}</p></div>',unsafe_allow_html=True)
    with st.expander('Cómo preparar el archivo'):
        st.write('1. Refresca la plantilla en el equipo que tenga Bloomberg conectado.\n2. Guarda el archivo cuando terminen las fórmulas BDP/BDH.\n3. Cárgalo aquí. La aplicación identifica los datos por pestaña, mes e ID.\n4. Revisa la fecha de corte y pulsa Calcular asignación.')
        st.caption('No es necesario pegar todo como valores si Excel ha guardado los valores calculados. Si falta una caché, la aplicación pedirá refrescar y guardar de nuevo.')
    st.stop()

content=upload.getvalue();digest=sha256(content).hexdigest()
if digest!=st.session_state.get('upload_digest'):
    reset_session()
    try:
        with st.spinner('Leyendo datos de Bloomberg…'):
            st.session_state['source_data']=load_excel(content,upload.name)
            st.session_state['upload_digest']=digest
    except InputError as exc:
        st.error(str(exc));st.stop()
    except Exception:
        st.error('No se ha podido leer este Excel. Comprueba que no esté protegido y que use la estructura de la plantilla Bloomberg.');st.stop()
data=st.session_state['source_data']
today=datetime.now(ZoneInfo('Europe/Madrid')).date()
closed=today.replace(day=1)-timedelta(days=1)
months=[m for m in data.histories['rv'] if m<=closed]
if not months:
    st.error('No hay cierres mensuales completos disponibles.');st.stop()
control_col,description=st.columns([1,3])
with control_col:
    cutoff=st.selectbox('Fecha de corte',options=list(reversed(months)),format_func=lambda d:d.strftime('%d/%m/%Y'),key=f'cut_{digest[:12]}')
with description:
    st.markdown(f'**Asignación para {next_month(cutoff):%m/%Y}**')
    st.caption(f'Últimos históricos: RV {max(data.histories["rv"]):%d/%m/%Y} · RF {max(data.histories["rf"]):%d/%m/%Y} · FX {max(data.histories["fx"]):%d/%m/%Y}')
key=sha256((digest+json.dumps(asdict(p),sort_keys=True)+cutoff.isoformat()).encode()).hexdigest()
run=st.button('Calcular asignación',type='primary',width='content')
if run:
    for k in ('result','exports','result_key'):st.session_state.pop(k,None)
    try:
        with st.spinner('Calculando señales, asignación e informes…'):
            result=calculate(data,p,cutoff,today)
            exports=(export_excel(result),export_pdf(result),export_parameters(result))
            st.session_state.update(result=result,exports=exports,result_key=key)
    except InputError as exc:
        st.error(str(exc));st.stop()
    except Exception:
        st.error('No se han podido generar los informes. Revisa el archivo y prueba de nuevo. No se muestran resultados de la ejecución anterior.');st.stop()
if st.session_state.get('result_key')!=key:
    if 'result' in st.session_state:st.info('Los datos o parámetros han cambiado. Pulsa Calcular asignación para actualizar los resultados y las descargas.')
    st.stop()

r=st.session_state['result'];xlsx,pdf,config=st.session_state['exports']
st.divider()
st.markdown('<div class="step">02 / RESULTADO MENSUAL</div>',unsafe_allow_html=True)
st.subheader(f'Propuesta de cartera · {r.month:%m/%Y}')
st.caption(f'Cierre {r.cutoff:%d/%m/%Y} · {r.position} · Score mixto {num(r.score,3)} · Duración RF {num(r.signals["Q48"])} años')
for c,b in zip(st.columns(4),[r.blocks[2],r.blocks[1],r.blocks[0],r.blocks[3]]):
    c.metric(b['name'],pct(b['recommended']),num(b['change']*100)+' pp vs. estratégico',delta_color='off')
st.caption('Los cambios se comparan con los pesos estratégicos aplicados, no con el mes anterior.')

t_alloc,t_signals,t_quality,t_method=st.tabs(['Asignación','Señales','Calidad de datos','Metodología'])
with t_alloc:
    left,right=st.columns([1.25,1])
    with left:
        st.markdown('### Distribución por bloques')
        dataframe(r.blocks,{'name':'Bloque','strategic':'Estratégico','recommended':'Recomendado','change':'Cambio (pp)'},
                  {'strategic':pct,'recommended':pct,'change':lambda x:num(x*100)})
    with right:
        chart=pd.DataFrame({'Bloque':[b['name'] for b in r.blocks],'Peso (%)':[b['recommended']*100 for b in r.blocks]}).set_index('Bloque')
        st.bar_chart(chart,color='#246B70',height=230)
    st.markdown('### Renta variable por región')
    dataframe(r.regions,{'name':'Región','within':'Peso dentro RV','portfolio':'Peso cartera','score':'Score'},
              {'within':pct,'portfolio':pct,'score':lambda x:num(x,3)})
    st.markdown('### Renta fija y duración')
    dataframe(r.fixed_income,{'name':'Segmento','within':'Peso RF','portfolio':'Peso cartera','score':'Score',
                             'yield_pct':'Yield (%)','oas_bp':'OAS (bp)','duration':'OAD (años)','duration_portfolio':'Contrib. duración cartera'},
              {'within':pct,'portfolio':pct,'score':lambda x:num(x,3),'yield_pct':num,'oas_bp':num,'duration':num,'duration_portfolio':num})
    st.caption('La duración se obtiene de la composición de segmentos RF. No incluye una señal direccional de tipos.')
with t_signals:
    s=r.signals
    st.markdown('### Contribución a la señal mixta')
    signals=[{'name':name,'score':s[f'D{i}'],'weight':s[f'E{i}'],'contribution':s[f'F{i}']} for name,i in [('Tendencia',27),('Valoración',28),('Riesgo',29)]]
    dataframe(signals,{'name':'Señal','score':'Score (-1 a +1)','weight':'Peso','contribution':'Contribución'},
              {'score':lambda x:num(x,3),'weight':pct,'contribution':lambda x:num(x,3)})
    with st.expander('Métricas globales',expanded=True):
        records=[]
        for name,i in [('Retorno RV 3m',6),('Retorno RV 6m',7),('Retorno RV 12m',8),('Amplitud sobre media 10m',9),('Volatilidad RV 6m / 36m',20),('Drawdown RV 12m',21),('HY / soberano 6m',22)]:
            records.append({'Métrica':name,'Valor':num(s[f'B{i}'],3) if i==20 else pct(s[f'B{i}']),'Score':num(s[f'D{i}'],3),'Contribución':num(s[f'F{i}'],3)})
        st.dataframe(pd.DataFrame(records),hide_index=True,width='stretch')
        c1,c2,c3=st.columns(3)
        c1.metric('Earnings yield forward',pct(s['B14']));c2.metric('EURIBOR 12m',pct(s['B15']));c3.metric('ERP proxy',pct(s['B16']))
    st.markdown('### Señales regionales RV')
    dataframe(r.regions,{'name':'Región','return6':'Ret. 6m','return12':'Ret. 12m','trend':'Trend','pe':'PER fwd','pb':'P/B',
                         'value':'Valor','volatility':'Vol. 90d','risk':'Riesgo','score':'Mixto','tilt':'Factor tilt'},
              {k:(pct if k in ('return6','return12','volatility') else lambda x:num(x,3)) for k in ('return6','return12','trend','pe','pb','value','volatility','risk','score','tilt')})
    st.markdown('### Señales de renta fija')
    dataframe(r.fixed_income,{'name':'Segmento','return6':'Ret. 6m','return12':'Ret. 12m','volatility':'Vol. 6m','trend':'Trend',
                             'risk':'Riesgo','fundamental':'Fundamental','score':'Mixto','tilt':'Factor tilt'},
              {k:(pct if k in ('return6','return12','volatility') else lambda x:num(x,3)) for k in ('return6','return12','volatility','trend','risk','fundamental','score','tilt')})
with t_quality:
    st.markdown('### Conciliaciones')
    dataframe(r.checks,{'control':'Control','actual':'Actual','expected':'Esperado','difference':'Diferencia','status':'Estado'},
              {'actual':lambda x:num(x,6),'expected':lambda x:num(x,6),'difference':lambda x:num(x,8)})
    st.markdown('### Observaciones que requieren revisión')
    for warning in r.warnings:st.warning(warning)
    st.caption('Huella del archivo: '+r.dataset.digest)
with t_method:
    st.markdown('### Metodología conservada de v0.4')
    st.write('La tendencia combina retornos a 3, 6 y 12 meses con la amplitud sobre la media de 10 meses. La valoración usa el earnings yield forward ponderado menos EURIBOR 12m. El régimen de riesgo combina volatilidad 6m/36m, drawdown 12m y comportamiento relativo HY/soberano.')
    st.write('La señal se traduce a RV de forma asimétrica alrededor del peso estratégico y dentro de los límites. El residual se reparte entre RF, liquidez y oro. Los scores regionales y RF redistribuyen los pesos dentro de sus bloques.')
    st.write('Las series se convierten a EUR sin cobertura y la caja se capitaliza ACT/360 con el tipo anterior. Se mantiene la cesta de niveles rebajados a 100 del Excel original; no se sustituye por un índice con rebalanceo mensual.')
    st.write('El Excel descargado incorpora fórmulas estándar y resultados calculados guardados. Las fórmulas se recalculan al abrirlo en Excel. Para el siguiente mes, carga la nueva plantilla en la aplicación.')
    st.json(asdict(r.params),expanded=False)

st.divider()
st.markdown('<div class="step">03 / DESCARGAR INFORMES</div>',unsafe_allow_html=True)
e,pd_col,j=st.columns([1,1,1])
suffix=r.month.strftime('%Y-%m')
e.download_button('Descargar Excel',xlsx,f'Asset_Allocation_{suffix}.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary',width='stretch',on_click='ignore')
pd_col.download_button('Descargar PDF de comité',pdf,f'Asset_Allocation_{suffix}.pdf','application/pdf',width='stretch',on_click='ignore')
j.download_button('Guardar configuración',config,f'Parametros_AA_{suffix}.json','application/json',width='stretch',on_click='ignore')
st.caption('El PDF incluye asignación, señales globales, detalle regional RV, señales RF y controles de datos. Las descargas corresponden a los parámetros y al corte mostrados.')
