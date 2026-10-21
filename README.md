# Asset Allocation · Streamlit

Aplicación mensual de Asset Allocation basada en el motor mixto v0.4. Carga la plantilla actualizada en Bloomberg, calcula señales y pesos y permite descargar un Excel autónomo con fórmulas, un PDF de comité y la configuración aplicada.

## Subir a GitHub y desplegar en Streamlit

1. Descomprime el proyecto.
2. Crea un repositorio en GitHub, o utiliza uno existente para esta aplicación.
3. Sube **el contenido de esta carpeta**, conservando las subcarpetas `aa/`, `.streamlit/` y `tests/`. `app.py` y `requirements.txt` deben quedar en la raíz del repositorio. No subas el ZIP ni los Excel de Bloomberg.
4. En [Streamlit Community Cloud](https://share.streamlit.io/), selecciona **Create app** y el repositorio.
5. Selecciona la rama correspondiente, `app.py` como archivo principal y **Python 3.12** en Advanced settings. Pulsa **Deploy**.

No se necesitan secretos, claves API, Microsoft Excel ni Bloomberg en el servidor. Si lo incorporas a un repositorio con otras aplicaciones, conserva `aa/` junto a `app.py`, añade estas dependencias al `requirements.txt` de ese repositorio y selecciona la ruta de este `app.py` al desplegar. Para evitar conflictos de dependencias entre aplicaciones, se recomienda un despliegue independiente desde un repositorio o carpeta propios.

Guía oficial: [desplegar una app desde GitHub en Streamlit](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app).

## Uso mensual

1. Refresca Bloomberg en el ordenador que tiene el Terminal y el Add-in.
2. Guarda el `.xlsx` o `.xlsm` cuando termine el cálculo. Se aceptan fórmulas **con valores calculados guardados** o una copia pegada como valores.
3. Carga el archivo en la aplicación. Selecciona el cierre mensual y revisa la configuración. La propuesta se identifica con el mes siguiente al cierre.
4. Pulsa **Calcular asignación**.
5. Revisa las pestañas Asignación, Señales y Calidad de datos. Descarga el Excel y el PDF.

La fecha por defecto se infiere de los históricos, no del nombre del archivo ni de una celda de control que pueda haber quedado anticuada. Si el control original difiere, se deja constancia. Para fechas actuales se usa Europe/Madrid.

**OAS:** el valor habitual de `INDEX_OAS_TSY` en los archivos utilizados por v0.4 está expresado en porcentaje. Por ejemplo, `0,487466` se convierte en `48,7466 bp`. Si un archivo ya lleva puntos básicos, selecciona esa unidad explícitamente. La elección queda en el Excel, el PDF y la configuración.

## Estructura de entrada

Se admiten la plantilla Bloomberg v0.2 de octubre y los motores v0.4 que conservan sus pestañas de entrada:

| Pestaña | Datos utilizados |
|---|---|
| `01_Universo` | IDs A5:A15, tipo de serie I, divisa J, estado P |
| `02_BBG_Snapshot` | IDs en A; campos G:Q; fecha en I. Las filas se identifican por ID, no por su orden. |
| `03_Hist_RV` | Fecha en A y seis series en B:G: US, Europa, Japón, Asia ex Japón, EM y oro |
| `04_Hist_RF` | Fecha en A y cinco series en B:F: soberano, cédulas, IG, HY y tipo de caja |
| `05_Hist_FX` | Fecha en A, EURUSD en B y EURJPY en C |
| `00_Control` | Fecha original de corte en B6, solo para contrastar |

Los históricos pueden comenzar en las filas 10 o 13, como en los archivos anteriores. Las cabeceras de Bloomberg se distinguen de las observaciones. Se requiere una observación por mes, al menos 37 cierres consecutivos de RV y RF y las divisas correspondientes. Se alinean por año y mes; se conservan las fechas originales. El cálculo ACT/360 utiliza los días entre cierres de RV, igual que el motor anterior.

IDs: `EQ_US`, `EQ_EU`, `EQ_JP`, `EQ_APXJ`, `EQ_EM`, `FI_SOV_IG`, `FI_COV`, `FI_IG`, `FI_HY`, `CASH_EUR`, `ALT_GOLD`.

Si cambia el orden de las columnas históricas o el universo de activos, hay que adaptar el importador y el modelo. La aplicación verifica divisas, IDs y campos numéricos; no valida en el Terminal que un ticker sea el índice aprobado. El archivo de entrada sigue siendo la fuente de los datos.

## Reglas de v0.4 conservadas

- Pesos estratégicos: RV 40%, RF 40%, liquidez 15%, oro 5%. RV limitada al 30%-60%.
- Peso de señales: tendencia 50%, valoración 25%, riesgo 25%.
- Trend: retornos 3/6/12m con pesos 20/30/35%, más amplitud sobre media 10m al 15%.
- Valoración: earnings yield forward ponderado menos EURIBOR 12m, neutral del 2,5% y escala del 2,5%.
- Riesgo: volatilidad RV 6m/36m al 40%, drawdown 12m al 30% y HY/soberano 6m al 30%.
- RV regional: pesos estratégicos 29,4% US, 60% Europa, 3,2% Japón, 1,6% Asia ex Japón y 5,8% EM. Scores 45% tendencia, 40% valoración relativa y 15% riesgo. Tilt relativo máximo 25%.
- RF: pesos internos 60,5% soberano, 10,75% cédulas, 16,25% IG y 12,5% HY. Scores 40% tendencia, 25% riesgo y 35% fundamental. Fundamental 60% Yield y 40% OAS. Tilt relativo máximo 20%.
- Residual defensivo proporcional: 2/3 RF, 1/4 caja y 1/12 oro con los pesos por defecto. Se deriva de los pesos estratégicos si se modifican.
- EUR sin cobertura. Caja capitalizada con el tipo del mes anterior y días reales/360.
- Se conserva la cesta de niveles rebajados a 100 y ponderados del Excel v0.4. No se cambia por una cartera rebalanceada mensualmente.

Los parámetros avanzados se pueden recuperar desde un JSON guardado por la aplicación. Ese JSON sustituye **todos** los controles laterales. Las constantes de las reglas regionales (40/60, 60/40, 45/40/15 y escalas relativas) siguen siendo las de v0.4.

### Corrección de la ventana RF

La copia de octubre anterior arrastraba un extremo de rango sin actualizar al trasladar las fórmulas de volatilidad RF. Esta aplicación usa los **seis retornos mensuales completos hasta el corte** y actualiza los dos extremos del rango del Excel. El peso total de cada bloque de octubre se mantiene, pero cambia ligeramente el reparto interno RF y su duración: **5,478537785 años** con el archivo de octubre utilizado para verificarla. Es una corrección de la ventana, no una metodología de inversión nueva.

## Descargas

### Excel

Incluye `12_Asignacion`, `10_Parametros`, `11_Senales`, `06_Series_EUR`, `10_Series_Motor`, `02_BBG_Snapshot`, `Datos`, `00_Control` y `Calidad`. El libro es una nueva salida autónoma; no sobrescribe ni conserva todas las pestañas de la plantilla de entrada. Contiene fórmulas estándar con valores calculados guardados. Los datos, fechas, divisas, snapshot y parámetros permiten seguir la asignación hasta sus inputs.

Las fórmulas se recalculan al abrirlo en Excel. Para un nuevo mes utiliza la aplicación; el PDF es una instantánea de la ejecución y no cambia al editar el Excel. Los parámetros avanzados y sus reglas están definidos en `aa/model.py` y se exportan en JSON.

### PDF

Cinco páginas A4 con el archivo de octubre:

1. Asignación por bloques, detalle RV y RF y duración.
2. Señales globales, valoración, régimen de riesgo, contribuciones y referencias.
3. Señales regionales RV y traducción del score a pesos.
4. Señales RF, fundamentales, pesos y contribuciones de duración.
5. Trazabilidad, conciliaciones, convenciones y limitaciones.

Puede crecer si los nombres o las observaciones son más largos. Las tablas repiten sus cabeceras al pasar de página.

## Validaciones y límites

Se bloquean archivos sin pestañas o datos calculados, fechas inválidas, duplicados mensuales, huecos históricos, meses abiertos, cierres ausentes, divisas no soportadas, snapshots fuera del lag, configuraciones que no suman 100% y denominadores nulos. No se interpolan históricos ni se reutiliza el último mes para cubrir uno ausente.

El lag del snapshot admite por defecto ±3 días alrededor del cierre, como v0.4. Si hay datos posteriores se indica expresamente. Los MSCI price return, los campos RF pendientes de confirmar y los avisos VIX/crédito sin actualizar permanecen visibles. El Yield se identifica como el campo de rendimiento recibido, sin certificar que sea YTW.

Los archivos se procesan en memoria de la sesión del **servidor Streamlit**. El código no los persiste en disco, no los envía a servicios externos ni usa caché de datos compartida. Al desplegar en la nube, los archivos sí se transmiten al servidor de tu despliegue; el control de acceso corresponde a la configuración de Streamlit. No se incluye autenticación propia.

Límite: 25 MB por archivo, 150 MB descomprimidos, 10.000 filas por histórico. Las macros y fórmulas recibidas no se ejecutan. Los textos de la plantilla se tratan como datos, no como instrucciones.

## Ejecutar localmente

Requiere Python 3.12. En la carpeta del proyecto:

```bash
python -m venv .venv
# macOS/Linux:
source .venv/bin/activate
# Windows PowerShell:
# .venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Las versiones se fijan en `requirements.txt` para repetir el entorno probado. No hay que instalar Node, un compilador LaTeX ni un programa de Office.

## Pruebas

```bash
python -m unittest discover -s tests -v
```

Las pruebas utilizan mercados sintéticos y comprueban límites, conciliaciones, sensibilidad de inputs, meses ausentes, lag, unidades OAS, caja con tipo anterior y días reales, seis observaciones de volatilidad RF, lectura de plantillas y exportación. GitHub Actions las ejecuta con cada cambio. La verificación con los datos reales de octubre se hizo localmente; esos datos no forman parte del repositorio.

## Organización

```text
app.py                 Interfaz Streamlit
requirements.txt       Dependencias de despliegue
aa/model.py            Lectura y cálculo v0.4
aa/exports.py          Excel y configuración JSON
aa/report.py           Informe PDF
aa/templates/v04.json  Etiquetas y fórmulas de referencia sin datos de mercado
tests/                 Pruebas con datos sintéticos
.streamlit/config.toml Apariencia y límite de carga
.github/workflows/     Pruebas automáticas en GitHub
```
