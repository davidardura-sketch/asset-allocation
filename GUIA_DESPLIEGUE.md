# Publicar la aplicación

## En GitHub

1. Descomprime `Asset_Allocation_Streamlit.zip`.
2. Entra en tu repositorio y elige **Add file → Upload files**.
3. Sube todos los archivos y carpetas del proyecto. En la raíz deben verse `app.py`, `requirements.txt` y la carpeta `aa`.
4. Confirma la subida con **Commit changes**.

No subas los Excel de Bloomberg ni el PDF de ejemplo: no son necesarios para que funcione.

## En Streamlit

1. Abre tu espacio de Streamlit Community Cloud, donde tienes las otras aplicaciones.
2. Pulsa **Create app** y selecciona el repositorio y su rama.
3. En **Main file path** indica `app.py`.
4. En **Advanced settings** elige **Python 3.12**.
5. Pulsa **Deploy**. No hay que introducir secretos ni claves de Bloomberg.

Si subes el proyecto dentro de una carpeta de un repositorio existente, selecciona la ruta completa, por ejemplo `asset-allocation/app.py`. Revisa las instrucciones de dependencias del README para ese caso.

## Cada mes

1. Guarda el Excel refrescado en Bloomberg.
2. Cárgalo en la aplicación.
3. Revisa la fecha de corte y pulsa **Calcular asignación**.
4. Revisa las advertencias y descarga **Excel** y **PDF de comité**.

El PDF incluye la información de señales y asignación. La configuración por defecto conserva los pesos y límites de v0.4.

## Si aparece un error

- **Falta un valor numérico guardado:** abre el archivo en Excel con Bloomberg, espera al refresco y vuelve a guardarlo.
- **Snapshot fuera del lag:** actualiza el snapshot para el cierre elegido. Un archivo antiguo no se actualiza con solo cambiar la fecha de control.
- **Falta un mes o hay duplicados:** corrige el histórico. No se interpolan datos.
- **No encuentra `aa`:** comprueba que has subido la carpeta completa, junto a `app.py`.
- **Problema al desplegar:** confirma Python 3.12 y que `requirements.txt` está junto al archivo principal. No necesita Node ni Microsoft Office.
