# Automatización de Formularios PHR N°6.1 (MINVU D.S. N°10)
### Diagnóstico Técnico - Social de las Familias | Comuna de Perquenco

Aplicación local desarrollada en Python para el procesamiento, auditoría y llenado masivo de fichas de diagnóstico habitacional rural a partir de la base de datos oficial en Excel.

---

## 🚀 Cómo Iniciar la Aplicación

Para iniciar la aplicación, simplemente haz **doble clic** en el archivo:
```text
Iniciar_App.bat
```
La aplicación se abrirá automáticamente en tu navegador web predeterminado en `http://localhost:8501`.

---

## 📌 Principio Rector: Fidelidad Estricta a la Fuente

En cumplimiento con las directrices de rigurosidad técnica y administrativa:
* **Cero Alucinación / Cero Inferencia Arbitraria**: Si un campo no existe en el Excel o no puede deducirse con absoluta certeza factual, **permanece estrictamente en blanco**.
* **Ítem 1.2 (Actividades Económicas - Tabla 5)**: 
  - Se vincula automáticamente la hoja `Actividad Económica` mediante cruce unificado por RUT y nombre normalizado.
  - Clasificación exacta y estandarizada en los 7 sectores MINVU: `Agricultura`, `Forestal`, `Pesca`, `Minería`, `Turismo Rural`, `Servicios` y `Otras (Especificar)`.
  - En las actividades clasificadas como **"Otras"**, se inserta automáticamente la especificación factual breve (ej. *Papelería creativa*, *Flores eternas*, *Artesanía en madera*, *Pastelería*, etc.) **justo abajo del rótulo `Otras (Especificar)`** en la celda oficial.
  - Para los postulantes sin actividad económica declarada en la base (89 casos), la Tabla 5 permanece **100% limpia y en blanco** (sin marcas inventadas).
* **Campos sin información en base**: Teléfono, Rol SII y Dirección de Terreno se mantienen en blanco (con opción configurable para replicar la dirección de residencia RSH si el usuario así lo define).

---

## 📊 Cómo Actualizar o Cambiar la Planilla Excel

Puedes cambiar o actualizar la planilla de postulantes de tres formas muy sencillas:

1. **Subir directamente desde la App:** En la barra lateral izquierda, utiliza la opción **"📤 Subir o actualizar planilla Excel (.xlsx)"** para arrastrar o cargar tu archivo nuevo.
2. **Copiar el archivo nuevo a la carpeta:** Si tienes un archivo nuevo (ej. `BASE PERQUENCO 07.10.2026.xlsx`), pégalo directamente en esta carpeta. La app lo detectará automáticamente y lo seleccionará en el desplegable de planillas (ordenado por fecha más reciente).
3. **Reemplazar el archivo existente:** Si guardas tus cambios sobre el mismo archivo Excel, la app detectará la actualización y recargará los datos de inmediato.
4. **Mapeo dinámico inteligente:** El lector busca las columnas por nombre (`NOMBRE`, `RUT`, `DIRECCION RSH`, `CONYUGE`, etc.), por lo que funciona correctamente incluso si agregas columnas nuevas o cambias su orden.

---

## 👁️ Rol de Gemini (API Key Cuota Gratuita)

Gemini actúa estrictamente como **los "ojos" del sistema**:
1. **Desglose de Direcciones**: Separa con precisión direcciones mixtas en calle, número y lote/sitio sin deformar nombres ni inventar numeración.
2. **Desambiguación de Género e Hijos**: Clasifica con exactitud el género biológico de cada hijo y pariente a partir de sus nombres propios chilenos, transformando términos genéricos como 'HIJO/A' o 'HIJA/O DE AMBOS' en 'Hijo' (Masculino) o 'Hija' (Femenino) en la Tabla 4 con sus respectivas edades.
3. **Detección Factual de Discapacidad**: Identifica y redacta fielmente la condición declarada en las notas del Excel.
4. **Redacción de Observaciones Factuales**: Sintetiza observaciones breves para la Tabla 4 basadas 100% en la edad y rol familiar real.
5. **Modo Offline**: Si no se ingresa clave de API, el sistema funciona de manera autónoma e instantánea mediante su motor determinista local.

---

## 📑 Pestañas de la Aplicación

1. **👁️ 1. Auditoría Individual:**
   - Permite seleccionar a cualquier postulante y ver en tiempo real cómo quedarán completadas las Tablas 1 a 5 antes de emitir los documentos.
   - Generación de prueba individual en DOCX descargable.
2. **📈 2. Métricas y Análisis del Padrón (Nuevo Dashboard):**
   - **Indicadores Clave (KPIs):** Postulantes (155), población beneficiaria total (393 habitantes), % Jefatura Femenina (83.2%), % Pueblo Mapuche (29.7%), % Familias con Discapacidad (11.6%), menores de 18 años (174 dependientes).
   - **Gráficos Demográficos Interactivos:** Jefatura de hogar por género (donut), pirámide etaria (barras), tamaño de grupos familiares y estado civil.
   - **Caracterización Económica y Laboral (Tabla 5):** Gráficos de cobertura laboral y sectores MINVU (Servicios, Agricultura, Otras, Forestal, Turismo Rural, Minería).
   - **Desglose de Microemprendimientos:** Tabla detallada de las 8 familias con oficios manuales/artesanales y su glosa normada.
   - **Tipologías Habitacionales y Territorialidad:** Gráficos de tipologías propuestas y factor de aislamiento (RE 3130).
3. **⚡ 3. Generación Masiva (155 Fichas):**
   - Procesamiento en lote a alta velocidad con barra de progreso, creación automática de los 155 archivos Word y descarga consolidada en formato ZIP.
4. **📊 4. Explorador de Base de Datos:**
   - Vista tabular interactiva para buscar, ordenar y filtrar los datos completos de los postulantes.

---

## 🛠️ Estructura del Proyecto

* `app.py`: Interfaz gráfica interactiva (Streamlit).
* `Iniciar_App.bat`: Lanzador directo para Windows.
* `src/excel_reader.py`: Lector y normalizador de la base Excel (`BASE PERQUENCO 02.10.2026.xlsx`).
* `src/gemini_auditor.py`: Motor de auditoría y consolidación (Gemini y motor local).
* `src/docx_generator.py`: Generador de documentos Word respetando bordes, columnas y tipografía oficial (`Gadugi` 9pt).
* `formularios_generados/`: Carpeta donde se guardan los archivos generados y el paquete ZIP de descarga.
