# Analítica de Facturación en Salud

**Facturación, radicación, glosas, recaudo y cartera.**

Aplicación en Python y Streamlit para cargar archivos de facturación, validar su estructura y explorar indicadores financieros y operativos mediante gráficos interactivos, tablas y descargas.

La aplicación inicia sin datos. Cada persona debe cargar su archivo al comenzar una sesión nueva. No requiere bases de datos, claves, APIs externas ni un dataset dentro del repositorio.

> Los resultados dependen de la calidad y estructura del archivo cargado. Esta herramienta facilita el análisis exploratorio y no reemplaza una auditoría financiera, contable, médica o contractual.

## Archivos del proyecto

Coloque estos archivos en la raíz del repositorio:

| Archivo | Función |
| --- | --- |
| `app.py` | Aplicación completa y navegación entre secciones. |
| `requirements.txt` | Dependencias de Python. |
| `README.md` | Documentación de instalación, uso y metodología. |

No es necesario crear subdirectorios ni subir archivos de facturación al repositorio.

## Instalación y ejecución local

Entorno utilizado para las pruebas: **Python 3.12**.

Abra una terminal en la carpeta del proyecto. Se recomienda usar un entorno virtual.

### Windows — PowerShell

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

### macOS o Linux

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m streamlit run app.py
```

Abra la dirección local que indique Streamlit en la terminal. Para detener la aplicación, presione `Ctrl+C`.

### Dependencias

- Streamlit: interfaz y estado de sesión.
- pandas y NumPy: transformación y análisis de datos.
- Plotly: gráficos interactivos.
- openpyxl: lectura y escritura de Excel.
- SciPy: pruebas y cálculos estadísticos.

Los rangos de versiones están definidos en `requirements.txt`.

## Despliegue en Streamlit Community Cloud

1. Cree un repositorio en GitHub.
2. Suba `app.py`, `requirements.txt` y este `README.md` a la raíz.
3. Entre a [Streamlit Community Cloud](https://share.streamlit.io/) y conecte su cuenta de GitHub.
4. Seleccione **Create app** y la opción para desplegar una aplicación existente.
5. Seleccione el repositorio y la rama donde subió los archivos.
6. Indique `app.py` como archivo principal.
7. En las opciones avanzadas, seleccione **Python 3.12**, el entorno utilizado para las pruebas.
8. Inicie el despliegue y espere a que se instalen las dependencias.
9. Abra la URL generada por Streamlit.

No se requieren secretos ni credenciales de servicios externos. No incluya datos de facturación en GitHub.

Documentación oficial: [despliegue](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy) y [dependencias](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies).

## Archivos de datos admitidos

| Formato | Comportamiento |
| --- | --- |
| `.xlsx` | Detecta las hojas y permite elegir una. Prefiere `Facturacion` cuando existe. |
| `.csv` | Intenta detectar codificación y separador. Permite escoger coma, punto y coma o tabulación. |

Los archivos `.xls` no están habilitados. Conviértalos a `.xlsx` o `.csv` antes de cargarlos.

Use una fila de encabezados, sin títulos o subtotales intercalados. Conserve identificadores y códigos como texto cuando necesite preservar ceros iniciales.

### Columnas mínimas

| Campo | Descripción |
| --- | --- |
| `ID_Factura` | Identificador de factura. |
| `Valor_Neto` | Importe neto del registro. |
| `Fecha_Radicacion` **o** `Periodo_Cierre` | Fecha de radicación o periodo de cierre en formato `AAAA-MM`. |

La ausencia de alguna columna mínima bloquea los análisis principales y muestra instrucciones para corregir el mapeo. Para los análisis temporales también se necesitan valores de fecha o periodo válidos.

### Columnas recomendadas

`Entidad_Pagadora`, `Estado_Factura`, `Valor_Pagado`, `Saldo_Cartera` y `Glosa_Inicial`.

Si faltan, la aplicación permite continuar y limita los análisis que dependen de ellas.

### Otras columnas reconocidas

| Grupo | Campos |
| --- | --- |
| Identificación | `Numero_Radicado`, `Codigo_Entidad`, `Codigo_Servicio` |
| Fechas | `Fecha_Servicio`, `Fecha_Emision`, `Fecha_Aceptacion`, `Fecha_Pago` |
| Operación | `Regimen`, `Sede`, `Servicio`, `Linea_Servicio`, `Cantidad`, `Canal_Radicacion`, `Dias_Emision_Radicacion`, `Soporte_Completo` |
| Importes | `Valor_Bruto`, `Descuento`, `Copago_Cuota` |

Los encabezados pueden contener espacios, guiones, tildes o diferencias entre mayúsculas y minúsculas. La aplicación propone equivalencias; el usuario debe revisarlas y pulsar **Aplicar correspondencia**. Una misma columna de origen no puede asignarse a dos campos.

### Ejemplo mínimo de CSV

Ejemplo ficticio, sin información personal:

```csv
ID_Factura;Periodo_Cierre;Valor_Neto
FAC-001;2026-01;1250000
FAC-002;2026-01;875000
FAC-003;2026-02;1500000
```

Guárdelo con extensión `.csv`. También puede usar esos encabezados en una hoja de Excel.

## Flujo de uso

1. Cargue un archivo desde la barra lateral.
2. Seleccione la hoja, o revise el separador y la codificación del CSV.
3. Abra **Configurar correspondencia de columnas**, revise las propuestas y aplíquelas.
4. Revise **Limpieza controlada y formatos**. Las conversiones están activas inicialmente; la eliminación de filas y duplicados requiere seleccionarla.
5. Consulte **Calidad de datos** para revisar conversiones fallidas, valores nulos e inconsistencias.
6. Aplique los filtros globales y seleccione una sección del menú.
7. En **Descargas**, elija el contenido y formato, pulse **Preparar descarga** y descargue el resultado.

**Restaurar datos originales** desactiva las transformaciones y conserva el mapeo. Algunos análisis requerirán volver a activar las conversiones.

**Reiniciar aplicación** limpia el estado de la sesión y vuelve a solicitar el archivo.

### Formatos monetarios y fechas

- Se reconocen importes numéricos y textos con símbolos de moneda, separadores y espacios.
- En modo automático, textos ambiguos como `1.234` quedan nulos. Seleccione **Decimal coma (1.234,56)** o **Decimal punto (1,234.56)** para resolverlos según el archivo.
- Las conversiones fallidas se informan; no se reemplazan automáticamente por cero.
- La etiqueta COP, USD o EUR cambia la presentación: **no convierte divisas**.
- Las fechas pueden convertirse con día antes del mes; revise esta opción para archivos con otro formato.
- Los periodos de cierre de texto deben seguir `AAAA-MM`.
- Si falta `Periodo_Cierre`, se deriva de `Fecha_Radicacion` cuando es posible.

## Secciones disponibles

| Sección | Contenido |
| --- | --- |
| Inicio y carga de datos | Información del archivo, mapeo y vista previa. |
| Calidad de datos | Nulos, duplicados, conversiones, reglas y puntaje orientativo. |
| Resumen ejecutivo | KPI, evolución y distribución de facturación. |
| Cierre mensual | Agregaciones mensuales, trimestrales y anuales; crecimiento y promedio móvil. |
| Entidades pagadoras | Rankings, comparación, detalle, concentración y Pareto. |
| Servicios y líneas | Facturación, glosas, recaudo y cartera por servicio o línea. |
| Sedes y regímenes | Comparación por sede o régimen. |
| Estados de facturación | Categorías detectadas, conteos y cruces. |
| Glosas | Frecuencia, severidad, distribución y registros destacados. |
| Recaudo y cartera | Pagos, saldos y antigüedad aproximada desde radicación. |
| Tiempos de radicación | Distribución, rangos y comparación de tiempos. |
| Análisis estadístico | Descriptivos, correlaciones y pruebas opcionales. |
| Valores atípicos | IQR, puntaje Z y percentiles; análisis global o por grupo. |
| Insights y alertas | Hallazgos calculados y umbrales editables. |
| Explorador de datos | Búsqueda, selección de columnas, orden y vistas de revisión. |
| Descargas | Exportaciones CSV, Excel y resumen de insights en texto. |
| Metodología | Fórmulas, supuestos y limitaciones. |

## Indicadores y reglas de interpretación

| Indicador | Cálculo |
| --- | --- |
| Facturación neta | Suma de `Valor_Neto`. |
| Facturas | Número de `ID_Factura` distintos y no nulos. |
| Ticket promedio | Suma de valor neto / facturas únicas. |
| Porcentaje de glosa | Suma de glosa inicial / suma de valor neto × 100. |
| Porcentaje de recaudo | Suma de valor pagado / suma de valor neto × 100. |
| Porcentaje de cartera | Suma de saldo de cartera / suma de valor neto × 100. |
| Frecuencia de glosa | Facturas únicas con glosa positiva / facturas únicas × 100. |
| Días de radicación | Fecha de radicación − fecha de emisión, si no existe el dato reportado. |
| Crecimiento | (Valor actual − valor anterior) / valor anterior × 100, entre periodos consecutivos. |
| Promedio móvil | Media de tres periodos consecutivos disponibles. |

**Los importes se suman por registro.** Si una factura aparece en varias líneas, los valores deben ser aditivos por línea. Repetir el total de la factura en cada línea sobreestima los indicadores. La aplicación no deduce ni corrige automáticamente esa situación.

Una factura presente en varios grupos se cuenta en cada uno; por eso los conteos de grupos o estados no siempre suman el total de facturas únicas. Los importes de filas sin ID pueden permanecer en las sumas aunque esas filas no cuenten como facturas únicas.

Los denominadores cero producen resultados no disponibles. Las tasas agregadas se calculan como cocientes de sumas, no como promedios de porcentajes individuales. Las sumas omiten valores nulos; revise su cobertura antes de interpretar resultados.

### Cierres y cartera

- Los pagos, glosas y saldos se agrupan por periodo de cierre de la factura. No representan flujos de caja por fecha de pago ni saldos históricos reconstruidos.
- Los filtros temporales usan radicación cuando existe como fecha convertida; en caso contrario, usan el periodo de cierre.
- Los meses ausentes quedan sin datos, no como facturación cero.
- Las comparaciones pueden involucrar meses incompletos o parcialmente filtrados.
- La antigüedad usa la fecha de radicación y una fecha de corte elegida; no equivale a mora contractual.
- Los umbrales de alertas y rangos de días son demostrativos y editables, no estándares normativos.

### Validaciones de calidad

Se revisan, según las columnas disponibles, duplicados, nulos, valores negativos, fechas inconsistentes, pagos/glosas/saldos superiores al neto y estas identidades, con tolerancia configurable:

```text
Valor_Neto ≈ Valor_Bruto − Descuento − Copago_Cuota
Saldo_Cartera ≈ Valor_Neto − Glosa_Inicial − Valor_Pagado
```

La segunda identidad es una hipótesis de revisión: una glosa inicial en disputa no necesariamente reduce el saldo contable. Una diferencia no demuestra por sí sola un error.

El puntaje de calidad promedia completitud, unicidad de filas, validez de conversión y consistencia exploratoria. **No reemplaza una auditoría de datos.**

### Estadística

Se incluyen correlaciones Pearson y Spearman, pruebas t de Student, Mann-Whitney U, ANOVA, Kruskal-Wallis, chi-cuadrado y Shapiro-Wilk, con controles de datos y supuestos según corresponda.

La independencia de observaciones debe validarse antes de las pruebas. Las líneas de una misma factura pueden estar relacionadas. No rechazar una hipótesis nula no demuestra que sea verdadera. La correlación no implica causalidad; algunas asociaciones resultan de identidades contables. No se aplica corrección automática por comparaciones múltiples.

La detección de atípicos requiere al menos cuatro valores por grupo. Un atípico puede ser legítimo y no se elimina automáticamente.

## Exportaciones

Se pueden descargar datos filtrados, cierre mensual, análisis por entidades/servicios/sedes, registros con glosa o cartera, atípicos, informe de calidad y resumen de insights.

- CSV: UTF-8 con BOM, punto y coma como separador y coma decimal.
- Excel: generación en memoria; los datos extensos se distribuyen en hojas según el límite de filas de Excel.
- Informe de calidad en Excel: hojas de reglas, nulos, conversiones y puntaje.
- Insights: CSV, Excel o texto.
- Los textos que podrían ejecutarse como fórmulas en una hoja de cálculo reciben un apóstrofo de protección.

## Privacidad y rendimiento

> Evite cargar nombres de pacientes, documentos de identidad, historias clínicas, diagnósticos u otra información personal o sensible. Utilice archivos anonimizados y siga las políticas de tratamiento de datos de su organización.

En un despliegue web, los archivos se cargan y procesan en la memoria del servidor de Streamlit, no exclusivamente en el navegador. El código de la aplicación no envía los datos a APIs externas ni los escribe en archivos persistentes.

La caché de datos incluye un token de sesión, tiene entradas limitadas y TTL de 15 minutos. Cerrar la pestaña o reiniciar no garantiza la eliminación inmediata de todas las entradas en caché; pueden permanecer hasta su expiración o desalojo. El código no incorpora autenticación propia.

El límite de carga depende de la configuración de Streamlit y la capacidad efectiva depende de los recursos del servidor. Excel puede requerir varias veces su tamaño en memoria. Los gráficos de detalle utilizan muestras reproducibles de hasta 5.000 registros; los indicadores usan todos los registros filtrados. Las tablas limitan la vista, mientras que las descargas conservan los registros correspondientes.

## Solución de problemas

| Situación | Acción sugerida |
| --- | --- |
| No se puede leer el CSV | Cambie el separador o la codificación; revise encabezados y estructura. |
| No se puede leer Excel | Verifique que sea `.xlsx`, no esté protegido y tenga una hoja con datos. |
| Faltan columnas mínimas | Corrija los encabezados o use el mapeo manual y aplíquelo. |
| Importes quedan nulos | Seleccione el formato decimal correcto y consulte las conversiones fallidas. |
| No quedan registros | Limpie los filtros o revise las exclusiones de limpieza. |
| No aparece un análisis | Verifique que estén disponibles y convertidas sus columnas necesarias. |
| Totales demasiado altos | Revise si el importe total de factura se repite por cada línea. |
| No se ejecuta una prueba | Revise tamaños de muestra, variabilidad, categorías y supuestos. |
| La descarga de Excel falla | Pruebe CSV o reduzca el conjunto filtrado. |
| Error al desplegar | Confirme la ubicación de los archivos, `app.py` como entrada y las dependencias; consulte los registros de Streamlit. |

## Verificación realizada

Durante el desarrollo se verificaron sintaxis, inicio sin datos, lectura CSV y Excel con varias hojas, conversiones, divisiones por cero, meses ausentes, exportaciones y navegación por las 17 secciones con datos completos y mínimos.

También se ejercitaron las rutas de pruebas estadísticas, el mapeo, casos sin columna mínima o con valores inválidos y procesamiento/agregación de 100.000 registros sintéticos. Estas verificaciones no constituyen una certificación contable ni garantizan el rendimiento con cualquier archivo o plan de alojamiento. No se incluyen datasets ni scripts de prueba en los archivos entregados.
