# Analítica de Facturación en Salud

**Facturación, radicación, glosas, recaudo y cartera**

Aplicación web desarrollada en Python y Streamlit para cargar archivos de facturación del sector salud, revisar su calidad y explorar indicadores financieros y operativos mediante tablas y gráficos interactivos.

La aplicación comienza sin datos. Cada usuario debe cargar un archivo en su sesión; no es necesario incorporar un dataset al repositorio.

> Los resultados dependen de la calidad y estructura del archivo cargado. Esta herramienta facilita el análisis exploratorio y no reemplaza una auditoría financiera, contable, médica o contractual.

## Contenido

- [Características](#características)
- [Archivos del proyecto](#archivos-del-proyecto)
- [Instalación y ejecución local](#instalación-y-ejecución-local)
- [Despliegue en Streamlit Community Cloud](#despliegue-en-streamlit-community-cloud)
- [Preparación del archivo de datos](#preparación-del-archivo-de-datos)
- [Guía de uso](#guía-de-uso)
- [Módulos de análisis](#módulos-de-análisis)
- [Indicadores y metodología](#indicadores-y-metodología)
- [Descargas](#descargas)
- [Privacidad y seguridad](#privacidad-y-seguridad)
- [Limitaciones](#limitaciones)
- [Solución de problemas](#solución-de-problemas)

## Características

- Carga de archivos Excel `.xlsx` y CSV `.csv` en memoria.
- Selección de hoja de Excel, con preferencia por `Facturacion` si existe.
- Detección de separador y codificación de CSV, con selección manual.
- Normalización de encabezados y asistente de correspondencia de columnas.
- Diagnóstico de nulos, duplicados, fechas inválidas, valores no numéricos e inconsistencias financieras.
- Limpieza configurable sobre una copia de los datos.
- Filtros globales y comparación del alcance completo frente al filtrado.
- Indicadores de facturación, glosas, recaudo, cartera y tiempos de radicación.
- Cierres mensuales, trimestrales y anuales.
- Análisis por entidades, servicios, líneas, sedes, regímenes y estados.
- Gráficos interactivos con Plotly, rankings y análisis de Pareto.
- Estadística descriptiva, correlaciones y pruebas estadísticas exploratorias.
- Detección de valores atípicos por IQR, puntaje Z y percentiles.
- Insights basados en reglas y alertas con umbrales editables.
- Exportaciones en CSV, Excel y texto, según el análisis.

No utiliza bases de datos, APIs externas, modelos de lenguaje ni claves de servicios externos.

## Archivos del proyecto

```text
app.py
requirements.txt
README.md
```

| Archivo | Función |
| --- | --- |
| `app.py` | Contiene la interfaz, lectura, validaciones, cálculos, gráficos y exportaciones. |
| `requirements.txt` | Declara las dependencias externas necesarias. |
| `README.md` | Documenta instalación, uso, metodología y limitaciones. |

Mantenga estos archivos en la raíz del repositorio. No es necesario crear subdirectorios ni subir archivos de facturación.

### Dependencias

```text
streamlit>=1.40
pandas>=2.1
numpy
plotly>=5.18
openpyxl>=3.1
scipy>=1.11
```

Las librerías estándar de Python no se incluyen en `requirements.txt`.

## Instalación y ejecución local

Como configuración recomendada, utilice Python 3.12 y un entorno virtual. Los comandos siguientes se ejecutan desde la carpeta que contiene `app.py` y `requirements.txt`.

### macOS y Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

### Windows, con PowerShell

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Si no desea activar el entorno, puede utilizar directamente su intérprete. Por ejemplo, en Windows:

```powershell
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m streamlit run app.py
```

Abra en el navegador la dirección local que muestre Streamlit en la terminal. Para detener la ejecución local, use `Ctrl+C` en esa terminal.

## Despliegue en Streamlit Community Cloud

1. Cree un repositorio en GitHub.
2. Suba `app.py`, `requirements.txt` y, opcionalmente, este `README.md` a la raíz.
3. Ingrese a Streamlit Community Cloud y conecte su cuenta de GitHub.
4. Seleccione **Create app** y la opción para desplegar una aplicación existente.
5. Seleccione el repositorio y la rama correspondiente.
6. Indique `app.py` como archivo principal.
7. Como recomendación, seleccione Python 3.12 en la configuración avanzada si está disponible.
8. Pulse **Deploy**.
9. Abra la dirección generada por la plataforma.

No se proporciona una URL de proyecto porque depende del despliegue que realice el propietario del repositorio.

**No suba el dataset a GitHub.** La aplicación solicita el archivo al usuario durante la sesión.

## Preparación del archivo de datos

### Formatos admitidos

| Formato | Disponibilidad | Consideraciones |
| --- | --- | --- |
| Excel `.xlsx` | Admitido | Permite seleccionar la hoja. |
| CSV `.csv` | Admitido | Permite detectar o seleccionar separador y codificación. |
| Excel `.xls` | No habilitado | Conviértalo a `.xlsx` antes de cargarlo. |

Use una fila de encabezados y una tabla de registros. Evite títulos adicionales, celdas combinadas y totales generales mezclados con los datos.

### Columnas mínimas

- `ID_Factura`.
- `Valor_Neto`.
- Al menos una de estas columnas: `Fecha_Radicacion` o `Periodo_Cierre`.

Si falta alguna condición mínima, la aplicación informa el problema y no ejecuta los análisis financieros principales. Puede corregir el archivo o usar el asistente de correspondencia.

### Columnas recomendadas

- `Entidad_Pagadora`.
- `Estado_Factura`.
- `Valor_Pagado`.
- `Saldo_Cartera`.
- `Glosa_Inicial`.

Su ausencia no impide continuar, pero limita los análisis que dependen de ellas.

### Estructura completa reconocida

| Grupo | Columnas |
| --- | --- |
| Identificación | `ID_Factura`, `Numero_Radicado` |
| Fechas y periodo | `Fecha_Servicio`, `Fecha_Emision`, `Fecha_Radicacion`, `Fecha_Aceptacion`, `Fecha_Pago`, `Periodo_Cierre` |
| Entidad | `Codigo_Entidad`, `Entidad_Pagadora`, `Regimen` |
| Operación | `Sede`, `Codigo_Servicio`, `Servicio`, `Linea_Servicio`, `Cantidad`, `Canal_Radicacion`, `Estado_Factura`, `Dias_Emision_Radicacion`, `Soporte_Completo` |
| Finanzas | `Valor_Bruto`, `Descuento`, `Copago_Cuota`, `Valor_Neto`, `Glosa_Inicial`, `Valor_Pagado`, `Saldo_Cartera` |

### Encabezados y correspondencias

La aplicación compara encabezados sin distinguir mayúsculas, tildes o pequeñas diferencias de espacios y guiones. Por ejemplo:

- `ID Factura` → `ID_Factura`.
- `Fecha Radicación` → `Fecha_Radicacion`.
- `Entidad Pagadora` → `Entidad_Pagadora`.
- `Línea de Servicio` → `Linea_Servicio`.

Las equivalencias se muestran antes de aplicarse. Cuando existen varias candidatas para un campo, debe elegir la columna manualmente. Cada columna fuente solo puede asignarse a un campo esperado.

### Fechas y periodos

- Seleccione **Día primero** o **Mes primero** según el formato de origen.
- Los valores no convertibles se registran como fechas inválidas.
- `Periodo_Cierre` debe tener formato `AAAA-MM`, por ejemplo `2026-01`; también se aceptan celdas Excel de tipo fecha.
- Si `Periodo_Cierre` no existe y se convierte `Fecha_Radicacion`, se deriva el periodo desde esa fecha.
- Si `Periodo_Cierre` existe pero es inválido, no se reemplaza automáticamente por la fecha de radicación.

### Valores monetarios

Se reconocen números y cadenas con símbolos, separadores y espacios. Puede seleccionar:

- **Automático**.
- **Colombiano:** `1.234,56`.
- **Internacional:** `1,234.56`.

En modo automático, un separador único seguido de tres dígitos puede ser ambiguo. La aplicación lo interpreta como miles y muestra una advertencia. Confirme el formato de la fuente antes de utilizar los totales.

Cambiar la etiqueta entre COP, USD, EUR o sin símbolo **no convierte tasas de cambio**.

### Facturas con varias líneas

Los montos se suman por registro y las facturas se cuentan por identificador único no nulo.

**Si una factura aparece en varias filas, los montos deben corresponder a cada línea.** No repita el valor total de la factura en cada fila, porque se sumaría varias veces. La aplicación no deduplica importes silenciosamente.

## Guía de uso

1. **Cargue el archivo** desde la barra lateral.
2. **Seleccione la hoja** de Excel o revise separador y codificación de CSV.
3. Abra **Configurar correspondencia de columnas**.
4. Revise las sugerencias y pulse **Aplicar correspondencia**.
5. Revise **Calidad de datos** antes de interpretar los indicadores.
6. Ajuste **Limpieza controlada**, formato monetario, orden de fechas y tolerancia financiera.
7. Aplique los **Filtros globales** disponibles.
8. Navegue por los módulos de análisis.
9. Configure rangos y alertas según su propósito analítico.
10. Descargue los resultados desde **Descargas**.

### Limpieza controlada

Las opciones permiten:

- Eliminar filas completamente vacías.
- Eliminar duplicados exactos, si se activa esta opción.
- Convertir fechas y columnas numéricas.
- Sustituir categorías vacías por `Sin informacion`.
- Excluir registros con valor neto inválido.
- Excluir fechas inválidas no vacías.

La aplicación conserva los datos originales en memoria y trabaja sobre una copia. Informa conversiones fallidas y exclusiones.

**Restaurar datos originales** desactiva las transformaciones de limpieza. La copia analítica puede conservar normalización de periodos y variables derivadas; los datos originales almacenados no se modifican. Sin conversión numérica, algunos análisis quedan desactivados.

**Reiniciar aplicación** limpia el estado de sesión y devuelve la interfaz al inicio.

### Filtros globales

Según las columnas disponibles, puede filtrar por fechas, año, mes, entidad, régimen, sede, servicio, línea, estado, canal, soporte, glosa, saldo y rango de valor neto.

Los filtros se aplican al conjunto utilizado por los módulos analíticos. El diagnóstico de calidad se calcula sobre los datos originales mapeados, no sobre el subconjunto filtrado.

**Seleccionar todos** y **Limpiar filtros** restablecen los filtros a su selección predeterminada. Si el conjunto queda vacío, la aplicación muestra una advertencia.

## Módulos de análisis

| Sección | Contenido principal |
| --- | --- |
| Inicio y carga de datos | Archivo, alcance, categorías y vista previa o muestra aleatoria. |
| Calidad de datos | Nulos, duplicados, conversiones, exclusiones y reglas de consistencia. |
| Resumen ejecutivo | KPI, comparación de alcance y evolución de indicadores. |
| Cierre mensual | Agregaciones mensuales, trimestrales o anuales, crecimiento y promedio móvil. |
| Entidades pagadoras | Rankings, comparaciones, Pareto y detalle por entidad. |
| Servicios y líneas | Participación, distribución, concentración y evolución por categoría. |
| Sedes y regímenes | Comparaciones de facturación y otros indicadores disponibles. |
| Estados de facturación | Facturas por estado, distribuciones y tablas cruzadas. |
| Glosas | Frecuencia, severidad, distribución y registros con mayores glosas. |
| Recaudo y cartera | Pagos, saldos, rankings y antigüedad aproximada. |
| Tiempos de radicación | Estadísticos, rangos configurables y casos extremos. |
| Análisis estadístico | Descriptivos, correlaciones y pruebas exploratorias. |
| Valores atípicos | IQR, puntaje Z y percentiles, globales o por categoría. |
| Insights y alertas | Hallazgos calculados, interpretaciones y sugerencias de revisión. |
| Explorador de datos | Búsqueda literal, selección de columnas, orden y vistas especiales. |
| Descargas | Exportaciones del conjunto filtrado y análisis disponibles. |
| Metodología | Fórmulas, supuestos, privacidad y limitaciones. |

La navegación del aplicativo usa algunas etiquetas sin tildes. Los análisis y controles se habilitan según la información disponible.

## Indicadores y metodología

### Indicadores financieros

```text
Facturación neta = suma de Valor_Neto

Porcentaje de glosa = suma de Glosa_Inicial / suma de Valor_Neto × 100

Porcentaje de recaudo = suma de Valor_Pagado / suma de Valor_Neto × 100

Porcentaje de cartera = suma de Saldo_Cartera / suma de Valor_Neto × 100

Ticket promedio = suma de Valor_Neto / número de ID_Factura únicos no nulos
```

Los porcentajes agregados son razones de sumas, no promedios simples de porcentajes individuales. Los denominadores iguales a cero producen un resultado no disponible, no infinito.

### Identidades financieras de revisión

```text
Valor_Neto ≈ Valor_Bruto - Descuento - Copago_Cuota

Saldo_Cartera ≈ Valor_Neto - Glosa_Inicial - Valor_Pagado
```

La tolerancia es configurable. Estas reglas identifican diferencias para revisión; no corrigen importes automáticamente.

**Una glosa inicial no siempre representa una deducción definitiva.** Confirme si la identidad de saldo corresponde al modelo contable del archivo.

### Frecuencia y severidad de glosa

- **Frecuencia:** facturas únicas con algún registro de glosa positiva / total de facturas únicas.
- **Severidad:** valor glosado total / facturación neta total.

No son indicadores equivalentes.

### Periodos y cierres

- Se utiliza `Periodo_Cierre` o el periodo derivado de radicación cuando el primero no existe.
- Los meses se ordenan cronológicamente.
- Los periodos intermedios sin registros se muestran como faltantes, no como facturación cero.
- El crecimiento requiere un periodo anterior consecutivo disponible.
- El promedio móvil requiere tres periodos consecutivos con datos.
- Los deltas del resumen comparan los últimos dos meses consecutivos dentro del filtro; las tarjetas muestran el total filtrado.
- Agrupar pagos por cierre de factura no equivale a agrupar movimientos por `Fecha_Pago`.

### Puntaje orientativo de calidad

Se calcula como promedio simple de cuatro componentes:

1. Completitud de celdas.
2. Unicidad del identificador de factura.
3. Proporción de filas sin fallos de validez detectados.
4. Proporción de filas sin inconsistencias detectadas.

> El puntaje es un indicador exploratorio de calidad y no reemplaza una auditoría de datos.

Las reglas sin columnas suficientes no se evalúan. Un puntaje alto no acredita el cumplimiento de reglas no evaluables ni certifica la calidad del archivo.

### Estadística y correlaciones

La sección incluye estadísticos descriptivos, correlaciones de Pearson o Spearman y las siguientes pruebas:

- Shapiro-Wilk.
- t de Student.
- Mann-Whitney U.
- ANOVA de una vía.
- Kruskal-Wallis.
- Chi-cuadrado de independencia.

Shapiro-Wilk utiliza una muestra reproducible de hasta 5.000 observaciones cuando corresponde. Las pruebas paramétricas incluyen chequeos de normalidad y homogeneidad; la independencia y el diseño de muestreo deben verificarse externamente.

No rechazar la hipótesis nula no demuestra que sea verdadera. No se aplica una corrección automática por múltiples comparaciones.

> La correlación no implica causalidad. Algunas relaciones pueden surgir directamente de las fórmulas contables entre las variables.

### Valores atípicos

El método IQR utiliza:

```text
IQR = Q3 - Q1
Límite inferior = Q1 - multiplicador × IQR
Límite superior = Q3 + multiplicador × IQR
```

El multiplicador inicial es 1,5 y puede modificarse. También se ofrecen puntaje Z y percentiles configurables.

Un valor atípico no es necesariamente un error. Los límites por grupos pequeños pueden ser inestables; el puntaje Z es sensible a distribuciones asimétricas.

### Pareto, alertas y antigüedad

- Pareto utiliza una referencia analítica del 80 %, no una regla obligatoria. Los montos negativos se excluyen de ese cálculo.
- Los umbrales de semáforos son demostrativos y editables; no representan estándares oficiales.
- La antigüedad se calcula desde radicación hasta una fecha de corte seleccionada, aplicada al saldo observado.
- Los registros radicados después del corte se excluyen de esa clasificación.
- El cálculo no reconstruye saldos históricos ni determina vencimientos contractuales.

## Descargas

Según las columnas y registros disponibles, se ofrecen:

- Datos filtrados.
- Cierre mensual.
- Análisis por entidades, servicios y sedes.
- Registros con glosa o cartera.
- Atípicos IQR de valor neto.
- Atípicos de la configuración seleccionada, si se ha utilizado ese módulo.
- Informe de nulos y reglas de calidad.
- Resumen de insights en tabla y texto.

Las tablas se exportan en CSV y, al activar **Preparar Excel**, en `.xlsx`. Los archivos se generan en memoria.

La exportación a Excel informa cuando se supera su límite de filas. En ese caso, utilice CSV. Las cadenas que podrían interpretarse como fórmulas se neutralizan durante la exportación.

## Privacidad y seguridad

> Evite cargar archivos que contengan nombres de pacientes, documentos de identidad, historias clínicas, diagnósticos u otra información personal o sensible. Utilice archivos anonimizados y siga las políticas de tratamiento de datos de su organización.

- No se solicitan datos de pacientes.
- No se envían datos a APIs externas.
- El código no guarda permanentemente los archivos cargados ni los resultados.
- Los originales, las transformaciones y las preferencias se conservan en memoria durante la sesión.
- La caché no persiste en disco, tiene un TTL de 600 segundos y un máximo de 16 entradas por función almacenada en caché.
- Si la versión de Streamlit admite caché por sesión, se utiliza ese alcance. En versiones anteriores se incorpora un token de sesión a las claves de caché para separar resultados entre sesiones.
- Reiniciar la interfaz no debe interpretarse como garantía de borrado inmediato de todas las entradas de caché en versiones anteriores.

**Procesamiento en memoria no significa procesamiento exclusivo en el computador del usuario.** En un despliegue remoto, los datos se procesan en el servidor que ejecuta Streamlit. Revise la autorización institucional y las condiciones del alojamiento antes de cargar información real.

No incorpore datos sensibles al repositorio ni publique capturas o exportaciones con información identificable.

## Limitaciones

- No es un sistema oficial de contabilidad, auditoría ni historia clínica.
- No reconstruye movimientos financieros históricos ni aplica reglas contractuales específicas.
- No convierte monedas ni utiliza tasas de cambio.
- No garantiza que un periodo esté completo.
- Un ID con líneas en varios estados o grupos puede contarse en más de una categoría; los porcentajes por categoría pueden no formar una partición de facturas únicas.
- La clasificación de pago completo o parcial se realiza por registro; un ID con varias líneas puede pertenecer a ambos conjuntos.
- Los soportes incompletos se reconocen mediante valores como `no`, `false`, `0`, `incompleto`, `incompleta` y `n`; otras codificaciones requieren revisión.
- Los insights describen resultados calculados; no determinan fraude, causas ni conductas de pagadores o equipos.
- El rendimiento depende del tamaño del archivo, las categorías y los recursos del servidor. No se garantiza una capacidad máxima universal.
- El límite del cargador depende de la configuración de Streamlit; el código no impone un límite adicional oculto.
- Algunos gráficos limitan categorías o utilizan muestras para mejorar la legibilidad; la interfaz lo indica cuando corresponde.
- Las pruebas locales no equivalen a validar el despliegue en una cuenta concreta ni todos los archivos posibles.

## Solución de problemas

| Situación | Revisión sugerida |
| --- | --- |
| El archivo no abre | Confirme que sea `.xlsx` o `.csv`; para CSV, pruebe otro separador o codificación. |
| La hoja no contiene datos | Seleccione otra hoja y revise que tenga encabezados y registros. |
| Faltan columnas mínimas | Revise el asistente de correspondencia o corrija encabezados en la fuente. |
| El mapeo genera duplicados | Asigne cada columna fuente una sola vez y evite colisiones con encabezados existentes. |
| Los valores monetarios son inesperados | Seleccione explícitamente el formato colombiano o internacional y revise las advertencias. |
| Las fechas son inválidas | Revise el orden día/mes y el contenido original en Calidad de datos. |
| No aparecen cierres | Verifique periodos `AAAA-MM` o active la conversión de fechas para derivarlos desde radicación. |
| Un análisis está desactivado | Revise las columnas necesarias y active las conversiones numéricas correspondientes. |
| Los filtros dejan el conjunto vacío | Pulse **Limpiar filtros** o ajuste las selecciones. |
| Una prueba estadística no se ejecuta | Revise tamaños de grupo, variabilidad, categorías y supuestos; utilice una alternativa no paramétrica cuando corresponda. |
| Los totales parecen duplicados | Compruebe si el total de factura se repite en varias líneas. Corrija el archivo de origen. |
| No se genera Excel | Active **Preparar Excel**; si supera el límite de filas, descargue CSV. |
| El despliegue no inicia | Confirme que `app.py` y `requirements.txt` estén en la raíz y revise los registros de despliegue. |

## Mantenimiento y licencia

Antes de modificar fórmulas o reglas, revise su coherencia con la metodología y pruebe archivos con columnas ausentes, valores inválidos, duplicados y filtros sin resultados.

El proyecto entregado no incluye un archivo `LICENSE`. Si desea distribuirlo con una licencia específica, el propietario debe definirla y agregarla al repositorio. Este README no concede por sí mismo una licencia de uso o redistribución.
