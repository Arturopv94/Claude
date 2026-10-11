# Práctica de Excel · Analyst, Customer Success Operations

Generador de libros de Excel para preparar la vacante
[Analyst, Customer Success Operations (Mitratech, remoto México)](https://job-boards.greenhouse.io/mitratech/jobs/8226722):
reportes de uso y salud de cuentas, health score, auditoría de reportes, calidad de datos y apoyo a CSMs.

Los datos son ficticios pero imitan a una empresa SaaS B2B: cuentas, CSMs, uso mensual del producto,
tickets de soporte, NPS, catálogo de reportes y un log de quién consulta cada reporte.

## Uso rápido

```bash
pip install xlsxwriter
python generar_practica_excel.py
```

Se crea la carpeta `salida/` con:

| Archivo | Para qué |
|---|---|
| `CS_Ops_Practica.xlsx` | Ejercicios: llenas las celdas **amarillas** y las columnas ✓/✗ te dicen si vas bien |
| `CS_Ops_Soluciones.xlsx` | El mismo libro (mismos datos) con todas las fórmulas resueltas |
| `csv/` | Las tablas en CSV para practicar Power Query |

Si no quieres correr nada, en `archivos_generados/` hay un set ya generado (semilla 42).

Opciones:

```bash
python generar_practica_excel.py --semilla 7          # datos nuevos para volver a practicar
python generar_practica_excel.py --cuentas 300        # más cuentas
python generar_practica_excel.py --fecha-corte 2026-06-30 --salida junio
```

## Ejercicios

| Hoja | Qué practicas | Relación con la vacante |
|---|---|---|
| **E1_Lookups** | VLOOKUP, INDEX/MATCH, XLOOKUP anidado, búsqueda aproximada (tiers), búsqueda de abajo hacia arriba, TRIM, IFERROR | Completar con datos del CRM las listas que mandan los CSMs |
| **E2_Agregaciones** | SUMIFS, COUNTIFS, AVERAGEIFS, MAXIFS, MINIFS, SUMPRODUCT, NPS | KPIs para liderazgo: ARR por segmento/región, book of business, backlog de soporte |
| **E3_Health_Score** | Health score ponderado (adopción, tendencia, soporte, NPS), IFS, SUMPRODUCT, RANK.EQ, formato condicional, nombres definidos | *"Build and refine health-score and KPI frameworks"* |
| **E4_Auditoria_Reportes** | Uso real por reporte, usuarios únicos (UNIQUE/FILTER), recencia, costo por vista, recomendación retirar / consolidar / automatizar / mantener | Auditar los reportes actuales: qué quitar, qué mejorar, qué agregar |
| **E5_Limpieza_Datos** | TRIM, PROPER, SWITCH, NUMBERVALUE, fechas en texto, duplicados, conciliación contra el maestro | Calidad de datos antes de reportar |
| **E6_Pivots** | 6 tablas dinámicas guiadas (agrupar por mes, % del total, campo calculado, slicers, GETPIVOTDATA) | Reportes rápidos y ad hoc |
| **E7_Dinamicas_365** | FILTER, UNIQUE, SORT, XLOOKUP de varias columnas, TEXTJOIN, LET, SEQUENCE, HSTACK | Reportes que se actualizan solos |
| **E8_Tendencias** | Uso mes a mes, MoM %, promedio móvil, churn, GRR, pipeline de renovaciones con ARR en riesgo | QBR/EBR y renovaciones |
| **E9_Dashboard_CSM** | Lista desplegable, KPIs dinámicos, top 5 con LET/FILTER/SORTBY/TAKE, gráfica | Tablero para cada CSM |
| E10_PowerQuery | Importar/combinar CSV, limpiar y agrupar sin fórmulas | Automatizar la ingesta de datos |
| Entrevista | Preguntas técnicas y de negocio (en inglés) | Practicar en voz alta |

La hoja **LEEME** muestra tu avance por ejercicio y tiene una tabla de traducción de funciones
inglés → español (por ejemplo XLOOKUP = BUSCARX) por si tu Excel está en español.

## Notas

- Requiere Excel 365 o Excel 2021+. LET, TAKE y HSTACK (en E7 y E9) solo existen en Excel 365.
- La fecha de corte (`AsOfDate`) es fija dentro del libro en lugar de `TODAY()`, así los resultados se pueden verificar.
- Las verificaciones ✓/✗ asumen los pesos y parámetros por defecto. Si los cambias para experimentar, es normal que algunas pasen a ✗.
