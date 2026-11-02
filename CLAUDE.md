# Contexto para Claude: proyecto BCAT-YOGUR

Trabajo de clase (gestión de proyectos, UJI). Para cada tema de la asignatura hay que hacer un documento sobre el mismo proyecto inventado. Marc deja los apuntes del tema (PDF) en una carpeta y pide el documento.

## El proyecto (datos que deben cuadrar en todos los documentos)

- **Nombre:** Sustitución de resistencias eléctricas por bombas de calor de alta temperatura (CO₂, R744) en el proceso de fabricación de yogur. **Acrónimo:** BCAT-YOGUR.
- **Qué:** 2 bombas de calor de ~175 kWt que sustituyen a las resistencias en pasteurización (90–95 °C), incubación (45–50 °C, ±1 °C) y CIP; recuperan calor de la instalación frigorífica; depósitos de acumulación; red hidráulica por niveles de temperatura; control PLC/SCADA. Las resistencias se quedan como respaldo.
- **Promotor:** [Empresa láctea] – Dirección General. **Cliente:** Producción y Mantenimiento. Nombres de empresa y personas siguen entre corchetes.
- **Plazo:** nov 2026 – mar 2028 (17 meses). Puesta en marcha antes del 31/12/2027. Obras en la línea solo en la parada de agosto de 2027.
- **Hitos:** H0 nov 26 acta · H1 ene 27 auditoría · H2 feb 27 integración térmica y tecnología · H3 abr 27 proyecto ejecutivo · H4 may 27 adjudicación y pedido · H5 sep 27 auxiliares y adaptación de consumidores (parada de agosto) · H6 nov 27 bombas y depósitos montados · H7 dic 27 puesta en marcha y acta de recepción · H8 mar 28 operación asistida, ahorros y cierre.
- **Presupuesto (sin IVA): 830.500 €** = ingeniería 45.000 · bombas de calor 330.000 · integración frigorífica 60.000 · depósitos 55.000 · red hidráulica 70.000 · adaptación de consumidores 80.000 · eléctrica y control 50.000 · montaje, puesta en marcha y formación 40.000 · dirección 25.000 · contingencia 10 % 75.500.
- **Objetivos:** ≥95 % del calor del proceso con bomba de calor, ahorro ≥65 % de la electricidad usada para calor (~1.700 MWh/año, ~255.000 €/año), retorno ≤4 años (estimado 3,3), desviación de coste ≤10 %, desviación de plazo ≤1 mes, 0 paradas no programadas.
- **Director del proyecto:** aprueba hasta 25.000 € por partida, usa la contingencia informando al promotor y puede mover fechas sin tocar H3, H7 ni H8. Lo demás lo aprueba la Dirección General.
- **Interesados (13, numeración fija):** internos 1 Dirección General (promotor), 2 Director del proyecto, 3 Producción, 4 Mantenimiento, 5 Operarios de planta; externos 6 Ingeniería de frío industrial, 7 Fabricante de las bombas de calor, 8 Empresa instaladora, 9 Servicio Territorial de Industria, 10 Clientes (cadenas de distribución), 11 Vecinos y empresas del polígono, 12 Organismo de control (OCA), 13 Compañía eléctrica.
  - Marc **quitó a propósito**: Calidad, comité de empresa/PRL, finanzas, banco, ayuntamiento, IDAE/IVACE y gestor de CAE. No los vuelvas a meter.
  - El acta todavía menciona ayudas/CAE en entregables, H3 y riesgos. Es una incoherencia conocida y pendiente de que Marc decida qué hacer; en los documentos nuevos no se incluyen ayudas ni CAE.
- **EDT (ver `EDT/generar_edt.py`):** 1 Gestión del proyecto · 2 Estudios previos · 3 Proyecto ejecutivo y legalización · 4 Instalación del sistema (4.1–4.7) · 5 Puesta en marcha y formación · 6 Verificación de resultados. Paquetes de trabajo 755.000 € + contingencia 75.500 €. Requisitos R01–R09 trazados a paquetes de la EDT.

Si un documento nuevo necesita un dato que no está aquí, sácalo del acta (`Acta de constitución/`) y mantén la coherencia con lo anterior.

## Estructura del repositorio

```
Plantilla_documentos_BCAT-YOGUR.docx   plantilla de formato (Marc la retoca en ONLYOFFICE)
Acta de constitución/                  acta (editada a mano, no tiene script)
Gestión de interesados/                documento + apuntes + matriz_interesados.py (genera la matriz)
EDT/                                   documento + apuntes + generar_edt.py (genera diagrama y documento)
_herramientas/docx_plantilla.py        clase Documento: abre la plantilla, limpia el ejemplo, añade títulos, tablas, figuras y pies numerados
README.md                              para los compañeros (María Jesús y Pablo), muy sencillo y cercano
```

- Una carpeta por documento, con los apuntes del tema dentro.
- Nombre del archivo: `Tema_BCAT-YOGUR.docx` (por ejemplo `Gestion_alcance_EDT_BCAT-YOGUR.docx`), sin "_v2" ni "_final".

## Cómo hacer un documento nuevo

1. Lee los apuntes del PDF de la carpeta. Se pueden leer con Read sin el parámetro `pages`, porque no hay `pdftoppm`. Sigue más o menos su estructura: los apartados del tema, las tablas que proponen y el ejemplo de clase.
2. Escribe un `generar_<tema>.py` en la carpeta usando `_herramientas/docx_plantilla.py` (ver `EDT/generar_edt.py` como modelo) y genera el `.docx`.
3. Valida con el script `validate.py` de la skill docx (`--original Plantilla_documentos_BCAT-YOGUR.docx`).
4. Si hay gráficas, hazlas con matplotlib con la misma estética: Arial y escala de grises (#404040, #D9D9D9, #F2F2F2, #7F7F7F), sin colores llamativos. Revisa la imagen con Read antes de dar el trabajo por acabado.
5. **Si un .docx ya lo ha retocado alguien a mano, no lo regeneres con el script:** edítalo directamente (python-docx o XML) para no pisar sus cambios.

## Estilo de los textos

- Español natural, como lo escribiría un estudiante. **Que no parezca hecho con IA**: frases cortas, palabras corrientes, sin flechas "→", sin "≈" (mejor "unos"), sin listas de sustantivos encadenados ni frases grandilocuentes.
- Marc prefiere documentos **sin relleno**. No hace falta introducción ni comentarios que repitan lo que ya dicen las tablas o la figura; ve directamente a lo que piden los apuntes.
- Usa los estilos de la plantilla: Heading 1/2/3 numerados ("1.", "1.1."), Normal, List Bullet, List Number, Etiqueta, Texto tabla, Cabecera tabla, Caption, Nota y Texto secundario.

## Detalles técnicos

- Windows, con Git Bash y PowerShell. Hay Python con python-docx, matplotlib, lxml y defusedxml. **No hay** Node, Word, LibreOffice, pandoc ni pdftoppm, así que no se pueden renderizar los .docx: avisa a Marc de que lo revise.
- En Git Bash usa `PYTHONIOENCODING=utf-8`; la consola es cp1252 y las tildes rompen los `print`.
- La plantilla está guardada con ONLYOFFICE:
  - `doc.styles["Heading 1"]` falla; busca los estilos recorriendo `doc.styles` por nombre (ya lo hace `Documento`).
  - Las tablas de la plantilla guardan los elementos fuera del orden del esquema; `Documento` los reordena.
- Al añadir XML a mano respeta el orden del esquema (pPr, tcPr, tblPr, tblCellMar: top, left, bottom, right) o la validación falla.
- Git: Marc hace los commits y el push. No hagas commit ni push si no te lo pide.
