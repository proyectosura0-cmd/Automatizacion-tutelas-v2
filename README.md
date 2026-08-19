# Automatizador de Tutelas — EPS SURA

Aplicación web para el Área Jurídica de EPS SURA que automatiza la generación
de contestaciones de acciones de tutela en formato .docx.

## Requisitos

- Python 3.10 o superior
- pip

## Instalación

```bash
cd tutelas-sura
pip install -r requirements.txt
```

## Configuración de plantillas

Copie los archivos .docx de plantillas en la carpeta `plantillas/`:

| Modelo                        | Nombre del archivo            |
|-------------------------------|-------------------------------|
| Transporte                    | Transporte.docx               |
| Medicamento Importado         | Medicamento_Importado.docx    |
| No INVIMA                     | No_INVIMA_2025.docx           |
| No PBS / Presupuestos Máximos | NO_PBS_Presup_Max.docx        |
| Silla de Ruedas               | Silla_de_Ruedas.docx          |
| Carencia Actual de Objeto     | Carencia_Actual_de_Objeto.docx|

También puede subir las plantillas directamente desde la interfaz web
en la pestaña **Plantillas**.

## Marcadores en las plantillas .docx

Las plantillas deben contener marcadores con el siguiente formato:

```
[ACCIONANTE]       [JUZGADO]          [RADICADO]
[NUMERO_RS]        [FECHA_CONTESTACION] [REPRESENTANTE_LEGAL]
[SERVICIO_SOLICITADO] [RESUMEN_HECHOS] [PRETENSIONES]
[SOLICITUD_FINAL]  [DIAGNOSTICO]      etc.
```

Lista completa en `services/word_generator.py` → función `_construir_mapa()`.

## Ejecución

```bash
python main.py
```

Abrir en el navegador: **http://localhost:8000**

La documentación de la API estará disponible en: **http://localhost:8000/docs**

## Estructura del proyecto

```
tutelas-sura/
├── main.py                     # FastAPI + uvicorn
├── database.py                 # SQLite con SQLAlchemy
├── models.py                   # Modelo ORM Caso
├── requirements.txt
├── routers/
│   ├── tutelas.py              # POST /generar, GET /descargar, plantillas
│   └── casos.py                # CRUD de casos + exportar CSV
├── services/
│   └── word_generator.py       # Lógica de reemplazo de marcadores
├── plantillas/                 # .docx de plantillas (manual o desde UI)
├── contestaciones/             # .docx generados
├── static/
│   └── index.html              # SPA completa
└── tutelas.db                  # Base de datos SQLite (auto-creada)
```

## Notas técnicas

- El reemplazador de marcadores concatena todos los `runs` del párrafo antes de
  buscar marcadores, lo que resuelve el problema de Word fragmentando el texto
  en múltiples `<w:r>` elementos.
- Los documentos generados incluyen al inicio una advertencia en rojo y negrita:
  `*** BORRADOR — PARA REVISIÓN DEL ABOGADO RESPONSABLE — NO RADICAR ***`
- Los archivos generados se guardan en `./contestaciones/` y son descargables
  desde la interfaz o directamente vía `/api/tutelas/descargar/{nombre}`.
