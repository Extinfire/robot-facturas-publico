# Robot Facturas SAT — Bot de Telegram

Bot de Telegram para registrar facturas (FEL Guatemala): recibe fotos o PDFs de facturas, extrae los datos con OCR (Claude o Ollama) y lectura de QR, los clasifica y genera reportes en Excel.

## Requisitos

- Python 3.10+
- Un bot de Telegram (créalo con [@BotFather](https://t.me/BotFather) y copia el token)
- Una API key de Anthropic (para OCR con Claude) **o** [Ollama](https://ollama.com) instalado localmente con un modelo de visión
- `plantilla_reporte.xlsx`: la plantilla de Excel para los reportes **no está incluida** en este repositorio. El bot necesita un libro con una hoja llamada `Masivo` donde escribe las facturas. Pídesela a quien te compartió el bot o crea la tuya.

## Instalación

**Windows:**

```bat
setup.bat
```

**Linux / macOS:**

```bash
./setup.sh
```

Luego copia la configuración de ejemplo y llena tus datos:

```bash
cp .env.ejemplo .env
```

Edita `.env`:

| Variable | Descripción |
|---|---|
| `TELEGRAM_BOT_TOKEN` | Token de tu bot (de @BotFather) |
| `ANTHROPIC_API_KEY` | API key de Anthropic para OCR (opcional si usas Ollama) |
| `OLLAMA_BASE_URL` / `OLLAMA_MODEL` | Config de Ollama para OCR local |
| `ADMIN_CHAT_ID` | Tu chat_id de Telegram para recibir alertas (envía un mensaje a @userinfobot para obtenerlo) |

## Uso

**Windows:** `iniciar.bat` — **Linux/macOS:** `python telegram_bot.py`

Comandos principales dentro de Telegram:

- Envía una **foto o PDF** de una factura y el bot la registra
- `/reporte` — genera el Excel de reportes
- `/buscar` — busca facturas por empresa, NIT, fecha o monto
- `/pendientes` — revisa facturas que fallaron en la extracción
- `/guia` — genera la guía de usuario en PDF

## Estructura

| Archivo | Qué hace |
|---|---|
| `telegram_bot.py` | Bot de Telegram: handlers, comandos, flujo de revisión |
| `sat_invoice_bot.py` | Extracción de datos de facturas (OCR + heurísticas) |
| `sat_verificador.py` | Verificación de facturas contra el portal de la SAT |
| `lector_qr.py` | Lectura del QR FEL de las facturas |
| `clasificacion.py` | Clasificación de gastos |
| `reportes.py` | Generación del Excel a partir de la plantilla |
| `comun.py` | Utilidades compartidas (perfiles de usuario, deduplicación) |
| `respaldo_nas.py` | Respaldo automático de la carpeta `datos/` a un NAS |
| `test_unitarios.py` | Tests (`pytest test_unitarios.py`) |

Los datos de usuarios se guardan en `datos/` y el catálogo de empresas en `data/empresas.json` (ambos ignorados por git; cada instalación mantiene los suyos).
