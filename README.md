# Generador de Actas - Universidad de Córdoba

MVP gratuito para:

- Subir una citación en PDF.
- Extraer fecha, hora, lugar, orden del día y personas citadas/invitadas.
- Marcar manualmente asistencia, inasistencia o excusa.
- Subir audio (MP3, WAV, M4A, MP4, AAC, OGG, FLAC) o una transcripción (TXT/DOCX/PDF).
- Transcribir audio con `faster-whisper`.
- Redactar el acta con un modelo local/gratuito en Google Colab.
- Generar un Word usando la plantilla institucional FGDC-025 suministrada.

## Arquitectura

- `docs/`: frontend estático publicable en GitHub Pages.
- `backend/`: API FastAPI para extracción, transcripción, redacción y Word.
- `colab/`: notebook para iniciar gratis el backend desde Google Colab.

## Uso rápido

1. Sube este repositorio a GitHub.
2. En GitHub activa **Settings > Pages > Deploy from a branch**, rama `main`, carpeta `/docs`.
3. Abre `colab/Backend_Actas_Colab.ipynb` en Google Colab y ejecuta todas las celdas.
4. Colab mostrará una URL pública temporal de Cloudflare Tunnel.
5. Abre la página de GitHub Pages y pega esa URL en **Servidor de procesamiento**.
6. Sube la citación, marca asistencia, sube audio o transcripción y genera el acta.

## Importante

Colab gratuito no es un servidor permanente. Cada vez que la sesión termine, hay que ejecutar nuevamente el notebook y, normalmente, pegar una nueva URL pública en la página.

La plantilla Word original está en `backend/assets/template_acta.docx` y el generador trabaja sobre ella para conservar encabezado, tablas, logos y estructura institucional.
