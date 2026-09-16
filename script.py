name: Ejecutar Tabulacion IA

on:
  workflow_dispatch: # Esto permite ejecutarlo manualmente con un botón

jobs:
  run-python-script:
    runs-on: ubuntu-latest # Servidor Linux gratuito

    steps:
      - name: Descargar el repositorio
        uses: actions/checkout@v3

      - name: Instalar Python
        uses: actions/setup-python@v4
        with:
          python-version: '3.10'

      - name: Instalar librerías
        run: |
          pip install pymupdf groq google-genai pandas openpyxl requests Pillow

      - name: Ejecutar el script
        env:
          GEMINI_API_KEY: ${{ secrets.GEMINI_API_KEY }}
          GROQ_API_KEY: ${{ secrets.GROQ_API_KEY }}
          KIMI_API_KEY: ${{ secrets.KIMI_API_KEY }}
          GMAIL_USER: ${{ secrets.GMAIL_USER }}
          GMAIL_APP_PASSWORD: ${{ secrets.GMAIL_APP_PASSWORD }}
        run: python script.py