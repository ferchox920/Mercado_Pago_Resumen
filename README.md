# Mercado_Pago_Resumen

Herramienta en Python para leer resúmenes de cuenta de Mercado Pago en PDF y generar archivos Excel filtrados.

## Uso

```bash
pip install pymupdf openpyxl
python mercadopago_pdf_a_excel.py archivo.pdf -o movimientos.xlsx
python mercadopago_pdf_a_excel.py archivo1.pdf archivo2.pdf --output-dir carpeta_salida
```

Sin argumentos abre la interfaz visual:

```bash
python mercadopago_pdf_a_excel.py
```

## Ejecutable

```bash
pyinstaller --onefile --noconsole mercadopago_pdf_a_excel.py
```
