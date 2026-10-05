# Mercado Pago · Resúmenes a Excel

**Convertí tus resúmenes de cuenta en PDF en planillas Excel organizadas y listas para revisar.**

Herramienta en Python que extrae **transferencias recibidas** y **liquidaciones de dinero** de los resúmenes de Mercado Pago. Podés usarla desde una interfaz gráfica o desde la terminal, con un archivo o varios a la vez.

Los documentos se procesan localmente en tu computadora.

## Funciones

- Selección de varios PDFs y una carpeta de destino desde la interfaz gráfica.
- Un Excel independiente por cada PDF, con movimientos, categorías y totales.
- Datos del titular y del período en las hojas de movimientos.
- Filtros, encabezados fijos y formato de moneda para facilitar la revisión.
- Aviso cuando el total exportado difiere de las entradas detectadas en el resumen.
- Registro del resultado de cada archivo y una barra de progreso en la interfaz.
- Nombres numerados para evitar sobrescribir archivos al generar por carpeta.

## Qué contiene el Excel

Cada archivo incluye cuatro hojas:

| Hoja | Contenido |
| --- | --- |
| **Movimientos** | Todos los movimientos incluidos en la exportación. |
| **Liquidación de dinero** | Movimientos de esta categoría. |
| **Transferencia recibida** | Transferencias recibidas y su contraparte. |
| **Resumen** | Cantidad y total por categoría, más el total general. |

Las hojas de movimientos contienen **fecha, tipo, descripción, contraparte y valor**.

> La exportación incluye únicamente las dos categorías indicadas. Una diferencia con el total de entradas puede deberse a otros tipos de ingreso presentes en el PDF; revisá el aviso y el resumen original.

## Instalación

Necesitás Python **3.10 o superior**. La interfaz gráfica utiliza Tkinter, que debe estar disponible en tu instalación de Python.

Cloná el repositorio e instalá las dependencias:

```bash
git clone https://github.com/ferchox920/Mercado_Pago_Resumen.git
cd Mercado_Pago_Resumen
python -m pip install pymupdf openpyxl
```

## Uso con interfaz gráfica

Ejecutá el programa sin argumentos:

```bash
python mercadopago_pdf_a_excel.py
```

1. Presioná **Seleccionar PDFs** y elegí los resúmenes que quieras procesar.
2. Presioná **Seleccionar carpeta destino**.
3. Presioná **Generar Excels**.
4. Revisá el registro de resultados y los archivos de la carpeta elegida.

Durante la generación, los botones de selección quedan bloqueados para mantener los archivos y el destino elegidos.

## Uso desde la terminal

### Un PDF con un nombre de salida específico

```bash
python mercadopago_pdf_a_excel.py "resumen.pdf" -o "movimientos.xlsx"
```

Si el archivo indicado con `-o` ya existe, se sobrescribe.

### Varios PDFs, un Excel por archivo

```bash
python mercadopago_pdf_a_excel.py "resumen_enero.pdf" "resumen_febrero.pdf" --output-dir "salida"
```

Los archivos se guardan como `resumen_enero_movimientos.xlsx` y `resumen_febrero_movimientos.xlsx`. Si un nombre ya existe, se agrega un sufijo como `_2`.

Si omitís `--output-dir`, los Excel se guardan en la carpeta desde la que ejecutás el comando. La opción `-o` admite un solo PDF y no se puede combinar con `--output-dir`.

Para consultar las opciones:

```bash
python mercadopago_pdf_a_excel.py --help
```

## Crear un ejecutable para Windows

Desde Windows, instalá PyInstaller y compilá el programa:

```bash
python -m pip install pyinstaller
python -m PyInstaller --onefile --noconsole mercadopago_pdf_a_excel.py
```

El ejecutable se genera en `dist/mercadopago_pdf_a_excel.exe`. Volvé a compilarlo después de actualizar el código para incorporar los cambios.

## Pruebas

Con las dependencias instaladas, ejecutá:

```bash
python -m unittest -v
```

Las pruebas usan datos ficticios y verifican la lectura de importes, el control de entradas, la conversión de PDF a Excel y la conservación de la selección durante el procesamiento.

## Documentos y archivos locales

El repositorio contiene el código, las pruebas y la documentación. Los PDFs, Excel generados, carpetas locales de documentos, archivos de compilación y accesos directos de Windows están excluidos mediante `.gitignore`.

## Alcance

El extractor trabaja con el texto del PDF y con la estructura de movimientos que reconoce el programa. No incluye OCR para documentos escaneados; un cambio en el formato del resumen puede requerir ajustar la extracción.

Proyecto independiente, sin afiliación con Mercado Pago.
