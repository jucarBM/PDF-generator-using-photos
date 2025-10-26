# Generador de Informes Fotográficos para Analistas Financieros

Este proyecto es un prototipo de escritorio que ayuda a los analistas financieros a consolidar en PDF las fotografías recopiladas durante sus visitas a clientes. La herramienta enlaza la evidencia visual con la información registrada en un archivo Excel y organiza todo en informes listos para compartir. Es un recurso de prueba y puede utilizarse libremente.

## Características principales

- **Automatiza la generación de informes**: crea un PDF por cada pestaña del libro de Excel, incorporando las imágenes asociadas a cada registro.
- **Tratamiento automático de imágenes**: corrige la orientación, ajusta el tamaño y comprime cada fotografía antes de añadirla al informe.
- **Interfaz gráfica simple**: permite seleccionar el archivo Excel, la carpeta de fotos, la carpeta de salida y los parámetros de maquetación (filas, columnas y calidad).
- **Procesamiento en segundo plano**: ejecuta la generación en un hilo dedicado para mantener la interfaz receptiva mientras avanza la barra de progreso.

## Requisitos

- Python 3.8 o superior
- Dependencias listadas en `requirements.txt`

Instale las dependencias con:

```bash
pip install -r requirements.txt
```

## Uso

1. Ejecute la aplicación:

   ```bash
   python main.py
   ```

2. Desde la ventana principal:
   - Seleccione el archivo de Excel que contiene los registros de visitas.
   - Elija la carpeta donde se encuentran las fotografías tomadas durante las visitas.
   - Indique la carpeta donde se guardarán los PDFs generados.
   - Ajuste el número de filas, columnas y la calidad de imagen si lo desea.

3. Presione **Generar PDFs** y espere a que la barra de progreso indique la finalización. Los informes se guardarán en la carpeta de salida, con un archivo por cada hoja del Excel.

## Licencia

Este proyecto es de uso libre para fines de evaluación y pruebas internas. Puede copiarlo, modificarlo y distribuirlo sin restricciones.
