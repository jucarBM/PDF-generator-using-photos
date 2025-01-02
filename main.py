import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
from reportlab.lib.pagesizes import A4
import threading
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Image, Paragraph, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from PIL import Image as PILImage
import io
from reportlab.lib.enums import TA_CENTER
import math

# Obtener el directorio base (necesario para PyInstaller)


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# El resto de tu código se mantiene igual, solo asegúrate de usar resource_path()
# cuando necesites acceder a archivos empaquetados

# La clase PhotoReportGUI se mantiene igual


def create_photo_report(excel_file, photos_dir, output_dir, rows=3, cols=2, page_size=A4, image_quality=80):
    """
    Genera un PDF por cada pestaña del Excel.

    Parámetros:
    - excel_file: ruta al archivo Excel
    - photos_dir: carpeta donde están las fotos
    - output_dir: carpeta donde se guardarán los PDFs generados
    - rows: número de filas por página
    - cols: número de columnas por página
    - page_size: tamaño de página (por defecto A4)
    - image_quality: calidad de las imágenes (0-100)
    """

    # Crear directorio de salida si no existe
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Leer Excel
    xlsx = pd.ExcelFile(excel_file)

    # Estilos
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=30,
        alignment=TA_CENTER
    )

    caption_style = ParagraphStyle(
        'Caption',
        parent=styles['Normal'],
        fontSize=10,
        alignment=TA_CENTER,
        spaceAfter=5
    )

    # Procesar cada hoja del Excel
    for sheet_name in xlsx.sheet_names:
        # Crear nombre del archivo PDF para esta hoja
        pdf_name = f"Visitas_{sheet_name}.pdf"
        pdf_path = os.path.join(output_dir, pdf_name)

        # Crear documento PDF
        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=page_size,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )

        elements = []
        df = pd.read_excel(excel_file, sheet_name=sheet_name)

        # Agregar título
        elements.append(Paragraph(f"Visitas {sheet_name}", title_style))
        elements.append(Paragraph("<br/>", styles['Normal']))

        # Procesar fecha de visita si existe
        df = df.dropna(how='all').reset_index(drop=True)
        if not df.empty and isinstance(df.iloc[0, 0], str) and "Fecha de visita" in df.iloc[0, 0]:
            elements.append(Paragraph(df.iloc[0, 0], caption_style))
            df = df.iloc[1:].reset_index(drop=True)

        photos_data = []
        for index, row in df.iterrows():
            if len(row) >= 2:
                nombre = str(row.iloc[0])
                codigo = str(row.iloc[1])

                if nombre and codigo and "Visitas" not in nombre:
                    photo_path = None
                    for filename in os.listdir(photos_dir):
                        if codigo in filename:
                            photo_path = os.path.join(photos_dir, filename)
                            break

                    if photo_path and os.path.exists(photo_path):
                        with PILImage.open(photo_path) as img:
                            # Rotar 180 grados si está de cabeza
                            img = img.rotate(180, expand=True)

                            # Rotar 270 grados adicionales si está en horizontal
                            if img.width > img.height:
                                img = img.rotate(270, expand=True)

                            # Comprimir imagen
                            img_byte_arr = io.BytesIO()
                            img.save(img_byte_arr, format='JPEG',
                                     quality=image_quality)
                            img_byte_arr = img_byte_arr.getvalue()

                            # Calcular dimensiones
                            available_height = (page_size[1] - 2*72) / rows
                            target_height = min(
                                3 * inch, available_height * 0.8)
                            aspect = img.width / img.height
                            new_width = target_height * aspect

                            caption = Paragraph(nombre, caption_style)
                            img_element = Image(io.BytesIO(
                                img_byte_arr), width=new_width, height=target_height)
                            photos_data.append([caption, img_element])

        if photos_data:
            # Calcular páginas necesarias
            photos_per_page = rows * cols
            num_pages = math.ceil(len(photos_data) / photos_per_page)

            for page in range(num_pages):
                start_idx = page * photos_per_page
                end_idx = min((page + 1) * photos_per_page, len(photos_data))
                page_photos = photos_data[start_idx:end_idx]

                table_data = []
                for i in range(0, len(page_photos), cols):
                    row_photos = page_photos[i:i + cols]
                    row_captions = []
                    row_images = []

                    while len(row_photos) < cols:
                        row_photos.append(['', ''])

                    for photo in row_photos:
                        row_captions.append(photo[0])
                        row_captions.append('')
                        row_images.append(photo[1])
                        row_images.append('')

                    table_data.append(row_captions[:-1])
                    table_data.append(row_images[:-1])

                if table_data:
                    table = Table(table_data, colWidths=[
                                  2.5*inch, 0.2*inch] * cols)
                    table.setStyle(TableStyle([
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('LEFTPADDING', (0, 0), (-1, -1), 5),
                        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                        ('TOPPADDING', (0, 0), (-1, -1), 2),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                    ]))
                    elements.append(table)

                    if page < num_pages - 1:
                        elements.append(PageBreak())

        # Generar PDF para esta hoja
        doc.build(elements)
        print(f"PDF generado: {pdf_path}")


class PhotoReportGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Informes Fotográficos")
        self.root.geometry("500x550")

        # Variables
        self.excel_path = tk.StringVar()
        self.photos_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.rows = tk.IntVar(value=3)
        self.cols = tk.IntVar(value=2)
        self.quality = tk.IntVar(value=80)

        # Crear interfaz
        self.create_widgets()

    # Métodos de navegación
    def browse_excel(self):
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")],
            initialdir=os.path.expanduser("~")
        )
        if filename:
            self.excel_path.set(filename)

    def browse_photos(self):
        dirname = filedialog.askdirectory(
            title="Seleccionar carpeta de fotos",
            initialdir=os.path.expanduser("~")
        )
        if dirname:
            self.photos_dir.set(dirname)

    def browse_output(self):
        dirname = filedialog.askdirectory(
            title="Seleccionar carpeta de salida",
            initialdir=os.path.expanduser("~")
        )
        if dirname:
            self.output_dir.set(dirname)

    def validate_inputs(self):
        if not self.excel_path.get():
            messagebox.showerror(
                "Error", "Por favor seleccione un archivo Excel")
            return False
        if not self.photos_dir.get():
            messagebox.showerror(
                "Error", "Por favor seleccione la carpeta de fotos")
            return False
        if not self.output_dir.get():
            messagebox.showerror(
                "Error", "Por favor seleccione la carpeta de salida")
            return False
        if not os.path.exists(self.excel_path.get()):
            messagebox.showerror(
                "Error", "El archivo Excel seleccionado no existe")
            return False
        if not os.path.exists(self.photos_dir.get()):
            messagebox.showerror(
                "Error", "La carpeta de fotos seleccionada no existe")
            return False
        return True

    def generate_reports(self):
        if not self.validate_inputs():
            return

        self.generate_button.state(['disabled'])
        self.status_var.set("Generando PDFs...")
        self.progress_var.set(0)

        thread = threading.Thread(target=self._generate_reports_thread)
        thread.start()

    def _generate_reports_thread(self):
        try:
            create_photo_report(
                excel_file=self.excel_path.get(),
                photos_dir=self.photos_dir.get(),
                output_dir=self.output_dir.get(),
                rows=self.rows.get(),
                cols=self.cols.get(),
                image_quality=self.quality.get()
            )
            self.root.after(0, lambda: self.status_var.set(
                "PDFs generados exitosamente"))
            messagebox.showinfo(
                "Éxito", "Los PDFs han sido generados correctamente")
        except Exception as e:
            self.root.after(0, lambda: self.status_var.set(f"Error: {str(e)}"))
            messagebox.showerror("Error", f"Error al generar PDFs:\n{str(e)}")
        finally:
            self.root.after(
                0, lambda: self.generate_button.state(['!disabled']))
            self.root.after(0, lambda: self.progress_var.set(100))

    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configurar expansión de grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        # Selector de Excel
        ttk.Button(
            main_frame,
            text="Seleccionar Archivo Excel",
            command=self.browse_excel,
            style='TButton',
            width=40
        ).grid(row=0, column=0, pady=(0, 5))

        ttk.Label(main_frame, textvariable=self.excel_path, wraplength=400).grid(
            row=1, column=0, pady=(0, 15)
        )

        # Selector de carpeta de fotos
        ttk.Button(
            main_frame,
            text="Seleccionar Carpeta de Fotos",
            command=self.browse_photos,
            style='TButton',
            width=40
        ).grid(row=2, column=0, pady=(0, 5))

        ttk.Label(main_frame, textvariable=self.photos_dir, wraplength=400).grid(
            row=3, column=0, pady=(0, 15)
        )

        # Selector de carpeta de salida
        ttk.Button(
            main_frame,
            text="Seleccionar Carpeta de Salida",
            command=self.browse_output,
            style='TButton',
            width=40
        ).grid(row=4, column=0, pady=(0, 5))

        ttk.Label(main_frame, textvariable=self.output_dir, wraplength=400).grid(
            row=5, column=0, pady=(0, 15)
        )

        # Frame de configuración
        config_frame = ttk.LabelFrame(
            main_frame, text="Configuración", padding="10")
        config_frame.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=(0, 15))

        # Filas
        ttk.Label(config_frame, text="Filas por página:").grid(
            row=0, column=0, sticky=tk.W, padx=5)
        ttk.Spinbox(config_frame, from_=1, to=10, width=5, textvariable=self.rows).grid(
            row=0, column=1, sticky=tk.W, padx=5, pady=5
        )

        # Columnas
        ttk.Label(config_frame, text="Columnas por página:").grid(
            row=1, column=0, sticky=tk.W, padx=5)
        ttk.Spinbox(config_frame, from_=1, to=10, width=5, textvariable=self.cols).grid(
            row=1, column=1, sticky=tk.W, padx=5, pady=5
        )

        # Calidad
        ttk.Label(config_frame, text="Calidad de imagen:").grid(
            row=2, column=0, sticky=tk.W, padx=5)
        quality_frame = ttk.Frame(config_frame)
        quality_frame.grid(row=2, column=1, sticky=(tk.W, tk.E))

        ttk.Scale(
            quality_frame,
            from_=1,
            to=100,
            orient=tk.HORIZONTAL,
            variable=self.quality,
            length=200
        ).grid(row=0, column=0)

        ttk.Label(quality_frame, textvariable=self.quality).grid(
            row=0, column=1, padx=5)

        # Barra de progreso
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(
            main_frame,
            length=300,
            mode='determinate',
            variable=self.progress_var
        )
        self.progress.grid(row=7, column=0, pady=(0, 10), sticky=(tk.W, tk.E))

        # Botón de generación
        generate_button_style = ttk.Style()
        generate_button_style.configure('Generate.TButton', padding=10)

        self.generate_button = ttk.Button(
            main_frame,
            text="Generar PDFs",
            command=self.generate_reports,
            style='Generate.TButton',
            width=30
        )
        self.generate_button.grid(row=8, column=0, pady=(0, 10))

        # Estado
        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(
            main_frame,
            textvariable=self.status_var,
            wraplength=400
        )
        self.status_label.grid(row=9, column=0)


def main():
    try:
        root = tk.Tk()
        app = PhotoReportGUI(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror(
            "Error", f"Error al iniciar la aplicación:\n{str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
