import os
from docxtpl import DocxTemplate
import pandas as pd
from docx2pdf import convert
import datetime

def generar_certificados(csv_file, template_file, output_folder):
    # Cargar el archivo CSV en un DataFrame
    df = pd.read_csv(csv_file)

    # Recorrer cada fila del DataFrame
    for _, row in df.iterrows():
        current_date = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Obtener los valores de la fila
        typecert            = row['type']
        date                = row['date']
        company             = row['company']
        equipment           = row['equipment']
        reference           = row['reference']
        id_code             = row['id_code']
        tag_code            = row['tag_code']
        brand               = row['brand']
        model               = row['model']
        fabrication_date    = row['fabrication_date']
        batch_number        = row['batch_number']
        concept             = row['concept']
        observations        = row['observations']
        next_date           = row['next_date']

        # Cargar la plantilla
        doc = DocxTemplate(template_file)

        # Reemplazar los marcadores con los valores de la fila
        context = {'type': typecert, 'date': date, 'company': company, 'equipment': equipment, 'reference': reference, 'id_code': id_code, 'tag_code': tag_code, 'brand': brand, 'model': model, 'fabrication_date': fabrication_date, 'batch_number': batch_number, 'concept': concept, 'observations': observations, 'next_date': next_date}
        doc.render(context)

        # Guardar el nuevo documento en la carpeta de certificados
        output_file = os.path.join(output_folder, f'{company}_certificado_{equipment}_{current_date}.docx')
        doc.save(output_file)

    print("¡Certificados generados con éxito!")

# Archivos de entrada
csv_file = 'certificados.csv'
template_file = 'certificado_equipo.docx'

# Carpeta de salida
output_folder = 'certificados'

# Crear la carpeta de certificados si no existe
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Generar los certificados
generar_certificados(csv_file, template_file, output_folder)

# Ruta de la carpeta de certificados
certificados_folder = 'certificados'

# Obtener la lista de archivos .docx en la carpeta de certificados
docx_files = [file for file in os.listdir(certificados_folder) if file.endswith('.docx')]

# Convertir los archivos .docx a .pdf
for docx_file in docx_files:
    docx_path = os.path.join(certificados_folder, docx_file)
    pdf_path = os.path.join(certificados_folder, docx_file.replace('.docx', '.pdf'))
    convert(docx_path, pdf_path)

print("¡Certificados convertidos a PDF exitosamente!")

# Eliminar los archivos .docx
for docx_file in docx_files:
    docx_path = os.path.join(certificados_folder, docx_file)
    os.remove(docx_path)

print("¡Archivos .docx eliminados!")
