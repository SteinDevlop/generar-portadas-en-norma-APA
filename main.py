from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.shared import Inches

def agregar_portada(titulo,estudiante,universidad,facultad,materia,docente,fecha):
    document = Document()
    # Tipo de letra y tamaño
    style = document.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # Interlineado
    paragraph_format = style.paragraph_format
    paragraph_format.line_spacing = 2

    # Margen
    sections = document.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
    title = document.add_paragraph()
    title_run = title.add_run(titulo)
    title_run.bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.add_paragraph(" ")
    for estudiante_singular in estudiante:
        document.add_paragraph(estudiante_singular[0]+".")
        document.add_paragraph(estudiante_singular[1]+".")
    document.add_paragraph(" ")
    document.add_paragraph(universidad+".")
    document.add_paragraph("Facultad de "+facultad+".")
    document.add_paragraph(materia+".")
    document.add_paragraph(docente+".")
    document.add_paragraph(fecha+".")
    for paragraph in document.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.save(f'docx/{titulo}.docx')
def capitalizar_mayores_tres(txt):
    resultado = txt.split()
    for i in range(len(resultado)):
        if len(resultado[i]) > 3:
            resultado[i] = resultado[i].capitalize()
    return " ".join(resultado)
def menu():
    titulo=capitalizar_mayores_tres(str(input("Titulo: ")))
    n_est = int(input("Numero de integrantes: "))
    estudiante=[]
    for i in range (1,n_est+1):
        nombre_est=capitalizar_mayores_tres(str(input("Nombre del estudiante: ")))
        codigo_est=capitalizar_mayores_tres(str(input("Codigo del estudiante: ")))
        estudiante.append([nombre_est,codigo_est])
    estudiante=sorted(estudiante, key=lambda x: x[0])
    print(estudiante)
    universidad=capitalizar_mayores_tres("universidad tecnologica de bolivar")
    
    facultades=["Ingenieria","Ciencias Basicas","Ciencias Sociales y Humanidades","Arquitectura","Derecho"]
    print("""Facultades:
Ingenieria - 1
Ciencias Basicas - 2
Ciencias Sociales y Humanidades - 3
Arquitectura - 4
Derecho - 5""")
    facultad=int(input("Indique la facultad: "))
    facultad=facultades[facultad]
    materia=capitalizar_mayores_tres(str(input("Materia: ")))
    docente=capitalizar_mayores_tres(str(input("Docente: ")))
    fecha=input("Dia de entrega: ")+" de "+input("Mes de entrega: ")+" de "+input("Año: ")
    agregar_portada(titulo,estudiante,universidad,facultad,materia,docente,fecha)

if __name__ == '__main__':
    menu()
