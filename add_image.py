from docx import Document
from docx.shared import Cm

documento = Document()

# Adicionar uma imagem
documento.add_picture('notebook.jpg', width=Cm(5.25))

documento.save('demo.docx')