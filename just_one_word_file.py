from docx import Document
import os


# 'Función que cambia las diagonales de una dirección windows a una que es leída por python
def cambiando_diagonales(cadena):
    simbolo="\\"
    for i in range(len(simbolo)):
        cadena = cadena.replace(simbolo[i], '/')
        # va contando caracter por caracter y si encuentra espacios
        # los reemplaza por un espacio en blanco
        return cadena


carpeta = input(str("Copia y pega la ubicación de tus archivos a unir: "))
carpeta = cambiando_diagonales(carpeta)
# Cambiar al directorio 'Documentos'
os.chdir(carpeta)
print(carpeta)
documentos = os.listdir(carpeta)
documentos = sorted(documentos, key=lambda x: int(x[:2]) if x[:2].isdigit() else float('inf'))
nuevos_documentos = documentos[-10::]
# la lista con sólo los últimos 9 dígitos

documentos = documentos[:-10:]
# la lista sin los últimos 9 dígitos

documentos_ordenados = nuevos_documentos + documentos
# la concatenación ordenada

print(nuevos_documentos)
print(documentos)
print(documentos_ordenados)
# Lista de archivos que se van a unir
# documentos = ['descripciones.docx', 'descripciones - copia.docx']
# Aqui empieza la manipulación de todos los archivos
# Crea un nuevo documento
merged_document = Document()

# Itera sobre cada archivo y agrega su contenido al nuevo documento
for documento_solo in documentos_ordenados:
    # documento_solo... corresónde al archivo 1 word que creará
    document = Document(documento_solo)
    for element in document.element.body:
        merged_document.element.body.append(element)

# Guarda el documento unido en un archivo nuevo
merged_document.save('descripciones_finales.docx')
