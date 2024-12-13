import streamlit as st
import pandas as pd
import os
import time
import locale
import calendar
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
import datetime
from collections import Counter
from PyPDF2 import PdfReader, PdfWriter
import io
from io import BytesIO
from reportlab.lib.utils import ImageReader

# Funciones
def dividir_texto(texto, width, max_width):
    palabras = texto.split()
    lineas = []
    linea_actual = palabras[0]

    for palabra in palabras[1:]:
        if c.stringWidth(linea_actual + " " + palabra, font_name, font_size) < max_width:
            linea_actual += " " + palabra
        else:
            lineas.append(linea_actual)
            linea_actual = palabra

    lineas.append(linea_actual)
    return lineas

def tarjetitas_empresa(auxiliar_y, elementos_a_escribir):

    # Definimos datos de la empresa
    id_empresa = f"ID: {elementos_a_escribir[0]}"
    nombre_empresa = elementos_a_escribir[1]
    provincia_empresa = elementos_a_escribir[2]
    productos_que_ofrece = elementos_a_escribir[3]
    productos_que_demanda = elementos_a_escribir[4]


    # Escribimos el ID de la empresa
    font_name = "Inter-Regular"
    font_size = 30
    c.setFillColorRGB(0, 0, 0)
    c.setFont(font_name, font_size)
    c.drawString(1150 , 1295 + auxiliar_y, id_empresa)

    # Escribimos el NOMBRE de la empresa
    font_name = "Inter-Regular"
    font_size = 37 
    c.setFillColorRGB(0, 0, 0)
    c.setFont(font_name, font_size)
    # Calcular tamaño del str
    text_width = c.stringWidth(nombre_empresa, font_name, font_size)
    if text_width > (1010 - 68 ):
        # Dividir el texto en líneas sin cortar palabras
        lineas = dividir_texto(nombre_empresa, text_width, 780)

        ii = 0   
        # Calcular la posición Y para las líneas
        for linea in lineas:
            c.drawString(68 , 1816 + auxiliar_y + ii , linea)
            ii -= 40     
    else:    
        c.drawString(68 , 1816 + auxiliar_y, nombre_empresa)

    # Escribimos el PROVINCIA de la empresa
    font_name = "Inter-Regular"
    font_size = 37
    c.setFillColorRGB(0, 0, 0)
    c.setFont(font_name, font_size)
    text_width = c.stringWidth(provincia_empresa, font_name, font_size)
    if text_width > (1350 - 1060 ):
        # Dividir el texto en líneas sin cortar palabras
        lineas = dividir_texto(provincia_empresa, text_width, 280)

        ii = 0   
        # Calcular la posición Y para las líneas
        for linea in lineas:
            c.drawString(1050 , 1816 + auxiliar_y + ii , linea)
            ii -= 40     
    else:    
        c.drawString(1050 , 1816 + auxiliar_y, provincia_empresa)

    # Escribimos los PRODUCTOS QUE OFRECE la empresa
    font_name = "Antonio-Medium"  # Nombre registrado de tu fuente personalizada
    font_size = 28  # Tamaño de la letra
    c.setFillColorRGB(1, 1, 1)
    c.setFont(font_name, font_size)
    text_width = c.stringWidth(productos_que_ofrece, font_name, font_size)
    if text_width > (1326 - 232 ):
        # Dividir el texto en líneas sin cortar palabras
        lineas = dividir_texto(productos_que_ofrece, text_width, 1520)

        ii = 0   
        # Calcular la posición Y para las líneas
        for linea in lineas:
            c.drawString(232 , 1717 + auxiliar_y + ii , linea)
            ii -= 40     
    else:    
        c.drawString(232 , 1717 + auxiliar_y , productos_que_ofrece)

    # Escribimos los PRODUCTOS QUE DEMANDA la empresa
    font_name = "Antonio-Medium"  # Nombre registrado de tu fuente personalizada
    font_size = 28  # Tamaño de la letra
    c.setFillColorRGB(1, 1, 1)
    c.setFont(font_name, font_size)
    text_width = c.stringWidth(productos_que_demanda, font_name, font_size)
    if text_width > (1326 - 232 ):
        # Dividir el texto en líneas sin cortar palabras
        lineas = dividir_texto(productos_que_demanda, text_width, 1520)

        ii = 0   
        # Calcular la posición Y para las líneas
        for linea in lineas:
            c.drawString(232 , 1500  + auxiliar_y + ii , linea)
            ii -= 40     
    else:    
        c.drawString(232 , 1500 + auxiliar_y , productos_que_demanda)


# Título del proyecto
st.title("Generador de Catálogo PDF")
st.write("---")
# Sección de carga de archivo Excel
st.header("Carga de Datos")
excel_file = st.file_uploader("Subir el archivo excel de la ronda de negocios", type=['xlsx', 'xls'])


data = None
if excel_file:
    try:
        df = pd.read_excel(excel_file,dtype=str)
        st.success("Archivo cargado con éxito")
        # st.dataframe(df.head())
    except Exception as e:
        st.error(f"Error al cargar el archivo: {e}")

# Sección de carga de imágenes
st.write("---")
st.header("Carga de Imágenes")
images = st.file_uploader("Subir el banner de la ronda de negocios correspondiente", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)     
if images:
    st.success(f"Se subió la imágen con éxito.")
    for image in images:
        img = Image.open(image)
        st.image(img, caption=image.name, use_container_width=True)
        # Convertir la imagen PIL a bytes
        img_buffer = BytesIO()
        img.save(img_buffer, format="PNG")  # Usa el formato adecuado
        img_buffer.seek(0)

# Sección de entrada de texto
st.write("---")
st.header("Carga de textos")
nombre_ronda = st.text_input("Nombre de la ronda de negocios (Aparecerá en el catálogo y en el pdf resultante)")
# espaciado_para_indice = st.text_input("Es el espaciado para el índice, por defecto 78:", "78")
# espaciado_para_indice = int(espaciado_para_indice)
st.write("---")

st.write("###### *NOTA: Los nombres de las columnas a usar deben de ser así:*")
col1, col2 = st.columns([1,1])
with col1:
    st.write("- Razón Social/Nombre de la empresa")
    st.write("- Productos que ofrece")
    st.write("- Productos que demanda")
with col2:    
    st.write("- Rubro de la empresa")
    st.write("- Provincia")
    st.write("- Id")

st.write("---")
# ------------------------------------------------------------------------------------------------------------------------------------------------
# Ingestamos excel
if st.button("Generar PDF"):
    try:
        # df = pd.read_excel("0 - Bases de Rondas/7 - Ronda Multisectorial de negocios Expo Bragado 2024 (respuestas).xlsx")
        # Manejar valores faltantes según el tipo de la columna
        # Convertir la columna a str para evitar conflictos
        # df['Teléfono de contacto (Código de área - Teléfono)'] = df['Teléfono de contacto (Código de área - Teléfono)'].astype(str)

        for col in df.columns:
            if df[col].dtype == 'object':  # Si la columna es de texto
                df[col].fillna("", inplace=True)
            else:  # Si la columna es numérica
                df[col].fillna(0, inplace=True)

        # Normalizamos: Aplicar str.strip() solo a las columnas de tipo string
        df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
        df["Razón Social/Nombre de la empresa"] = df["Razón Social/Nombre de la empresa"].str.title()
        df["Productos que ofrece"] = df["Productos que ofrece"].astype(str).str.capitalize()
        df["Productos que demanda"] = df["Productos que demanda"].astype(str).str.capitalize()


        # Definir la longitud máxima permitida
        longitud_maxima = 472
        # Recortar la cadena si sobrepasa la longitud máxima y añadir "..."
        df['Productos que ofrece'] = df['Productos que ofrece'].apply(lambda x: x[:longitud_maxima] + '...' if len(x) > longitud_maxima else x)
        df['Productos que demanda'] = df['Productos que demanda'].apply(lambda x: x[:longitud_maxima] + '...' if len(x) > longitud_maxima else x)

        # Tamaño de la hoja A4
        custom_page_size = (1414, 2000)  

        # Nombre del pdf
        pdf_output_path = nombre_ronda

        pdf_buffer = io.BytesIO()
        # Crear el objeto Canvas, especificando el archivo de salida y el tamaño de la página
        c = canvas.Canvas(pdf_buffer, pagesize= custom_page_size)

        # Registramos las FUENTES a usar
        font_path_antonio = "Fuentes/Antonio/Antonio-Bold.ttf"  
        pdfmetrics.registerFont(TTFont("Antonio-Bold", font_path_antonio))

        font_path_inter = "Fuentes/Inter/Inter-Light.ttf"  
        pdfmetrics.registerFont(TTFont("Inter-Light", font_path_inter))

        font_path_inter = "Fuentes/Inter/Inter-Regular.ttf"  
        pdfmetrics.registerFont(TTFont("Inter-Regular", font_path_inter))

        font_path_inter = "Fuentes/Inter/Inter-Bold.ttf"  
        pdfmetrics.registerFont(TTFont("Inter-Bold", font_path_inter))

        font_path = "Fuentes/Antonio/Antonio-Medium.ttf" 
        pdfmetrics.registerFont(TTFont("Antonio-Medium", font_path))

        # Hacemos la PORTADA
        width, height = custom_page_size
        

        img_reader = ImageReader(img_buffer)
        c.drawImage(img_reader, 0, 250, width  , height - 200)


        # LE ESCRIBIMOS
        font_name = "Inter-Bold"  # Nombre registrado de tu fuente personalizada
        font_size = 102  # Tamaño de la letra
        c.setFillColorRGB(0, 0, 0)
        c.setFont(font_name, font_size)
        # Calcular la posición X centrada
        x1 = 0  # Coordenada X izquierda
        x2 = 1414  # Coordenada X derecha
        text_width = c.stringWidth("Solicitud de entrevistas", font_name, font_size)
        centered_x = (x2 - x1 - text_width) / 2 + x1
        c.drawString(centered_x , 150 , "Solicitud de entrevistas")

        c.showPage()

        # Hacemos la primer hoja
        
        # Hacemos la primer hoja
        image = Image.open("imgs/2 - catálogo.png")
        width, height = custom_page_size
        c.drawImage("imgs/2 - catálogo.png", 0, 0, width, height)

        # LE ESCRIBIMOS nombre de la ronda
        font_name = "Inter-Regular"  # Nombre registrado de tu fuente personalizada
        font_size = 30  # Tamaño de la letra
        c.setFillColorRGB(0, 0, 0)
        c.setFont(font_name, font_size)
        c.drawString(28 , 28 , nombre_ronda)

        # Obtenemos el primer RUBRO (rubro actual)
        for indice, elemento in df.iterrows():
            rubro_actual  = elemento['Rubro de la empresa']
            break

        # cantidad_de_rubros = [rubro_actual]
        cantidad_de_rubros = []
        # Para la paginación
        paginación = 1

        # Para el conteo de empresas, al llegar a 3 reinicia.
        conteo_empresas = 0

        # Escribimos el rubro actual
        font_name = "Inter-Bold"  # Nombre registrado de tu fuente personalizada
        font_size = 52  # Tamaño de la letra
        c.setFillColorRGB(0, 0, 0)
        c.setFont(font_name, font_size)
        c.drawString(44 , 1920 , rubro_actual)

        # Escribimos el NÚMERO de la PÁGINA
        font_name = "Inter-Regular"  # Nombre registrado de tu fuente personalizada
        font_size = 32  # Tamaño de la letra
        c.setFillColorRGB(0, 0, 0)
        c.setFont(font_name, font_size)
        c.drawString(1350 , 28 , str(paginación))

        # Acá iterariamos en el df
        for indice, elemento in df.iterrows():
            # ID de la empresa
            id = elemento["Id"] # por el momento ponemos el cp como ID pa probar

            # Nombre de la empresa
            empresa = elemento["Razón Social/Nombre de la empresa"]
                    
            # Provincia
            provincia = elemento["Provincia"]

            # Rubro de la empresa
            rubro = elemento['Rubro de la empresa']
            cantidad_de_rubros.append(rubro)

            # Oferta y demanda
            productos_que_ofrece = elemento['Productos que ofrece']

            productos_que_demanda = elemento['Productos que demanda']
            
            elementos_a_escribir = [id, empresa, provincia, productos_que_ofrece, productos_que_demanda]
            
            # Si el rubro cambia o es la primera empresa, crear una nueva hoja
            if rubro != rubro_actual or conteo_empresas == 3:
                # Checkeo de cantidad de empresas pa poner parte blanca
                if conteo_empresas == 1:
                    image = Image.open("imgs/imágen blanca.png")
                    c.drawImage("imgs/imágen blanca.png",   10 , 70, 1414, 1200)
                elif conteo_empresas == 2:
                    image = Image.open("imgs/imágen blanca.png")
                    c.drawImage("imgs/imágen blanca.png",   10 , 70, 1414, 600)

                # Cambiamos de hoja
                c.showPage()

                # Creamos nueva hoja
                image = Image.open("imgs/2 - catálogo.png")
                width, height = custom_page_size
                c.drawImage("imgs/2 - catálogo.png", 0, 0, width, height)

                # LE ESCRIBIMOS nombre de la ronda
                font_name = "Inter-Regular"  # Nombre registrado de tu fuente personalizada
                font_size = 30  # Tamaño de la letra
                c.setFillColorRGB(0, 0, 0)
                c.setFont(font_name, font_size)
                c.drawString(28 , 28 , nombre_ronda)

                # Reiniciar variables
                conteo_empresas = 0 
                rubro_actual = rubro
                paginación += 1

                # Escribimos el rubro actual
                font_name = "Inter-Bold"  # Nombre registrado de tu fuente personalizada
                font_size = 52  # Tamaño de la letra
                c.setFillColorRGB(0, 0, 0)
                c.setFont(font_name, font_size)
                c.drawString(44 , 1920 , rubro_actual)

                # Escribimos el NÚMERO de la PÁGINA
                font_name = "Inter-Regular"  # Nombre registrado de tu fuente personalizada
                font_size = 32  # Tamaño de la letra
                c.setFillColorRGB(0, 0, 0)
                c.setFont(font_name, font_size)
                c.drawString(1350 , 28 , str(paginación))
                
            else : pass 
        
            # Auxiliar de Y para ir bajando las tarjetitas de las empresas
            if conteo_empresas == 0:
                auxiliar_y = 0
            elif conteo_empresas == 1:
                auxiliar_y = -598
            elif conteo_empresas == 2:
                auxiliar_y = -1196

            # Escribimos las tarjetitas
            tarjetitas_empresa(auxiliar_y, elementos_a_escribir)
        
            # Sumamos al contador de empresas
            conteo_empresas += 1

        # Una vez terminado hacemos el checkeo por fuera
        if conteo_empresas == 1:
            image = Image.open("imgs/imágen blanca.png")
            c.drawImage("imgs/imágen blanca.png",   10 , 70, 1414, 1200)
        elif conteo_empresas == 2:
            image = Image.open("imgs/imágen blanca.png")
            c.drawImage("imgs/imágen blanca.png",   10 , 70, 1414, 600)    

        c.showPage()

        # Hacemos el índice
        # Creamos nueva hoja
        image = Image.open("imgs/0 - índice.png")
        width, height = custom_page_size
        c.drawImage("imgs/0 - índice.png", 0, 0, width, height)

        # LE ESCRIBIMOS nombre de la ronda
        font_name = "Inter-Regular"  # Nombre registrado de tu fuente personalizada
        font_size = 30  # Tamaño de la letra
        c.setFillColorRGB(0, 0, 0)
        c.setFont(font_name, font_size)
        c.drawString(28 , 28 , nombre_ronda)

       # Contamos las cantidades de cada rubro
        contador_rubros = Counter(cantidad_de_rubros)

        # Diccionario para almacenar el índice de páginas por rubro
        indice = {}

        # Inicializar el número de página
        pagina_actual = 1

        # Número de empresas por página
        empresas_por_pagina = 3

        # Calcular las páginas para cada rubro
        for rubro, cantidad in contador_rubros.items():
            # Calcular cuántas páginas necesita este rubro
            paginas_necesarias = (cantidad // empresas_por_pagina) + (1 if cantidad % empresas_por_pagina > 0 else 0)
          
            # Asignar el rango de páginas para este rubro
            pagina_inicio = pagina_actual
            pagina_fin = pagina_inicio + paginas_necesarias - 1

            # Guardar el rango de páginas para el rubro
            indice[rubro] = (pagina_inicio, pagina_fin)

            # Actualizar el número de página para el siguiente rubro
            pagina_actual = pagina_fin + 1

        # Imprimir el índice para comprobar
        # for rubro, (pagina_inicio, pagina_fin) in indice.items():
        #     st.write(f"Rubro: {rubro}, Páginas: {pagina_inicio}-{pagina_fin}")


        i = 1600
        # Escribimos el NÚMERO de la PÁGINA
        font_name = "Inter-Regular"  # Nombre registrado de tu fuente personalizada
        font_size = 42  # Tamaño de la letra
        c.setFillColorRGB(0, 0, 0)
        c.setFont(font_name, font_size)

        max_width = 600  # Ancho máximo para los rubros antes de los puntos
        end_position = 1200  # Posición fija donde empiezan los números de página

        # Calculamos la distancia entre rubros dependiendo de la cantidad
        num_rubros = len(indice.keys())
        espacio_total = 1520  # Altura disponible en la página
        distancia_min = 50  # Distancia mínima
        distancia_max = 100  # Distancia máxima

        # Calculamos la distancia ideal para los rubros
        distancia = espacio_total // num_rubros

        # Aseguramos que la distancia esté dentro del rango permitido
        distancia = max(distancia_min, min(distancia, distancia_max))

        # Verificar si la distancia calculada hace que los rubros se salgan del espacio
        espacio_utilizado = distancia * (num_rubros - 1)  # El espacio ocupado por todos los rubros (sin el primero)

        if espacio_utilizado > espacio_total:
            # Ajustamos la distancia si el espacio utilizado supera el total disponible
            distancia = espacio_total // (num_rubros - 1)  # Calculamos una distancia que no exceda el total

            # Aseguramos que la distancia esté dentro de los límites
            distancia = max(distancia_min, min(distancia, distancia_max))

        # Imprimimos la cantidad de rubros y la distancia calculada
        # st.write(f"Número de rubros: {num_rubros}")
        # st.write(f"Distancia entre rubros: {distancia}")

        # Generamos el índice
        for rubro, (pagina_inicio, pagina_fin) in indice.items():
            # Calcular la longitud del rubro
            rubro_width = c.stringWidth(rubro, font_name, font_size)
            
            # Calcular cuántos puntos se necesitan para llenar el espacio
            dots = '.' * int((end_position - (60 + rubro_width)) / c.stringWidth('.', font_name, font_size))
            
            # Crear la línea completa con puntos y número de página
            if pagina_inicio == pagina_fin:
                linea = f"{rubro} {dots} {pagina_inicio}"
            else:    
                linea = f"{rubro} {dots} {pagina_inicio}-{pagina_fin}"
            
            # Escribir la línea en el PDF
            c.drawString(60, i, linea)
            i -= distancia  # Reducimos la posición 'i' según la distancia calculada




        c.save()

        # Leer el PDF generado en el buffer
        pdf_buffer.seek(0)
        reader = PdfReader(pdf_buffer)

        # Crear un nuevo PDF reorganizado
        writer = PdfWriter()
        pages = list(reader.pages)

        # Reorganizar las páginas: mover la última página a la segunda posición
        last_page = pages.pop(-1)
        pages.insert(1, last_page)

        for page in pages:
            writer.add_page(page)

        # Guardar el PDF reorganizado en otro buffer
        final_pdf_buffer = io.BytesIO()
        writer.write(final_pdf_buffer)
        final_pdf_buffer.seek(0)

        # Permitir descarga del archivo reorganizado
        st.download_button(
            label="Descargar Catálogo PDF",
            data=final_pdf_buffer,
            file_name=f"{pdf_output_path}.pdf",
            mime="application/pdf"
        )

        st.success("¡Catálogo generado con éxito!")
    except Exception as e:
        st.error(f"Error en la generación: || {e} ||")
