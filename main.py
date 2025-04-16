import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from openpyxl import load_workbook
from fpdf import FPDF
from datetime import date
import fitz  # PyMuPDF

def leer_Excel(filename):
    print("📘 Leyendo Excel...")
    wb = load_workbook(filename)
    print("Hojas:", wb.sheetnames)

    hojas_deseadas = ["OFICINA MARZO 2025", "JARDIN MARZO 2025"]
    for nombre_hoja in hojas_deseadas:
        if nombre_hoja in wb.sheetnames:
            ws = wb[nombre_hoja]
            Leer_Hoja(ws)
        else:
            print(f"❌ Hoja '{nombre_hoja}' no encontrada.")

def Leer_Hoja(ws):
    today = str(date.today())
    todayGroup = today.split("-")
    año = todayGroup[0]

    for row in range(2, ws.max_row + 1):
        birthday = ws['M' + str(row)].value
        if birthday:
            birthday = str(birthday).split(" ")[0]
            birthdayGroup = birthday.split("-")

            if birthdayGroup[1] == todayGroup[1] and birthdayGroup[2] == todayGroup[2]:
                print("🎉 Cumpleaños encontrado")

                nombre = ws['F' + str(row)].value.split(" ")[0]
                paterno = ws['D' + str(row)].value.strip()
                materno = ws['E' + str(row)].value.strip()
                correo = ws['T' + str(row)].value.strip()
                nombrepdf = "cumple.pdf"

                crear_pdf(nombre, paterno, materno, birthdayGroup[2], birthdayGroup[1], año, nombrepdf)
                nombreImagen = crear_imagen(nombrepdf)
                enviar_correo(nombreImagen, correo, nombre, paterno, materno)
                enviar_correo(nombreImagen, 'mferrerol@junji.cl', nombre, paterno, materno)

def crear_pdf(nombre, apellidoPaterno, apellidoMaterno, dia, mes, año, nombrePdf):
    print("📄 Creando PDF...")
    class PDF(FPDF):
        def header(self):
            self.image("image_1.png", keep_aspect_ratio=True, w=self.epw)
            self.ln(10)

        def footer(self):
            self.set_y(-45)
            self.image("image_2.png", keep_aspect_ratio=True, w=self.epw)

    meses = {
        "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
        "05": "Mayo", "06": "Junio", "07": "Julio", "08": "Agosto",
        "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
    }

    if len(dia) == 1: dia = "0" + dia
    if len(mes) == 1: mes = "0" + mes

    texto_principal00 = '¡Feliz Cumpleaños!'
    texto_principal1 = "En este día tan especial, deseamos que tengas un cumpleaños repleto de amor y felicidad. Que tus sueños se cumplan y estés siempre en compañía de la alegría, éxito y momentos llenos de sorpresas maravillosas."
    texto_principal2 = f"{nombre} {apellidoPaterno} {apellidoMaterno}"
    final1 = "¡Te deseamos lo mejor!"
    final2 = "Dirección Regional JUNJI Biobío"

    pdf = PDF('P', 'mm', 'A4')
    pdf.set_margin(0)
    pdf.add_page()
    pdf.set_text_color(9, 100, 175)
    pdf.set_font('Helvetica', 'B', 30)
    pdf.multi_cell(0, 20, texto_principal00, ln=True, align='C')
    pdf.set_font('Helvetica', 'BU', 40)
    pdf.multi_cell(0, 20, texto_principal2, ln=True, align='C')
    pdf.set_font('Helvetica', 'B', 20)
    pdf.multi_cell(0, 15, texto_principal1, ln=True, align='C')
    pdf.set_font('times', 'B', 30)
    pdf.set_text_color(1, 168, 158)
    pdf.ln()
    pdf.multi_cell(0, 20, final1, ln=True, align='C')
    pdf.multi_cell(0, 20, final2, ln=True, align='C')
    pdf.output(nombrePdf)

def crear_imagen(filename):
    print("🖼️ Creando imagen desde PDF...")
    doc = fitz.open(filename)
    page = doc.load_page(0)
    pix = page.get_pixmap()
    imagename = filename.replace(".pdf", ".png")
    pix.save(imagename)
    return imagename

def enviar_correo(filename, correo, nombre, paterno, materno):
    print(f"📧 Enviando correo a {correo}")
    remitente = '08junjibiobio@junji.cl'
    asunto = 'JUNJI te desea un feliz cumpleaños'
    cuerpo = f"""
    <html>
        <body>
        <img src='cid:image1'/>
        </body>
    </html>
    """

    username = remitente
    password = 'Tijunji2017'

    mensaje = MIMEMultipart()
    mensaje['From'] = remitente
    mensaje['To'] = correo
    mensaje['Subject'] = asunto

    with open(filename, "rb") as img_file:
        image = MIMEImage(img_file.read())
        image.add_header("Content-ID", "<image1>")
        mensaje.attach(image)

    mensaje.attach(MIMEText(cuerpo, 'html'))

    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    server.login(username, password)
    server.sendmail(remitente, correo, mensaje.as_string())
    server.quit()

def main():
    print("🚀 Iniciando proceso...")
    archivo_excel = "cumple_funcionario_correo.xlsx"
    leer_Excel(archivo_excel)

if __name__ == "__main__":
    main()
