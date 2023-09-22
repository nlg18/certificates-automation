from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib.pagesizes import A4, landscape

from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.units import mm

import pandas as pd
import PyPDF2
import smtplib

import openpyxl
from PIL import Image, ImageDraw, ImageFont
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders

import os

# import time
# time.sleep(10)

solicit = pd.read_excel('./solicit.xlsx')

solicit = pd.DataFrame(data=solicit, columns=['Nome completo', 'E-mail'])
solicit.rename(columns={'Nome completo': 'Nome', 'E-mail': 'Email'}, inplace=True)

for i in range(1):
    d = canvas.Canvas("ementa.pdf")
    d.setPageSize(landscape(A4))
    d.drawInlineImage("ementa.jpg", 0, 0, width=840, height=600)
    style = getSampleStyleSheet()["Normal"]
    style.fontSize = 20
    style.fontName = "Helvetica-Bold"
    style.textColor = colors.orange
    p = Paragraph(" ", style)
    text_width = 655
    text_height = 680
    p.wrap(text_width, text_height)
    p.drawOn(d, 178, 380)
    d.showPage()
    d.save()


def write_centered_text(canvas, text, y):
    width, height = A4
    text_width = canvas.stringWidth(text)
    x = (width - text_width) / 2
    canvas.drawString(x, y, text)


for index, row in solicit.iterrows():
    email = row['Email']
    name = row['Nome']

    ######## criação certificado (passo 1 de 2)

    c = canvas.Canvas("certif_" + str(index + 1) + "20230629.pdf")
    c.setPageSize(landscape(A4))
    c.drawInlineImage("certificado.jpg", 0, 0, width=840, height=600)

    p = Paragraph(name, style)

    text_width = 855
    text_height = 680

    p.wrap(text_width, text_height)

    cent_number = 490 - 5 * len(name)
    if cent_number < 178:
        cent_number = 178
    if cent_number > 450:
        cent_number = 400

    p.drawOn(c, cent_number, 380)

    y = 380
    # write_centered_text(c, name, y)

    c.showPage()
    c.save()

    ######## envio do certificado por email (passo 2 de 2)

    msg = MIMEMultipart()
    msg['From'] = 'pro.nicolaslg@gmail.com'
    # msg['From'] = 'renatacesjf@ufjf.br'

    # msg['To'] = 'treina.critt@ufjf.br'
    msg['To'] = 'pro.nicolaslg@gmail.com'
    # msg['To'] = email

    msg['Subject'] = 'Certificado de Participação'

    espaco_indice = name.index(" ")
    corpo = f"""
    Olá, {name[0:espaco_indice]}!  

    Segue em anexo o seu certificado de participação.

    Fique atento ao @crittufjf no Instagram para futuros eventos!

    Atenciosamente,
    CRITT, UFJF.
    """

    msg.attach(MIMEText(corpo, 'plain'))

    filename = "certif_" + str(index + 1) + "20230629.pdf"
    attachment = open(filename, 'rb')
    p = MIMEBase('application', 'octet-stream')
    p.set_payload((attachment).read())
    encoders.encode_base64(p)
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    # msg.attach(p)

    ementa_filename = "ementa.pdf"
    ementa_attachment = open(ementa_filename, 'rb')
    ementa_part = MIMEBase('application', 'octet-stream')
    ementa_part.set_payload((ementa_attachment).read())
    encoders.encode_base64(ementa_part)
    ementa_part.add_header('Content-Disposition', "attachment; filename= %s" % ementa_filename)
    # msg.attach(ementa_part)

    pdf_combinado = PyPDF2.PdfMerger()
    pdf_combinado.append(filename)
    pdf_combinado.append(ementa_filename)
    # pdf_combinado.write("./certif/certificados.pdf")
    pdf_combinado.write("certificado_" + str(index + 1) + "20230629.pdf")

    filename_combinado = "certificado_" + str(index + 1) + "20230629.pdf"
    attachment_combinado = open(filename_combinado, 'rb')

    p_combinado = MIMEBase('application', 'octet-stream')
    p_combinado.set_payload((attachment_combinado).read())
    encoders.encode_base64(p_combinado)
    p_combinado.add_header('Content-Disposition', "attachment; filename= %s" % filename_combinado)
    msg.attach(p_combinado)

    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.starttls()
    # smtpObj.login('renatacesjf@ufjf.br', 'builjgiyhemnjirv')
    # todo
    # senha = manage your google account >security>How you sign in to Google>verificação em duas etapas > senhas de app
    smtpObj.login('pro.nicolaslg@gmail.com', senha)
    smtpObj.send_message(msg)
    smtpObj.quit()

    # os.remove(filename)
    # os.remove(ementa_filename)
    c = None
    d = None
    attachment.close()
    attachment = None
    ementa_attachment.close()
    ementa_attachment = None
    attachment_combinado.close()
    attachment_combinado = None
    pdf_combinado.close()
    pdf_combinado = None