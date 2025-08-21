import keyring
import pandas as pd
from datetime import datetime
from pptx import Presentation
import win32com.client
from io import BytesIO
from PIL import Image
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import os

#Inicializar variáveis
username = os.environ["USERNAME"]
user_Email = "fabio.aguincha@axians.com"
user_Password = keyring.get_password("Robot_Aniversarios", user_Email)
server_SMTP = "smtp.office365.com"
port_SMTP= 587
from_Email = "data-ai.axianspt@axians.com"
temp_path=fr"C:\Users\{username}\Documents"


#Ler o ficheiro Excel, que deve estar na pasta mapeada do Onedrive
file_path = fr"C:\Users\{username}\VINCI Energies\GO-Data & AI - General\6 - PROJECTS\03 - DATA EXCELLENCE CENTER\RPAs - internos\HappyBirthday\Aniversários.xlsx"
df = pd.read_excel(file_path)

#Filtrar pelas pessoas que fazem anos hoje
today=datetime.today()
today=datetime(2025,9,22)       #data martelada
birthdays_today = df[
    (df["aniversário"].dt.day == today.day) &
    (df["aniversário"].dt.month == today.month)
]
if birthdays_today.empty:
    print("Ninguém faz anos hoje.")
    exit()

#Ler ficheiro PowerPoint, que deve estar na pasta mapeada do OneDrive
pptx_path = fr"C:\Users\{username}\VINCI Energies\GO-Data & AI - General\6 - PROJECTS\03 - DATA EXCELLENCE CENTER\RPAs - internos\HappyBirthday\parabens template.pptx"
ppt_app = win32com.client.Dispatch("PowerPoint.Application")
ppt_app.Visible = False
presentation = ppt_app.Presentations.Open(pptx_path, WithWindow=False)
slide = presentation.Slides(1)

for _, row in birthdays_today.iterrows():
    name = row['Nome']
    email = row['email']

    # Substituir [nome] pelo nome da pessoa
    for shape in slide.Shapes:
        if shape.HasTextFrame:
            shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text.replace("[nome]", name)
    
    # Exportar slide para imagem temporária
    temp_path_name = os.path.join(temp_path, f"temp_slide_{name}.png")
    slide.Export(temp_path_name, 'PNG')
    with open(temp_path_name, 'rb') as f:
        img_data = f.read()
    os.remove(temp_path_name)

    # Preparar email
    msg = MIMEMultipart("related")
    msg["Subject"] = f"Parabéns, {name}!"
    msg["From"] = from_Email
    msg["To"] = email

    html = '<html><body><img src="cid:slide_img"></body></html>'
    msg.attach(MIMEText(html, "html"))

    img = MIMEImage(img_data)
    img.add_header('Content-ID', '<slide_img>')
    msg.attach(img)

    # Enviar email
    with smtplib.SMTP(server_SMTP, port_SMTP) as server:
        server.starttls()
        server.login(user_Email, user_Password)
        server.sendmail(from_Email, email, msg.as_string())

    print(f"Sent birthday email to {name} ({email})")

# Close PowerPoint
presentation.Close()
ppt_app.Quit()