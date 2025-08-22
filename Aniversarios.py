import keyring
import pandas as pd
from datetime import datetime
from pptx import Presentation
import win32com.client
import os

#Inicializar variáveis
username = os.environ["USERNAME"]
user_Email = os.environ["MY_EMAIL"]
user_Password = keyring.get_password("Robot_Aniversarios", user_Email)
from_Email = os.environ["FROM_EMAIL"]
temp_path=fr"C:\Users\{username}\Documents"


#Ler o ficheiro Excel, que deve estar na pasta mapeada do Onedrive
file_path = fr"C:\Users\{username}\VINCI Energies\GO-Data & AI - General\6 - PROJECTS\03 - DATA EXCELLENCE CENTER\RPAs - internos\HappyBirthday\Aniversários.xlsx"
df = pd.read_excel(file_path)

#Filtrar pelas pessoas que fazem anos hoje
today=datetime.today()
birthdays_today = df[
    (df["aniversário"].dt.day == today.day) &
    (df["aniversário"].dt.month == today.month)
]
if birthdays_today.empty:
    print("Ninguém faz anos hoje.")
    exit()

#Ler ficheiro PowerPoint, que deve estar na pasta mapeada do OneDrive
pptx_path = fr"C:\Users\{username}\VINCI Energies\GO-Data & AI - General\6 - PROJECTS\03 - DATA EXCELLENCE CENTER\RPAs - internos\HappyBirthday\parabens template.pptx"

#Para cada pessoa:
for _, row in birthdays_today.iterrows():
    name = row['Nome'].split()[0]
    email = row['email']

    #Abrir PowerPoint
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    presentation = ppt_app.Presentations.Open(pptx_path, WithWindow=False)
    slide = presentation.Slides(1)

    # Substituir [nome] pelo nome da pessoa
    for shape in slide.Shapes:
        if shape.HasTextFrame:
            if shape.TextFrame.HasText:
                shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text.replace("[nome]", name)
    
    
    # Exportar slide para imagem temporária
    temp_path_name = os.path.join(temp_path, f"temp_slide_{name}.png")
    slide.Export(temp_path_name, 'PNG')


    # Abrir Outlook
    outlook = win32com.client.Dispatch('Outlook.Application')

    #Preparar email
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = f"Parabéns, {name}!"
    mail.HTMLBody = f'<html><body><img src="cid:slideimage"></body></html>'
    attachment = mail.Attachments.Add(temp_path_name)
    attachment.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "slideimage"
    )

    #Enviar
    mail.Send()
    print(f"Sent birthday email to {name} ({email})")

    #Limpar ficheiros temporários e fechar PowerPoint
    os.remove(temp_path_name)
    presentation.Close()
    ppt_app.Quit()
