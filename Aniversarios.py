import keyring
import pandas as pd
from datetime import datetime
from pptx import Presentation
import win32com.client
import os
from dotenv import load_dotenv
import pythoncom
import time
import sys

#Inicializar variáveis
load_dotenv(r"C:\Users\fabio.aguincha\Documents\Repos\Robot_Aniversarios\.env")
username = os.getenv("MY_USERNAME")
user_Email = os.getenv("MY_EMAIL")
user_Password = keyring.get_password("Robot_Aniversarios", user_Email)
from_Email = os.getenv("FROM_EMAIL")
temp_path=fr"C:\Users\{username}\Documents"


#Ler o ficheiro Excel, que deve estar na pasta mapeada do Onedrive
file_path = fr"C:\Users\{username}\VINCI Energies\GO-Data & AI - General\6 - PROJECTS\03 - DATA EXCELLENCE CENTER\RPAs - internos\HappyBirthday\Aniversários.xlsx"
df = pd.read_excel(file_path)

#Filtrar pelas pessoas que fazem anos hoje
today=datetime.today()
#today=datetime(2025,10,23)
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

    """
    def get_outlook_application(max_retries=3):
        for attempt in range(max_retries):
            try:
                # Initialize COM
                pythoncom.CoInitialize()
            
                # Try to get existing Outlook instance
                try:
                    outlook = win32com.client.GetActiveObject('Outlook.Application')
                    return outlook
                except pythoncom.com_error:
                    # If no existing instance, try to create one
                    outlook = win32com.client.Dispatch('Outlook.Application')
                    # Give Outlook time to initialize
                    time.sleep(2)
                    return outlook
                
            except Exception as e:
                print(f"Attempt {attempt + 1} failed: {e}")
                if attempt < max_retries - 1:
                    time.sleep(5)  # Wait before retry
                else:
                    print("Failed to connect to Outlook after all retries")
                    return None
    
        return None

    # Use it in your script
    outlook = get_outlook_application()
    if outlook is None:
        sys.exit(1)
    """
    
    # Abrir Outlook
    outlook = win32com.client.Dispatch('Outlook.Application')

    #Preparar email
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.BCC = "joana.morgado@axians.com; goncalo.vasconcelos@axians.com; joao.monteiro-simoes@axians.com; catarina.antunes-santos@axians.com"
    mail.SentOnBehalfOfName = from_Email
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
