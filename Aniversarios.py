# %%
import os
import pandas as pd
from datetime import datetime


#Ler o ficheiro, que deve estar na pasta mapeada da Onedrive
username = os.environ["USERNAME"]
file_path = fr"C:\Users\{username}\VINCI Energies\GO-Data & AI - General\6 - PROJECTS\03 - DATA EXCELLENCE CENTER\RPAs - internos\HappyBirthday\Aniversários.xlsx"
df = pd.read_excel(file_path)

#Filtrar pelas pessoas que fazem anos hoje
today=datetime.today()
today=datetime(2025,9,22)
birthdays_today = df[
    (df["aniversário"].dt.day == today.day) &
    (df["aniversário"].dt.month == today.month)
]

# Show first few rows
print(birthdays_today)
# %%
