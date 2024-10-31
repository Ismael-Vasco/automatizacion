import os
import requests
import pandas as pd
from datetime import datetime
import schedule
import time

# Configura tus credenciales y el dataset
CLIENT_ID = 'TU_CLIENT_ID'
CLIENT_SECRET = 'TU_CLIENT_SECRET'
TENANT_ID = 'TU_TENANT_ID'
GROUP_ID = 'TU_GROUP_ID'
DATASET_ID = 'TU_DATASET_ID'
POWER_BI_API_URL = 'https://api.powerbi.com/v1.0/myorg/groups/{}/datasets/{}/tables/{}/rows'

def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        'grant_type': 'client_credentials',
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'scope': 'https://analysis.windows.net/powerbi/api/.default'
    }
    response = requests.post(url, data=data)
    response.raise_for_status()
    return response.json()['access_token']

def upload_files_to_powerbi(folder_path):
    token = get_access_token()
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}

    for filename in os.listdir(folder_path):
        if filename.endswith('.csv'):  # Cambia la extensión si es necesario
            file_path = os.path.join(folder_path, filename)
            df = pd.read_csv(file_path)
            rows = df.to_dict(orient='records')

            data = {'rows': rows}
            response = requests.post(POWER_BI_API_URL.format(GROUP_ID, DATASET_ID, 'NombreDeLaTabla'), headers=headers, json=data)
            response.raise_for_status()
            print(f'Subido: {filename}')

def job():
    folder_path = 'ruta/a/tu/carpeta'
    upload_files_to_powerbi(folder_path)

# Programar el trabajo cada mes
schedule.every().month.at("01:00").do(job)  # Cambia la hora si es necesario

while True:
    schedule.run_pending()
    time.sleep(1)
