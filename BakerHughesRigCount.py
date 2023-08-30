import requests
import xlrd
import pandas as pd
import os
from io import BytesIO

def fetch_and_process_rig_data(url ,sheet_name):
    # Fetch data
    dl = requests.get(url)
    
    if dl.status_code == 200:
        # Read the Excel file into a Pandas DataFrame
        with BytesIO(dl.content) as bio:
            rig_count = pd.read_excel(bio, engine='pyxlsb', sheet_name='US Oil & Gas Split')

        rig_df = pd.DataFrame(rig_count[5:])
        rig_df.columns = rig_df.iloc[0]
        rig_df = rig_df.iloc[1:].reset_index(drop=True)
        rig_df['Date'] = rig_df['Date'].apply(lambda x: xlrd.xldate_as_datetime(x, 0))
    

    return rig_df

if __name__ == '__main__':
    url = 'https://rigcount.bakerhughes.com/static-files/e916080f-555e-4e18-9969-46f62fbef9d6'
    sheet_name = 'US Oil & Gas Split'

    rig_df = fetch_and_process_rig_data(url, sheet_name)
    print(rig_df.head())
