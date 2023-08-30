import requests
import xlrd
import pandas as pd
import os

def fetch_and_process_rig_data(url, save_path, sheet_name):
    # Fetch data
    dl = requests.get(url)
    
    # Save to disk
    with open(save_path, 'wb') as rig_data:
        rig_data.write(dl.content)
    try:

        # Read and clean up Excel data"
        rig_count = pd.read_excel(save_path, engine='pyxlsb', sheet_name=sheet_name)
        rig_df = pd.DataFrame(rig_count[5:])
        rig_df.columns = rig_df.iloc[0]
        rig_df = rig_df.iloc[1:].reset_index(drop=True)
        rig_df['Date'] = rig_df['Date'].apply(lambda x: xlrd.xldate_as_datetime(x, 0))
    
    finally:
        os.remove(save_path)

    return rig_df

if __name__ == '__main__':
    url = 'https://rigcount.bakerhughes.com/static-files/e916080f-555e-4e18-9969-46f62fbef9d6'
    save_path = 'F:/Downloads/north_america_rotary_rig_count_jan_2000_-_current.xlsb'
    sheet_name = 'US Oil & Gas Split'

    rig_df = fetch_and_process_rig_data(url, save_path, sheet_name)
    print(rig_df.head())
