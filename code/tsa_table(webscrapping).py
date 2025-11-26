import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta


# cwly = pp_end_date- timedelta(days=364)+timedelta(days=7)
# PWly_date=pp_end_date- timedelta(days=364)

def fetch_and_process_tsa_data(url, year):
    response = requests.get(url)

    if response.status_code == 200:
        soup = BeautifulSoup(response.content, "html.parser")
        table = soup.find("table")
        
        if table:
            # Extract table headers
            headers = [header.text.strip() for header in table.find_all("th")]

            
            if not headers:
                print(f"No headers found for {year}. Skipping.")
                return pd.DataFrame()
            
            # Extract table rows
            rows = []
            for row in table.find_all("tr")[1:]:  # Skip header row
                cells = [cell.text.strip() for cell in row.find_all("td")]
                rows.append(cells)
            
            # Create a DataFrame
            df_tsa_web = pd.DataFrame(rows, columns=headers)
            return df_tsa_web
        
        else:
            return pd.DataFrame()
    else:
        print(f"Failed to fetch the webpage for {year}. Status code: {response.status_code}")
        return pd.DataFrame()

def create_tsa_web_table ():
    # URLs for 2024 and 2025 data
    urls = {
        "2025": "https://www.tsa.gov/coronavirus/passenger-throughput",
        "2024": "https://www.tsa.gov/travel/passenger-volumes/2024"
    }

    # Fetch and process data for both years
    data_2024 = fetch_and_process_tsa_data(urls["2024"], "2024")
    data_2025 = fetch_and_process_tsa_data(urls["2025"], "2025")

    combined_data = pd.concat([data_2024, data_2025], ignore_index=True)

    combined_data['Date'] = pd.to_datetime(combined_data['Date'], errors='coerce')
    combined_data = combined_data.dropna(subset=['Date'])

    df_tsa_web = combined_data.sort_values(by='Date').reset_index(drop=True)
    df_tsa_web['Week End'] = df_tsa_web['Date'] + pd.to_timedelta((5 - df_tsa_web['Date'].dt.dayofweek) % 7, unit='d')

    df_tsa_web['Numbers']= df_tsa_web['Numbers'].str.replace(',', '').astype(int)

    return df_tsa_web


def create_df_tsa_table(tsa_data,df_tsa_web,format_percentage,end_date,pp_end_date,cwly_date,pwly_date):
    df_tsa_data=tsa_data.copy()
    df_tsa_data['depart_date']=pd.to_datetime(df_tsa_data['depart_date'])
    df_tsa_data['Week End'] = df_tsa_data['depart_date'] + pd.to_timedelta((5 - df_tsa_data['depart_date'].dt.dayofweek) % 7, unit='d')

    df_tsa = pd.DataFrame()

    #current year 
    a=df_tsa_data[(df_tsa_data['Week End'] == end_date.strftime("%d-%b-%Y"))]['passengers'].sum()
    b=(df_tsa_web[df_tsa_web['Week End'] == end_date.strftime("%d-%b-%Y")])['Numbers'].sum()
    df_tsa.loc['CY','priceline_passengers']= a
    df_tsa.loc['CY','TSA_passengers'] = b
    df_tsa.loc['CY','%'] = round(a*100/b,2)


    #last year
    a=df_tsa_data[(df_tsa_data['Week End'] == cwly_date.strftime("%d-%b-%Y"))]['passengers'].sum()
    b=(df_tsa_web[df_tsa_web['Week End'] == cwly_date.strftime("%d-%b-%Y")])['Numbers'].sum()
    df_tsa.loc['LY','priceline_passengers']= a
    df_tsa.loc['LY','TSA_passengers'] = b
    df_tsa.loc['LY','%'] = round(a*100/b,4)

    df_tsa.loc['CY','YoY'] = ((df_tsa.loc['CY','%']-df_tsa.loc['LY','%'])*100)


    #current year previous week
    a=(df_tsa_data[df_tsa_data['Week End'] == pp_end_date.strftime("%d-%b-%Y")])['passengers'].sum()
    b=(df_tsa_web[df_tsa_web['Week End'] ==pp_end_date.strftime("%d-%b-%Y")])['Numbers'].sum()
    df_tsa.loc['CY_PW','priceline_passengers']= a
    df_tsa.loc['CY_PW','TSA_passengers'] = b
    df_tsa.loc['CY_PW','%'] = round(a*100/b,2)

    #last year previous week
    a=(df_tsa_data[df_tsa_data['Week End'] == pwly_date.strftime("%d-%b-%Y")])['passengers'].sum()
    b=(df_tsa_web[df_tsa_web['Week End'] ==pwly_date.strftime("%d-%b-%Y")])['Numbers'].sum()
    df_tsa.loc['LY_PW','priceline_passengers']= a
    df_tsa.loc['LY_PW','TSA_passengers'] = b
    df_tsa.loc['LY_PW','%'] = round(a*100/b,4)

    df_tsa.loc['CY_PW','YoY'] = ((df_tsa.loc['CY_PW','%']-df_tsa.loc['LY_PW','%'])*100)
    df_tsa['%']=df_tsa['%'].apply(format_percentage)

    return df_tsa