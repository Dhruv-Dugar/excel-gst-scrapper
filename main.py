import pandas as pd
import requests
from bs4 import BeautifulSoup
# optional dependencies, but internal functions require them
import openpyxl
import lxml
import html5lib
import streamlit as st

# importing the data from the excel file
file = st.file_uploader("Upload file", type=["xlsx"])
if file is not None:
    df = pd.read_excel(file)
    df.describe()

    # making new dataframes for the required data
    required_data = pd.DataFrame()
    required_data = df[['GSTIN/UIN of Recipient', 'Reciever Name']].drop_duplicates()
    required_data['Count of Invoices'] = required_data['Reciever Name'].map(df['Reciever Name'].value_counts())
    required_data['Total Invoice Value'] = required_data['Reciever Name'].map(df.groupby('Reciever Name')['Invoice value'].sum())

    # defining the headers for making API calls
    headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:134.0) Gecko/20100101 Firefox/134.0',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-GB,en;q=0.5',
    'Accept-Encoding': 'gzip, deflate, br, zstd',
    'Referer': 'https://www.knowyourgst.com/gst-number-search/dugar-polychem-07ADKPD7356H1ZF/',
    'Content-Type': 'application/x-www-form-urlencoded',
    'Origin': 'https://www.knowyourgst.com',
    'DNT': '1',
    'Sec-GPC': '1',
    'Connection': 'keep-alive',
    'Cookie': 'csrftoken=3OrQeqBtkXwnNXNPYXipQEvKAWzJHc3DdffIyimcROKpoN1zK7HrfHepINF2EYdy; csrftoken=3OrQeqBtkXwnNXNPYXipQEvKAWzJHc3DdffIyimcROKpoN1zK7HrfHepINF2EYdy',
    'Upgrade-Insecure-Requests': '1',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-User': '?1',
    'Priority': 'u=0, i',
    'TE': 'trailers'
    }

    def getAddress(receiver_name: str, gstin: str):
        receiver_name = receiver_name.replace(' ', '-')
        receiver_name = receiver_name.replace('M/S', 'ms')
        url = f'https://www.knowyourgst.com/gst-number-search/{receiver_name}-{gstin}/'
        response = requests.get(url, headers=headers)
        soup = BeautifulSoup(response.text, 'html.parser')
        table = soup.find('table')
        table_data = str(table)
        print(table_data)
        df2 = pd.read_html(table_data)[0]
        address = df2["Details"][2]
        return address

    with st.spinner("Fetching data from the internet..."):    
        # enumerate through the data and get the address
        for index, row in required_data.iterrows():
            try:
                address = getAddress(row['Reciever Name'], row['GSTIN/UIN of Recipient'])
                required_data.loc[index, 'Address'] = address
                print(address)
            except Exception as e:
                print(f"Failed to get address for {row['Reciever Name']}: {e}")
                required_data.loc[index, 'Address'] = "failure"
                address = ""
            
    # export the required data to a new excel file in descending order of total invoice value
    required_data = required_data.sort_values(by='Total Invoice Value', ascending=False)
    required_data.to_excel("required_data.xlsx")

    # display the required data
    st.write(required_data)
    with open('required_data.xlsx', 'rb') as f:
        st.download_button('Download final file', f)