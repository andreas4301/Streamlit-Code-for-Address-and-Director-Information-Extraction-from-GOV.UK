import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
from io import StringIO

# File upload
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

# Function to extract company information
def check_company_info(company_number):
    search_url = f"https://find-and-update.company-information.service.gov.uk/company/{company_number}"
    response = requests.get(search_url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        company_name_tag = soup.find("h1", class_="heading-xlarge")
        address_tag = soup.find("dd", class_="text data")
        
        company_name = company_name_tag.text.strip() if company_name_tag else None
        address = address_tag.text.strip() if address_tag else None
        
        # Handle foreign companies (those with 'FC' at the start of the company number)
        if company_number.startswith('FC'):
            return company_name, address
        
        # For non-foreign companies, extract the postcode (last two words of address)
        if address:
            postcode = " ".join(address.split()[-2:])  # Last two words for UK postcodes
            return company_name, postcode
        
    return None, None, None

# Function to get officer surnames
def get_company_officers(company_number):
    officers_url = f"https://find-and-update.company-information.service.gov.uk/company/{company_number}/officers"
    response = requests.get(officers_url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        officer_names = soup.find_all("span", id=lambda x: x and x.startswith("officer-name-"))
        surnames = []
        for tag in officer_names:
            full_name = tag.text.strip().split()
            if full_name:
                surnames.append(full_name[-1])  # Get surname (last part of the name)
        return " ".join(surnames)
    return ""

if uploaded_file is not None:
    # Read the Excel file
    df = pd.read_excel(uploaded_file)

    # Ensure the columns are as expected
    if len(df.columns) >= 3:
        df['Company Number'] = df.iloc[:, 0]  # Assuming first column is company number
        df['Company Name'] = df.iloc[:, 1]    # Assuming second column is company name
        df['URL'] = df.iloc[:, 2]             # Save the third column (URL)

        # Lists to store results
        company_names = []
        postcodes_or_addresses = []
        officer_surnames = []

        for index, row in df.iterrows():
            company_number = row['Company Number']
            
            # Fetch company info
            company_name, postcode_or_address = check_company_info(company_number)
            
            # Fetch officer surnames
            surnames = get_company_officers(company_number)

            # Save results
            company_names.append(company_name)
            postcodes_or_addresses.append(postcode_or_address)
            officer_surnames.append(surnames)

        # Create a new DataFrame with the results
        result_df = pd.DataFrame({
            'Company Number': df['Company Number'],
            'Company Name': company_names,
            'URL': df['URL'],  # Keep the URL from the original file
            'Postcode or Address': postcodes_or_addresses,
            'Officer Surnames': officer_surnames
        })

        # Convert to CSV
        csv_output = StringIO()
        result_df.to_csv(csv_output, index=False)
        csv_output.seek(0)

        # Provide the download button
        st.download_button(
            label="Download CSV",
            data=csv_output.getvalue(),
            file_name="company_info_output.csv",
            mime="text/csv"
        )
    else:
        st.write("Invalid file format. Make sure the first three columns are: Company Number, Name, and URL.")
