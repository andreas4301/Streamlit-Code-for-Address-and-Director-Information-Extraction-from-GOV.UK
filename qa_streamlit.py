import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd

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
        
        if company_name_tag and address_tag:
            postcode = address_tag.text.strip().split()[-1]
            return company_name_tag.text.strip(), postcode
    return None, None

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
        df['URL'] = df.iloc[:, 2]             # Assuming third column is URL (not used in this case)

        # Lists to store results
        company_names = []
        postcodes = []
        officer_surnames = []

        for index, row in df.iterrows():
            company_number = row['Company Number']
            
            # Fetch company info
            company_name, postcode = check_company_info(company_number)
            
            # Fetch officer surnames
            surnames = get_company_officers(company_number)

            # Save results
            company_names.append(company_name)
            postcodes.append(postcode)
            officer_surnames.append(surnames)

        # Add results to DataFrame
        df['Company Name'] = company_names
        df['Postcode'] = postcodes
        df['Officer Surnames'] = officer_surnames

        # Save to a new Excel file
        output_file = "company_info_output.xlsx"
        df[['Company Number', 'Company Name', 'Postcode', 'Officer Surnames']].to_excel(output_file, index=False)

        st.write("Data extracted and saved successfully!")
        st.download_button(label="Download Excel file", data=output_file, file_name=output_file)

    else:
        st.write("Invalid file format. Make sure the first three columns are: Company Number, Name, and URL.")
