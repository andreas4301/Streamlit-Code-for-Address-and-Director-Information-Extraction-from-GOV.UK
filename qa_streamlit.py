import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd

# Function to fetch company info
def check_company_info(company_number):
    search_url = f"https://find-and-update.company-information.service.gov.uk/company/{company_number}"
    response = requests.get(search_url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        company_name_tag = soup.find("h1", class_="heading-xlarge")
        address_tag = soup.find("dd", class_="text data")
        if company_name_tag and address_tag:
            company_name = company_name_tag.text.strip()
            address = address_tag.text.strip()
            return company_name, address
    return None, None

# Function to fetch company officers
def get_company_officers(company_number):
    officers_url = f"https://find-and-update.company-information.service.gov.uk/company/{company_number}/officers"
    response = requests.get(officers_url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        officer_names = soup.find_all("span", id=lambda x: x and x.startswith("officer-name-"))
        officers = [tag.text.strip() for tag in officer_names]
        return officers
    return []

# Streamlit UI
st.title("Company Info Extractor")

uploaded_file = st.file_uploader("Upload an Excel file with Company Numbers", type=["xlsx"])

if uploaded_file:
    # Read the Excel file
    df = pd.read_excel(uploaded_file)

    # Create new columns to store the company name, address, and officers
    df['Company Name'] = ""
    df['Address'] = ""
    df['Officers'] = ""

    # Process each company number in the file
    for index, row in df.iterrows():
        company_number = str(row[0])  # Assuming company number is in the first column
        company_name, address = check_company_info(company_number)
        officers = get_company_officers(company_number)

        # Update the dataframe with fetched info
        df.at[index, 'Company Name'] = company_name
        df.at[index, 'Address'] = address
        df.at[index, 'Officers'] = ", ".join(officers)

    # Display the dataframe in Streamlit
    st.write(df)

    # Save the result to a new Excel file
    output_file = "company_info_output.xlsx"
    df.to_excel(output_file, index=False)
    st.download_button(label="Download Excel file", data=open(output_file, "rb"), file_name=output_file)
