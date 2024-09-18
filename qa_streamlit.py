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

# Function to split the long string into 8-character company numbers
def split_company_numbers(input_string):
    return [input_string[i:i+8] for i in range(0, len(input_string), 8)]

# Streamlit UI
st.title("Company Info Extractor")

# Input field for a big string of company numbers
company_numbers_string = st.text_input("Enter a string of company numbers (8 characters each, no spaces):")

if company_numbers_string:
    # Split the string into individual company numbers
    company_numbers = split_company_numbers(company_numbers_string)
    
    # Create a dataframe to store the results
    df = pd.DataFrame(company_numbers, columns=['Company Number'])
    df['Company Name'] = ""
    df['Address'] = ""
    df['Officers'] = ""

    # Process each company number
    for index, row in df.iterrows():
        company_number = row['Company Number']
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
