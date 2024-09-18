import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment

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

# Input field for a space-separated string of company numbers
company_numbers_string = st.text_area("Enter a space-separated list of company numbers (8 characters each):")

if company_numbers_string:
    # Split the string into individual company numbers by spaces
    company_numbers = company_numbers_string.split()

    # Create a dataframe to store the results
    df = pd.DataFrame(company_numbers, columns=['Company Number'])
    df['Company Name'] = ""
    df['Address'] = ""

    max_officers = 0  # Track the maximum number of officers

    # Process each company number
    officer_data = []  # To hold officer data for all companies
    for index, company_number in enumerate(company_numbers):
        company_name, address = check_company_info(company_number)
        officers = get_company_officers(company_number)

        # Update max_officers if this company has more officers
        max_officers = max(max_officers, len(officers))

        # Store data for later use
        officer_data.append({
            'Company Number': company_number,
            'Company Name': company_name,
            'Address': address,
            'Officers': officers
        })

    # Add officer columns dynamically based on the max number of officers
    for i in range(1, max_officers + 1):
        df[f'Officer {i}'] = ""

    # Populate the dataframe with company info and officer data
    for index, data in enumerate(officer_data):
        df.at[index, 'Company Name'] = data['Company Name']
        df.at[index, 'Address'] = data['Address']
        for i, officer in enumerate(data['Officers']):
            df.at[index, f'Officer {i+1}'] = officer

    # Display the dataframe in Streamlit
    st.write(df)

    # Save the result to a new Excel file with appropriate column widths
    output_file = "company_info_output.xlsx"
    wb = Workbook()
    ws = wb.active

    # Write the dataframe to the Excel sheet
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # Auto-adjust column widths
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = length + 2

    # Save the file
    wb.save(output_file)

    # Provide download option
    st.download_button(label="Download Excel file", data=open(output_file, "rb"), file_name=output_file)
