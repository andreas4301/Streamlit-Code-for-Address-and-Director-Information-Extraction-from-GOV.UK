import streamlit as st
import requests
from bs4 import BeautifulSoup

# Input fields
company_number = st.text_input("Enter Company Number")

if company_number:
    # Functions to fetch company info
    def check_company_info(company_number):
        search_url = f"https://find-and-update.company-information.service.gov.uk/company/{company_number}"
        response = requests.get(search_url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            company_name_tag = soup.find("h1", class_="heading-xlarge")
            address_tag = soup.find("dd", class_="text data")
            st.write(f"Company Name: {company_name_tag.text.strip()}")
            st.write(f"Registered Office Address: {address_tag.text.strip()}")
        else:
            st.write("Failed to fetch company data")

    def get_company_officers(company_number):
        officers_url = f"https://find-and-update.company-information.service.gov.uk/company/{company_number}/officers"
        response = requests.get(officers_url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            officer_names = soup.find_all("span", id=lambda x: x and x.startswith("officer-name-"))
            st.write("Officers:")
            for tag in officer_names:
                st.write(tag.text.strip())
        else:
            st.write("Failed to fetch officer data")

    check_company_info(company_number)
    get_company_officers(company_number)
