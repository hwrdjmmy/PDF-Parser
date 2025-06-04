# PDF-Parser
This was used for the DC Home Health project for extracting responses on a PDF Form

import os
import fitz  # PyMuPDF
from openpyxl import Workbook

# Define the folder containing PDF files
folder_path = r"C:\Users\James Howard\DHCF_Survey"  # Use raw string (r) to handle Windows paths

# Dictionary to store data for each file
file_data = {}
all_field_names = set()  # To collect all unique field names

# Loop through all files in the folder
for filename in os.listdir(folder_path):
    # Check if the file is a PDF
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(folder_path, filename)
        print(f"Processing file: {filename}")
        
        # Open the PDF
        pdf = fitz.open(pdf_path)
        
        # Initialize a dictionary to hold form field data for this file
        file_data[filename] = {}

        # Loop through the pages of the PDF
        for page_num in range(len(pdf)):
            page = pdf.load_page(page_num)
           
            # Get all form fields (widgets) on the page
            form_fields = page.widgets()
           
            # Loop through each form field (widget)
            for field in form_fields:
                field_name = field.field_name  # Name of the form field
                field_value = field.field_value  # Value entered in the form field

                # Store the field value keyed by field name
                file_data[filename][field_name] = field_value
                all_field_names.add(field_name)  # Collect unique field names
        
        # Close the PDF
        pdf.close()

# Prioritize specific field names like Q1 and Q2
priority_fields = ["Q1","Q2","Q3.1.0","Q3.1.1","Q3.1.2","Q3.1.3","Q3.1.4","Q3.1.5","Q3.1.6","Q3.1.7","Q3.1.8","Q3.1.9","Q3.1.10",
                   "Q3.1.11","Q3.1.12","Q3.1.13","Q3.1.14","Q3.2.0","Q3.2.1","Q3.2.2","Q3.2.3","Q3.2.4","Q3.2.5","Q3.2.6","Q3.2.7",
                   "Q3.2.8","Q3.2.9","Q3.2.10","Q3.2.11","Q3.2.12","Q3.2.13","Q3.2.14","Q4.1","Q4.2","Q4.3","Q4.4","Q4.5",
                   "Q4.6","Q4.7","Q5.1.1","Q5.1.2","Q5.1.3","Q5.1.4","Q5.1.5","Q5.1.6","Q5.1.7","Q5.1.8","Q5.1.9","Q5.1.10",
                   "Q5.1.11","Q5.2.1","Q5.2.2","Q5.2.3","Q5.2.4","Q5.2.5","Q5.2.6","Q5.2.7","Q5.2.8","Q5.2.9","Q5.2.10","Q5.2.11",

                   "Q6.0.0","Q6.0.1","Q6.0.2","Q6.0.3",
                   "Q6.1.0","Q6.1.1","Q6.1.2","Q6.1.3",
                   "Q6.2.0","Q6.2.1","Q6.2.2","Q6.2.3",
                   "Q6.3.0","Q6.3.1","Q6.3.2","Q6.3.3",
                   "Q6.4.0","Q6.4.1","Q6.4.2","Q6.4.3",
                   "Q6.5.0","Q6.5.1","Q6.5.2","Q6.5.3",
                   "Q6.6.0","Q6.6.1","Q6.6.2","Q6.6.3",
                   "Q6.7.0","Q6.7.1","Q6.7.2","Q6.7.3",
                   "Q6.8.0","Q6.8.1","Q6.8.2","Q6.8.3",
                   "Q6.9.0","Q6.9.1","Q6.9.2","Q6.9.3",
                   "Q6.10.0","Q6.10.1","Q6.10.2","Q6.10.3",

                   "Q7.0.0","Q7.0.1","Q7.0.2","Q7.0.3",
                   "Q7.1.0","Q7.1.1","Q7.1.2","Q7.1.3",
                   "Q7.2.0","Q7.2.1","Q7.2.2","Q7.2.3",
                   "Q7.3.0","Q7.3.1","Q7.3.2","Q7.3.3",
                   "Q7.4.0","Q7.4.1","Q7.4.2","Q7.4.3",
                   "Q7.5.0","Q7.5.1","Q7.5.2","Q7.5.3",
                   "Q7.6.0","Q7.6.1","Q7.6.2","Q7.6.3",
                   "Q7.7.0","Q7.7.1","Q7.7.2","Q7.7.3",
                   "Q7.8.0","Q7.8.1","Q7.8.2","Q7.8.3",
                   "Q7.9.0","Q7.9.1","Q7.9.2","Q7.9.3",
                   "Q7.10.0","Q7.10.1","Q7.10.2","Q7.10.3",

                   "Q7.0.0.0.0","Q7.0.0.0.1","Q7.0.0.0.2","Q7.0.0.0.3",
                   "Q7.0.0.1.0","Q7.0.0.1.1","Q7.0.0.1.2","Q7.0.0.1.3",
                   "Q7.0.0.2.0","Q7.0.0.2.1","Q7.0.0.2.2","Q7.0.0.2.3",
                   "Q7.0.0.3.0","Q7.0.0.3.1","Q7.0.0.3.2","Q7.0.0.3.3",
                   "Q7.0.0.4.0","Q7.0.0.4.1","Q7.0.0.4.2","Q7.0.0.4.3",
                   "Q7.0.0.5.0","Q7.0.0.5.1","Q7.0.0.5.2","Q7.0.0.5.3",
                   "Q7.0.0.6.0","Q7.0.0.6.1","Q7.0.0.6.2","Q7.0.0.6.3",
                   "Q7.0.0.7.0","Q7.0.0.7.1","Q7.0.0.7.2","Q7.0.0.7.3",
                   "Q7.0.0.8.0","Q7.0.0.8.1","Q7.0.0.8.2","Q7.0.0.8.3",
                   "Q7.0.0.9.0","Q7.0.0.9.1","Q7.0.0.9.2","Q7.0.0.9.3",
                   "Q7.0.0.10.0","Q7.0.0.10.1","Q7.0.0.10.2","Q7.0.0.10.3",

                   "Q8","Q9","Q10","Q11","Q12","Q13","Q14","Q15","Q15a","Q15b","Q16.0","Q16.1","Q16.2",
                   "Q16.3","Q17","Q18.0","Q18.1","Q18.2","Q19.0","Q19.1","Q19.2","Q20","Q21","Q22","Q23","Q24","Q25","Q26","Q27",
                   "Q28","Q29","Q30","Q31","Q32","Q33","Q34","Q34a","Q35","Q36","Q37.0.0","Q37.0.1","Q37.1.0","Q37.1.1","Q37.2.0",
                   "Q37.2.1","Q38","Q39","Q39a","Q40","Q41","Q42","Q42a","Q43","Q44","Q45","Q46","Q46a","Q47","Q48","Q49","Q50",
                   "Q50a","Q51","Q52","Q52a.0.0","Q52a.0.1","Q52a.1.0","Q52a.1.1","Q52a.2.0","Q52a.2.1","Q52b","Q53","Q54","Q55",
                   "Q55a","Q55a.0.0","Q55a.0.1","Q55a.1.0","Q55a.1.1","Q55a.2.0","Q55a.2.1","Q55b.0.0","Q55b.0.1","Q55b.1.0","Q55b.1.1",
                   "Q55b.2.0","Q55b.2.1","Q56","Q57","Q58","Q58a","Q58b","Q58c","Q58d","Q58e","Q59.1","Q59.2","Q59.3","Q60","Q61",
                   "Q61a","Q61b","Q61c","Q62","Q63","Q64","Q65","Q65a","Q65b","Q66","Q67","Q68","Q69","Q70","Q71","Q72","Q72a",
                   "Q73","Q74","Q75","Q75a","Q76","Q77","Q77a","Q78.1","Q78.2","Q78.3","Q78a","Q79.1","Q79.2","Q79.3","Q79a",
                   "Q80","Q81","Q81a.1","Q81a.2","Q81a.3","Q81b.1","Q81b.2","Q81b.3","Q82","Q83","Q84","Q85","Q86","Q86a","Q87",
                   "Q87a","Q87b","Q88","Q89","Q90","Q91","Q92","Q93.0.0","Q93.0.1","Q93.1.0","Q93.1.1","Q93.2.0","Q93.2.1",
                   "Q93.3.0","Q93.3.1","Q94"]

reordered_field_names = priority_fields + sorted(
    field for field in all_field_names if field not in priority_fields
)

# Create a new Excel workbook and select the active worksheet
workbook = Workbook()
sheet = workbook.active

# Write header row: "File Name", "HH_Agency" followed by reordered field names
header = ["File Name", "HH_Agency"] + reordered_field_names
sheet.append(header)

# Write data for each file
for filename, fields in file_data.items():
    row = [filename]  # Start row with the file name
    
    # Add the HH_Agency column based on conditional logic
    if filename == "2024_DHCF_HHRSD_ABA HOME HEALTH CARE_Fillable.pdf": 
        hh_agency = "ABA"
    elif filename == "2024_DHCF_HHRSD_ProviderSurvey_Fillable_v6_20241108_Kristi Chittum_Americare In Home Nursing_Updated.pdf": 
        hh_agency = "Americare in Home Nursing "
    elif filename == "2024_DHCF_HHRSD_ProviderSurvey_Fillable ASAP.pdf": 
        hh_agency = "ASAP/Palisades Health Partners"
    elif filename == "Provider Operations Survey 1_Capital Care Inc.pdf": 
        hh_agency = "Capital Care Inc"
    elif filename == "2024_DHCF_HHRSD_ProviderSurvey_Fillable_v6_20241108 CV HHC.pdf": 
        hh_agency = "Capitol View"
    elif filename == "2024 _DHCF_HHRSD_ProviderSurvey_Fillable_v1_Community Care.pdf": 
        hh_agency = "Community Care Nursing"
    elif filename == "2024_DHCF_HHRSD_ProviderSurvey_Fillable (Direct Care Home Health Services LLC).pdf": 
        hh_agency = "Direct Care Home Health"
    elif filename == "Provider Survey Health Management Inc.pdf": 
        hh_agency = "HMI"
    elif filename == "HT 2024_DHCF_HHRSD_ProviderSurvey_Fillable_v6_20241108 Human Touch Home Care Agency.pdf": 
        hh_agency = "Human Touch "
    elif filename == "Ideal Nursing Services - 2024_DHCF_HHRSD_ProviderSurvey_Fillable_v1.pdf": 
        hh_agency = "Ideal "
    elif filename == "2024_DHCF_HHRSD_ProviderSurvey_Immaculate_v6_20241108.pdf": 
        hh_agency = "Immaculate "
    elif filename == "2024_DHCF_HHRSD_ProviderSurvey_IntegratedCommunityServices.pdf": 
        hh_agency = "Integrated Community Service"
    elif filename == "Provider Operations Survey_KBC Nursing Agency and Home Health Care inc.pdf": 
        hh_agency = "KBC"
    elif filename == "Maxim Healthcare Services 2024 HH Provider Survey.pdf": 
        hh_agency = "Maxim "
    elif filename == "2024_DHCF_HHRSD_ProviderSurvey_Fillable_v1 _MEIGERHEALTH1.pdf": 
        hh_agency = "Meiger Health"
    elif filename == "2024_DHCF_HHRSD_ProviderSurvey_Fillable_v1-1.pdf_Open_Systems_Healthcarepdf.pdf": 
        hh_agency = "Open Systems Healthcare"
    elif filename == "2024_DHCF_HHRSD_ProviderSurvey_Optum WCH-SBerry.pdf": 
        hh_agency = "Optum WCH"
    elif filename == "2024_DHCF_HHRSD_ProviderSurvey_Fillable (Premier Health Services Inc).pdf": 
        hh_agency = "Premier "
    elif filename == "2024_DHCF_HHRSD_ProviderSurvey_Fillable_v6_20241108_Premiumselecthomecare2.pdf": 
        hh_agency = "Premium Select"
    elif filename == "DHCF PROVIDER SURVEY 2024_PHRI Updated.pdf": 
        hh_agency = "Professional Healthcare Resources Inc."
    elif filename == "2024 T&N Reliable Nursing Care Provider Survey.pdf": 
        hh_agency = "T&N Reliable "
    else:
        hh_agency = None  # Or set a default value if needed
    
    row.append(hh_agency)  # Add the HH_Agency value
    
    # Add field values aligned with reordered headers
    for field_name in reordered_field_names:
        row.append(fields.get(field_name, None))  # Add the field value or None if not present
    sheet.append(row)

# Save the Excel file
output_file = r"C:\Users\James Howard\DHCF_Survey\Operation Survey.xlsx"
workbook.save(output_file)

print(f"Form data successfully written to {output_file}")
print(filename)
