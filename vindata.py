import streamlit as st
import pandas as pd
import requests
import openpyxl
import os
import io
import json

def vin_data(file_bytes, original_filename):
    wb = openpyxl.load_workbook(file_bytes)
    res = len(wb.sheetnames)

    if res > 1:
        raw_vin_data = pd.read_excel(file_bytes, 'Vehicle & Asset List', header=3)
    else:
        raw_vin_data = pd.read_excel(file_bytes, header=3)

    for column in raw_vin_data.columns:
        if 'vin' in column.lower():
            raw_vin_data.rename(columns={column: 'VIN'}, inplace=True)

    base_url = 'https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVin/'

    vin_data = pd.DataFrame(columns=[
        'VIN', 'VIN Mask', 'Model Year', 'Manufacturer', 'Make', 'Model', 'Trim', 
        'Weight Class', 'Body/Cab Type', 'Body Class', 'Drive Type', 'Fuel Type', 
        'Engine Model', 'Engine Configuration', 'Engine Cyl', 'Displacement (Litres)', 
        'Engine Horse Power', 'Transmission', 'Speeds', 'Error Test'
    ])

    values = [raw_vin_data['VIN'][i] for i in raw_vin_data.index if pd.notna(raw_vin_data['VIN'][i])]

    results = []
    ind = 0

    for value in values:
        value = str(value).replace(" ", "")
        url = base_url + value + '?format=json'
        response = requests.get(url, verify=False)
        try:
            data = response.json()
            decoded_values = {item['Variable']: item['Value'] for item in data['Results']}
            results.append({
                'VIN': value, 
                'VIN Mask': decoded_values.get('Vehicle Descriptor', 'N/A'), 
                'Model Year': decoded_values.get('Model Year', 'N/A'), 
                'Manufacturer': decoded_values.get('Manufacturer Name', 'N/A'), 
                'Make': decoded_values.get('Make', 'N/A'), 
                'Model': decoded_values.get('Model', 'N/A'), 
                'Trim': decoded_values.get('Trim', 'N/A'), 
                'Weight Class': decoded_values.get('Gross Vehicle Weight Rating From', 'N/A'),
                'Body/Cab Type': decoded_values.get('Cab Type', 'N/A'), 
                'Body Class': decoded_values.get('Body Class', 'N/A'), 
                'Drive Type': decoded_values.get('Drive Type', 'N/A'),
                'Fuel Type': decoded_values.get('Fuel Type - Primary', 'N/A'), 
                'Engine Model': decoded_values.get('Engine Model', 'N/A'), 
                'Engine Configuration': decoded_values.get('Engine Configuration', 'N/A'),
                'Engine Cyl': decoded_values.get('Engine Number of Cylinders', 'N/A'), 
                'Displacement (Litres)': decoded_values.get('Displacement (L)', 'N/A'), 
                'Engine Horse Power': decoded_values.get('Engine Brake (hp) From', 'N/A'),
                'Transmission': decoded_values.get('Transmission Style', 'N/A'), 
                'Speeds': decoded_values.get('Transmission Speeds', 'N/A'), 
                'Error Test': decoded_values.get('Error Text', 'N/A')
            })
            ind += 1
        except json.JSONDecodeError:
            results.append({
                'VIN': value, 
                'VIN Mask': 'Error', 
                'Model Year': 'Error', 
                'Manufacturer': 'Error', 
                'Make': 'Error', 
                'Model': 'Error', 
                'Trim': 'Error', 
                'Weight Class': 'Error',
                'Body/Cab Type': 'Error', 
                'Body Class': 'Error', 
                'Drive Type': 'Error',
                'Fuel Type': 'Error', 
                'Engine Model': 'Error', 
                'Engine Configuration': 'Error',
                'Engine Cyl': 'Error', 
                'Displacement (Litres)': 'Error', 
                'Engine Horse Power': 'Error',
                'Transmission': 'Error', 
                'Speeds': 'Error', 
                'Error Test': 'Error: Incorrect VIN, no data exists'
            })
            ind += 1
        except requests.exceptions.Timeout:
            return "Timed out"

    results_df = pd.DataFrame(results)
    results_df.drop_duplicates(subset=['VIN'], inplace=True)

    buffer = io.BytesIO()
    processed_filename = os.path.splitext(original_filename)[0] + "_VIN_data.xlsx"

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        results_df.to_excel(writer, index=False, sheet_name='Vehicle Data')
        workbook = writer.book
        worksheet = writer.sheets['Vehicle Data']

        for idx, column in enumerate(worksheet.columns):
            max_length = 0
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[chr(65 + idx)].width = adjusted_width

    buffer.seek(0)
    return buffer, processed_filename

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css?family=Open+Sans');
body {
    font-family: 'Open Sans', sans-serif;
}
</style>
""", unsafe_allow_html=True)

st.image("https://www.tdtyres.com/wp-content/uploads/2018/12/kisspng-car-michelin-man-tire-logo-michelin-logo-5b4c286206fa03.5353854915317177300286.png")

st.title("VIN Vehicle Data")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx", "csv"])

if uploaded_file:
    file_bytes = uploaded_file.read()
    buffer, processed_filename = vin_data(io.BytesIO(file_bytes), uploaded_file.name)
    st.success(f'File "{uploaded_file.name}" successfully processed.')
    st.download_button(label="Download Processed File", data=buffer, file_name=processed_filename)

st.markdown('''
This application checks customer VINs with the [National Highway Traffic Safety Administration API](https://vpic.nhtsa.dot.gov/api/) to retrieve vehicle information based on the VIN. This application can handle large volumes of VINs but greater numbers of uploaded VINs will slow down processing time. Processing 2200 VINs takes roughly 25 minutes. When uploading large numbers of VINs please be patient and do not close out the application while processing.
**Input Document Requirements:**
- The uploaded document containing the VINs must follow the standard [Michelin Connected Fleet Deployment Template.](https://view.officeapps.live.com/op/view.aspx?src=https%3A%2F%2Fraw.githubusercontent.com%2FChanMichelin%2FautovinMCF%2Fmain%2Fexamples%2FMCF%2520Deployment%2520Template.xlsx&wdOrigin=BROWSELINK) This application cannot decipher different document formats. If an error is indicated with a file you upload, please check the uploaded document follows the formatting guidelines.
- The VIN column must include the VINs the user wants to query. This is the only field necessary to retrieve vehicle data. 
- Make sure the input document is not open on your computer. If the input document is open, a permission error will occur.
***Example Input File:*** [***VIN Example***](https://michelingroup.sharepoint.com/:x:/r/sites/ProcessImprovement/_layouts/15/Doc.aspx?sourcedoc=%7BFA264B31-B424-418C-8D1C-C0E5F001094E%7D&file=MCF%20Deployment%20Template.xlsx&action=default&mobileredirect=true&wdsle=0)
***Note:*** If you are interested in vehicle information regarding VINs recorded in a different format/document: download the MCF Deployment Template linked above, then copy and paste the VINs into the VIN column and upload this document for bulk processing.
**Output Document Description:**
- This application processes all the VINs regardless of VIN accuracy or vehicle type. 
- If the VIN is inaccurate or relates to a lift/trailer not present in the NHTSA database the 'Error' column will indicate what type of error is occurring for user reference. 
- An error code of 0 indicates there was no issue with the VIN. 
- This file provides information on vehicle make, model, year, and manufacturer as well as more detailed information pertaining to trim, engine type, primary fuel etc. 
- When a cell is empty, but the error column reports there was no issue processing the VIN (error code is 0) this indicates that data on this vehicle specification is not recorded within the NHTSA database. 
- The output Excel file will have the same name as the original document followed by _VIN_data. 
***Example Output File:*** [***VIN Example_VIN_data***](https://michelingroup.sharepoint.com/:x:/r/sites/ProcessImprovement/_layouts/15/Doc.aspx?sourcedoc=%7B7481464E-023E-4E40-9007-34AE4022EECE%7D&file=VIN%20Example_VIN_data.xlsx&action=default&mobileredirect=true&wdsle=0)
If you are interested in a list of accurate VINs that relate to CAN compatible vehicles excluding trailers and lifts, please refer to the [Automated VIN Decoding Application.](https://autovin.streamlit.app/)
''')
