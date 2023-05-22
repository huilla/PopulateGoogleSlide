# This is a script for a streamlit app that extracts data from a Google Sheet
# and uses this data to populate a Google Slide.

# Preparation:
# 1. Create a Google Sheet 'Employee_Data' with the given employee data
# 2. Create a Google Slide 'Employee_Presentation' with the given template and placeholders
# 3. Create a new project on Google Cloud Platform
# 4. Enable Google Drive API & Google Sheets API & Google Slides API
# 5. Create a service account and download credentials.json
# 6. Share Google Sheet & Google Slide with the service account email address
# 5. Continue with the script (see below)

'# Extract information from a Google Sheet to a Google Slide'

# Import the necessary libraries
# If not installed use: pip install 'library name'
import streamlit as st
import gspread
from google.oauth2 import service_account
from googleapiclient.discovery import build

def main():
    # Load the service account credentials for Google Sheet and open the sheet
    sa = gspread.service_account(filename='credentials.json')
    sheet = sa.open('Employee_Data')
    wks = sheet.worksheet('Sheet1')
    
    # Get all the values from the worksheet
    data = wks.get_all_values()
    st.title('Employee Data')
    # Display the data in a table using Streamlit
    st.table(data)

    # Write a function that extracts the necessary information from the Google Sheet
    def extract_information(name):
        name_cells = wks.findall(name)  # find the cell with the given name
        if not name_cells:
            print(f"No matching name found: {name}")
            exit(1)

        row_number = name_cells[0].row
        data = wks.row_values(row_number)

        # Extract the necessary information
        id = data[0]
        name = data[1]
        occupation = data[2]
        country = data[3]
        age = data[4]

        # Call the function to update the Google Slide
        write_to_google_slide(slide_id, id, name, country, age, occupation)

    # Load the service account credentials for Google Slides
    slides_credentials = service_account.Credentials.from_service_account_file('credentials.json')
    slides_service = build('slides', 'v1', credentials=slides_credentials)
    slide_id = '1kIIMyB9DB6X3XNTX0ofYhbV23oshT6lBvfHCecAUH0s'

    # Create a function to write data to the Google Slide
    def write_to_google_slide(slide_id, emp_id, emp_name, emp_country, emp_age, emp_occupation):
        presentation = slides_service.presentations().get(presentationId=slide_id).execute()
        slide_objects = presentation['slides'][0]['pageElements']

        # Iterate through the slide objects and populate the placeholders with data
        for obj in slide_objects:
            if 'shape' in obj and 'text' in obj['shape']:
                text = obj['shape']['text']
                if 'textElements' in text:
                    for text_element in text['textElements']:
                        if 'textRun' in text_element and 'content' in text_element['textRun']:
                            if text_element['textRun']['content'] == '[Placeholder for Name]':
                                text_element['textRun']['content'] = emp_name
                            elif text_element['textRun']['content'] == '[Placeholder for ID]':
                                text_element['textRun']['content'] = emp_id
                            elif text_element['textRun']['content'] == '[Placeholder for Occupation]':
                                text_element['textRun']['content'] = emp_occupation
                            elif text_element['textRun']['content'] == '[Placeholder for Country]':
                                text_element['textRun']['content'] = emp_country
                            elif text_element['textRun']['content'] == '[Placeholder for Age]':
                                text_element['textRun']['content'] = emp_age

            # Prepare the requests to update the text content
            reqs = [
                {'replaceAllText': {
                    'containsText': {'text': '[Placeholder for Name]'},
                    'replaceText': emp_name
                }},
                {'replaceAllText': {
                    'containsText': {'text': '[Placeholder for ID]'},
                    'replaceText': emp_id
                }},
                {'replaceAllText': {
                    'containsText': {'text': '[Placeholder for Occupation]'},
                    'replaceText': emp_occupation
                }},
                {'replaceAllText': {
                    'containsText': {'text': '[Placeholder for Country]'},
                    'replaceText': emp_country
                }},
                {'replaceAllText': {
                    'containsText': {'text': '[Placeholder for Age]'},
                    'replaceText': emp_age
                }},
            ]

            # Update the modified slide objects in the presentation
            request = slides_service.presentations().batchUpdate(
                presentationId=slide_id,
                body={
                    'requests': reqs
                }
            )
            request.execute()

    # Open the Google Slide in the browser showing the data from the Google Sheet
    def open_google_slide(slide_id):
        # Construct the presentation URL
        slide_url = f"https://docs.google.com/presentation/d/{slide_id}/edit"

        # Open the slide URL using the default web browser
        slides_service.presentations().get(presentationId=slide_id).execute()
        print("Opening Google Slide...")
        webbrowser.open(slide_url)
        
    # Add a text input field for the user to enter a name
    name_to_extract = st.text_input("Enter the name to extract information:")

    # Trigger the extraction and update the Google Slide when a name is entered
    if name_to_extract:
        extract_information(name_to_extract)
        open_google_slide(slide_id)
    
    # Test the prototype by asking the user to enter a name (e.g. John Doe)
    name_to_extract = input("Enter the name to extract information: ")
    extract_information(name_to_extract)
    open_google_slide(slide_id)

if __name__ == '__main__':
    main()
