import streamlit as st
import pandas as pd
from datetime import datetime
import requests
from requests.auth import HTTPBasicAuth
import time
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import tempfile
import os

def parse_page_ranges(input_str):
    try:
        page_numbers = set()
        for part in input_str.split(','):
            part = part.strip()
            if '-' in part:
                start, end = map(int, part.split('-'))
                if start > end:
                    return None
                page_numbers.update(range(start, end + 1))
            else:
                page_numbers.add(int(part))
        return sorted(page_numbers)
    except ValueError:
        return None

def split_pdf(file_row, page_numbers):
    try:
        # Get the PDF file path
        temp_file_path = file_row['Temp File Path']
        if not temp_file_path or not os.path.exists(temp_file_path):
            # If temp file does not exist, create it from the uploaded file
            uploaded_file = file_row.get('Data', None)
            if uploaded_file is not None:
                temp_file_path = os.path.join(st.session_state.temp_dir.name, file_row['File Name'])
                with open(temp_file_path, 'wb') as pdf_file:
                    pdf_file.write(uploaded_file.getvalue())
            else:
                st.error(f"No data found for PDF file {file_row['File Name']}")
                return None

        # Read the PDF
        with open(temp_file_path, 'rb') as pdf_file:
            pdf_reader = PdfReader(pdf_file)
            pdf_writer = PdfWriter()

            # Adjust page numbers to zero-based indexing
            max_page_number = len(pdf_reader.pages)
            selected_pages = [p - 1 for p in page_numbers if 1 <= p <= max_page_number]
            if not selected_pages:
                st.error("No valid page numbers selected.")
                return None

            # Create a new PDF with selected pages
            for page_num in selected_pages:
                pdf_writer.add_page(pdf_reader.pages[page_num])

            # Write the output PDF
            output_pdf_name = f"split_{file_row['File Name']}"
            output_pdf_path = os.path.join(st.session_state.temp_dir.name, output_pdf_name)
            with open(output_pdf_path, 'wb') as out_pdf_file:
                pdf_writer.write(out_pdf_file)

            return output_pdf_path
    except Exception as e:
        st.error(f"An error occurred while splitting the PDF: {e}")
        return None

def main():
    # Set the title of the app
    st.title('Email Conversion and PDF Manipulation App')

    # Detailed instructions
    st.markdown("""
    <h2><b>Instructions:</b></h2>

    1. **Upload Files:**
       - Use the file uploader below to upload your EML, MSG, and PDF files.
       - You can select multiple files at once.
       - Duplicate files will be ignored.

    2. **Select an Action:**
       - **Combine PDFs:**
         - Click on the 'Combine PDFs' button.
         - Enter the indices of the files you want to combine in a comma-separated list with either the index of the specific file or with a range of indices. (e.g., '1, 2-4').
         - If you include EML or MSG files, you will need to provide your Zamzar API key for conversion. If the user does not already have a key, they can sign up for a free one at [THIS LINK](https://developers.zamzar.com/signup?plan=test)—includes 100 free conversions per month, after which the user can either pay or use a new email to generate another key. If the selected files do not include a .eml/.msg attatchement, the user can substitute any text for the API key to continue combining PDFs as normal.
         - Click 'Convert and Combine Selected Files' to start the process.
       - **Split PDF:**
         - Click on the 'Split PDF' button.
         - Enter the index of the PDF file you want to split.
         - Specify the page numbers or ranges in the same manner as above to include in the new PDF (e.g., '1, 3-5').
         - Click 'Split PDF' to get the new PDF.

    3. **Download Results:**
       - After processing, you can download the combined or split PDF file.

    4. **Repeat or Upload More Files:**
       - You can upload more files at any time.
       - After each action, the app will loop back to allow you to perform another action.
       - The app supports multiple actions without losing previously uploaded files.
       - When users choose to access files they've already converted, those files are stored in a temporary cache to eliminate redundant API requests.
    
    
    """, unsafe_allow_html=True)

    # Initialize session state variables
    if 'file_details' not in st.session_state:
        st.session_state.file_details = []
    if 'converted_files' not in st.session_state:
        st.session_state.converted_files = {}
    if 'temp_dir' not in st.session_state:
        st.session_state.temp_dir = tempfile.TemporaryDirectory()
    if 'api_key' not in st.session_state:
        st.session_state.api_key = ''
    if 'user_input' not in st.session_state:
        st.session_state.user_input = ''
    if 'action_selected' not in st.session_state:
        st.session_state.action_selected = None

    # File uploader for multiple files (always available)
    uploaded_files = st.file_uploader(
        "Upload your EML, MSG, and PDF files (you can select multiple files):",
        type=['eml', 'msg', 'pdf'],
        accept_multiple_files=True
    )
    if uploaded_files:
        # Process each uploaded file
        for uploaded_file in uploaded_files:
            file_name = uploaded_file.name
            file_type = os.path.splitext(file_name)[1]
            file_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            # Check if the file is already in the list to avoid duplicates
            if not any(fd['File Name'] == file_name for fd in st.session_state.file_details):
                st.session_state.file_details.append({
                    'File Name': file_name,
                    'Date Modified': file_date,
                    'File Type': file_type,
                    'Data': uploaded_file,
                    'Temp File Path': None
                })
            else:
                # Duplicate file - ignore without warning
                pass

    if st.session_state.file_details:
        # Create a DataFrame from the file details
        df = pd.DataFrame(st.session_state.file_details)
        df.index = df.index + 1  # Start index at 1

        # Display the DataFrame
        st.write("Available Files:")
        st.dataframe(df[['File Name', 'File Type']])

        # Action selection
        st.write("Select an action:")
        col1, col2 = st.columns(2)
        with col1:
            if st.button('Combine PDFs'):
                st.session_state.action_selected = 'combine'
        with col2:
            if st.button('Split PDF'):
                st.session_state.action_selected = 'split'

        if st.session_state.action_selected == 'combine':
            # Prompt the user to input the indices
            st.session_state.user_input = st.text_input(
                "Choose the files you want to convert (e.g., '1, 2-4'): ",
                value=st.session_state.user_input
            )

            if st.session_state.user_input:
                # Parse the user input
                selected_indices = set()
                for part in st.session_state.user_input.split(','):
                    part = part.strip()
                    if '-' in part:
                        try:
                            start, end = map(int, part.split('-'))
                            if start > end:
                                st.error(f"Invalid range '{part}': start index is greater than end index.")
                                st.stop()
                            selected_indices.update(range(start, end + 1))  # Keep one-based indexing
                        except ValueError:
                            st.error(f"Invalid range in input: '{part}'")
                            st.stop()
                    else:
                        try:
                            selected_indices.add(int(part))
                        except ValueError:
                            st.error(f"Invalid number in input: '{part}'")
                            st.stop()

                # Filter the DataFrame based on the selected indices
                selected_indices = sorted([i for i in selected_indices if 1 <= i <= len(df)])  # Keep one-based indexing

                if not selected_indices:
                    st.error("No valid indices selected.")
                    st.stop()

                selected_df = df.loc[selected_indices]
                st.write("Selected Files:")
                st.dataframe(selected_df[['File Name', 'File Type']])

                # Prompt for Zamzar API key (required for .eml and .msg to .pdf conversion)
                st.session_state.api_key = st.text_input(
                    'Enter your Zamzar API key (required for .eml and .msg to .pdf conversion):',
                    type='password',
                    value=st.session_state.api_key
                )

                # Button to start conversion and combination
                if st.button('Convert and Combine Selected Files'):
                    # Copy the selected DataFrame to avoid modifying the original
                    processed_df = selected_df.copy()

                    # Count the number of .eml and .msg files
                    num_email_files = processed_df[processed_df['File Type'].isin(['.eml', '.msg'])].shape[0]
                    email_files_processed = 0  # Initialize the counter

                    # Initialize the progress bar if there are .eml or .msg files
                    if num_email_files > 0:
                        progress_bar = st.progress(0)
                        progress_text = st.empty()  # Placeholder for progress text

                    # Iterate through the selected files
                    for idx, (index, row) in enumerate(processed_df.iterrows(), start=1):
                        file_name = row['File Name']
                        file_type = row['File Type']
                        uploaded_file = row.get('Data', None)
                        temp_file_path = row.get('Temp File Path', None)

                        if file_type in ['.eml', '.msg']:
                            if st.session_state.api_key == '':
                                st.error('API key is required for .eml and .msg to .pdf conversion.')
                                st.stop()

                            # Check if this file has already been converted
                            if file_name in st.session_state.converted_files:
                                st.write(f"Using cached PDF for {file_name}")
                                converted_info = st.session_state.converted_files[file_name]
                                output_pdf_name = converted_info['File Name']
                                output_pdf_path = converted_info['Temp File Path']
                            else:
                                st.write(f"Converting {file_name} to PDF...")
                                # Zamzar API conversion
                                endpoint = "https://sandbox.zamzar.com/v1/jobs"
                                target_format = "pdf"

                                # Prepare the file for upload
                                data_content = {'target_format': target_format}
                                files = {'source_file': (file_name, uploaded_file.getvalue())}

                                response = requests.post(
                                    endpoint,
                                    data=data_content,
                                    files=files,
                                    auth=HTTPBasicAuth(st.session_state.api_key, '')
                                )

                                if response.status_code != 201:
                                    st.error(f"Error starting conversion job for {file_name}.")
                                    st.stop()

                                job = response.json()
                                job_id = job['id']

                                # Poll the job status
                                status_endpoint = f"https://sandbox.zamzar.com/v1/jobs/{job_id}"
                                with st.spinner(f"Converting {file_name}..."):
                                    while True:
                                        response = requests.get(
                                            status_endpoint,
                                            auth=HTTPBasicAuth(st.session_state.api_key, '')
                                        )
                                        job_status = response.json()
                                        status = job_status['status']
                                        if status == 'successful':
                                            break
                                        elif status == 'failed':
                                            st.error(f"The conversion job for {file_name} failed.")
                                            st.stop()
                                        time.sleep(5)  # Wait before polling again

                                # Download the converted PDF
                                target_file_id = job_status['target_files'][0]['id']
                                download_endpoint = f"https://sandbox.zamzar.com/v1/files/{target_file_id}/content"
                                output_pdf_name = os.path.splitext(file_name)[0] + '.pdf'
                                output_pdf_path = os.path.join(
                                    st.session_state.temp_dir.name,
                                    output_pdf_name
                                )  # Save in temporary folder

                                response = requests.get(
                                    download_endpoint,
                                    stream=True,
                                    auth=HTTPBasicAuth(st.session_state.api_key, '')
                                )
                                if response.status_code == 200:
                                    with open(output_pdf_path, 'wb') as pdf_file:
                                        for chunk in response.iter_content(chunk_size=1024):
                                            if chunk:
                                                pdf_file.write(chunk)
                                    st.write(f"PDF created: {output_pdf_name}")

                                    # Update the session_state with converted file
                                    st.session_state.converted_files[file_name] = {
                                        'File Name': output_pdf_name,
                                        'File Type': '.pdf',
                                        'Date Modified': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                        'Temp File Path': output_pdf_path
                                    }

                                    # Update progress bar
                                    email_files_processed += 1
                                    progress = email_files_processed / num_email_files
                                    progress_bar.progress(progress)
                                    progress_text.text(
                                        f"Converted {email_files_processed} out of {num_email_files} email files"
                                    )
                                else:
                                    st.error(f"Failed to download the converted PDF for {file_name}.")
                                    st.stop()
                            # Update processed_df with converted file info
                            processed_df.at[index, 'File Name'] = output_pdf_name
                            processed_df.at[index, 'File Type'] = '.pdf'
                            processed_df.at[index, 'Date Modified'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            processed_df.at[index, 'Temp File Path'] = output_pdf_path

                        elif file_type == '.pdf':
                            if temp_file_path and os.path.exists(temp_file_path):
                                output_pdf_path = temp_file_path
                            else:
                                if uploaded_file is not None:
                                    output_pdf_path = os.path.join(st.session_state.temp_dir.name, file_name)
                                    with open(output_pdf_path, 'wb') as pdf_file:
                                        pdf_file.write(uploaded_file.getvalue())
                                    processed_df.at[index, 'Temp File Path'] = output_pdf_path
                                else:
                                    st.error(f"No data found for PDF file {file_name}")
                                    st.stop()
                        else:
                            st.error(f"Unsupported file type: {file_type}")
                            st.stop()

                        # Update processed_df with Temp File Path
                        processed_df.at[index, 'Temp File Path'] = output_pdf_path

                    # Combine the PDF files
                    st.write("Combining PDF files...")
                    pdf_merger = PdfMerger()

                    for index, row in processed_df.iterrows():
                        if row['File Type'] == '.pdf':
                            pdf_file_path = row['Temp File Path']
                            # Open the PDF file in binary mode
                            try:
                                with open(pdf_file_path, 'rb') as pdf_file:
                                    pdf_reader = PdfReader(pdf_file)
                                    pdf_merger.append(pdf_reader)
                            except Exception as e:
                                st.error(f"Error reading PDF file {row['File Name']}: {e}")
                                st.stop()

                    # Write the combined PDF
                    combined_pdf_name = f'combined_output_{datetime.now().strftime("%Y%m%d%H%M%S")}.pdf'
                    combined_pdf_path = os.path.join(st.session_state.temp_dir.name, combined_pdf_name)
                    with open(combined_pdf_path, 'wb') as combined_pdf_file:
                        pdf_merger.write(combined_pdf_file)
                    pdf_merger.close()

                    # Provide a download link
                    with open(combined_pdf_path, 'rb') as f:
                        st.success("Combined PDF is ready!")
                        download_clicked = st.download_button(
                            label="Download Combined PDF",
                            data=f,
                            file_name=combined_pdf_name,
                            mime="application/pdf"
                        )

                    # Automatically reset the action selected
                    if download_clicked:
                        # Update the file list to replace '.eml' and '.msg' files with their converted '.pdf' versions
                        updated_file_details = []
                        for file_detail in st.session_state.file_details:
                            file_name = file_detail['File Name']
                            if file_name in st.session_state.converted_files:
                                # Replace with converted PDF
                                converted_file = st.session_state.converted_files[file_name]
                                updated_file_details.append({
                                    'File Name': converted_file['File Name'],
                                    'Date Modified': converted_file['Date Modified'],
                                    'File Type': converted_file['File Type'],
                                    'Data': None,  # Data is None because we have Temp File Path
                                    'Temp File Path': converted_file['Temp File Path']
                                })
                            else:
                                updated_file_details.append(file_detail)
                        st.session_state.file_details = updated_file_details

                        # Clear the previous user input and action
                        st.session_state.user_input = ''
                        st.session_state.action_selected = None

                        # Rerun the app to reflect changes
                        st.experimental_rerun()

                else:
                    st.info("Click the button to convert and combine selected files.")
            else:
                st.info("Please enter the indices of the files you want to combine.")

        elif st.session_state.action_selected == 'split':
            # Prompt the user to select the index of the file to split
            split_file_input = st.text_input(
                "Choose the index of the PDF file you want to split (e.g., '1'): ",
                value=st.session_state.get('split_file_input', '')
            )
            st.session_state.split_file_input = split_file_input

            if split_file_input:
                try:
                    split_index = int(split_file_input.strip())
                    if split_index < 1 or split_index > len(df):
                        st.error("Invalid index selected.")
                        st.stop()
                    selected_file = df.loc[split_index]
                    if selected_file['File Type'] != '.pdf':
                        st.error("Selected file is not a PDF.")
                        st.stop()
                    else:
                        st.write(f"Selected file: {selected_file['File Name']}")
                        # Prompt for page ranges
                        page_ranges_input = st.text_input(
                            "Enter the page numbers or ranges you want to include (e.g., '1, 3-5'): ",
                            value=st.session_state.get('page_ranges_input', '')
                        )
                        st.session_state.page_ranges_input = page_ranges_input

                        if page_ranges_input:
                            # Button to split and download
                            if st.button('Split PDF'):
                                # Validate and parse the page ranges
                                page_numbers = parse_page_ranges(page_ranges_input)
                                if page_numbers is None:
                                    st.error("Invalid page ranges entered.")
                                    st.stop()
                                else:
                                    # Perform splitting
                                    split_pdf_path = split_pdf(selected_file, page_numbers)
                                    if split_pdf_path:
                                        # Provide download link
                                        with open(split_pdf_path, 'rb') as f:
                                            st.success("Split PDF is ready!")
                                            download_clicked = st.download_button(
                                                label="Download Split PDF",
                                                data=f,
                                                file_name=f"split_{selected_file['File Name']}",
                                                mime="application/pdf"
                                            )
                                        if download_clicked:
                                            # Reset action
                                            st.session_state.action_selected = None
                                            st.session_state.split_file_input = ''
                                            st.session_state.page_ranges_input = ''
                                            # Rerun the app to loop back
                                            st.experimental_rerun()
                                    else:
                                        st.error("Failed to split the PDF.")
                                        st.stop()
                        else:
                            st.info("Please enter the page ranges to include.")
                except ValueError:
                    st.error("Invalid index entered.")
                    st.stop()
            else:
                st.info("Please enter the index of the PDF file you want to split.")

        else:
            st.info("Please select an action to perform.")
    else:
        st.info("Please upload your EML, MSG, and PDF files to proceed.")

if __name__ == '__main__':
    main()
