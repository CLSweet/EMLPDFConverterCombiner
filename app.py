import streamlit as st
import pandas as pd
from datetime import datetime
import requests
from requests.auth import HTTPBasicAuth
import time
from PyPDF2 import PdfMerger, PdfReader
import tempfile
import os

def main():
    # Set the title of the app
    st.title('EML Conversion and PDF Combination App')

    # Initialize session state variables
    if 'file_details' not in st.session_state:
        st.session_state.file_details = []
    if 'converted_files' not in st.session_state:
        st.session_state.converted_files = {}
    if 'uploaded_files' not in st.session_state:
        st.session_state.uploaded_files = None
    if 'temp_dir' not in st.session_state:
        st.session_state.temp_dir = tempfile.TemporaryDirectory()
    if 'api_key' not in st.session_state:
        st.session_state.api_key = ''
    if 'user_input' not in st.session_state:
        st.session_state.user_input = ''

    # File uploader for multiple files
    if st.session_state.uploaded_files is None:
        uploaded_files = st.file_uploader(
            "Upload your EML and PDF files (you can select multiple files):",
            type=['eml', 'pdf'],
            accept_multiple_files=True
        )
        if uploaded_files:
            st.session_state.uploaded_files = uploaded_files
            # Process each uploaded file
            for uploaded_file in uploaded_files:
                file_name = uploaded_file.name
                file_type = os.path.splitext(file_name)[1]
                file_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                st.session_state.file_details.append({
                    'File Name': file_name,
                    'Date Modified': file_date,
                    'File Type': file_type,
                    'Data': uploaded_file,
                    'Temp File Path': None
                })

    if st.session_state.file_details:
        # Create a DataFrame from the file details
        df = pd.DataFrame(st.session_state.file_details)
        df.index = df.index + 1  # Start index at 1

        # Display the DataFrame
        st.write("Available Files:")
        st.dataframe(df[['File Name', 'Date Modified', 'File Type']])

        # Prompt the user to input the indices
        st.session_state.user_input = st.text_input("Choose the files you want to convert (e.g., '1, 2-4'): ", value=st.session_state.user_input)

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
            st.dataframe(selected_df[['File Name', 'Date Modified', 'File Type']])

            # Prompt for Zamzar API key (required for .eml to .pdf conversion)
            st.session_state.api_key = st.text_input('Enter your Zamzar API key (required for .eml to .pdf conversion):', type='password', value=st.session_state.api_key)

            # Button to start conversion and combination
            if st.button('Convert and Combine Selected Files'):
                # Copy the selected DataFrame to avoid modifying the original
                processed_df = selected_df.copy()

                # Count the number of .eml files
                num_eml_files = processed_df[processed_df['File Type'] == '.eml'].shape[0]
                eml_files_processed = 0  # Initialize the counter

                # Initialize the progress bar if there are .eml files
                if num_eml_files > 0:
                    progress_bar = st.progress(0)
                    progress_text = st.empty()  # Placeholder for progress text

                # Iterate through the selected files
                for idx, (index, row) in enumerate(processed_df.iterrows(), start=1):
                    file_name = row['File Name']
                    file_type = row['File Type']
                    uploaded_file = row.get('Data', None)
                    temp_file_path = row.get('Temp File Path', None)

                    if file_type == '.eml':
                        if st.session_state.api_key == '':
                            st.error('API key is required for .eml to .pdf conversion.')
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

                            response = requests.post(endpoint, data=data_content, files=files, auth=HTTPBasicAuth(st.session_state.api_key, ''))

                            if response.status_code != 201:
                                st.error(f"Error starting conversion job for {file_name}.")
                                st.stop()

                            job = response.json()
                            job_id = job['id']

                            # Poll the job status
                            status_endpoint = f"https://sandbox.zamzar.com/v1/jobs/{job_id}"
                            with st.spinner(f"Converting {file_name}..."):
                                while True:
                                    response = requests.get(status_endpoint, auth=HTTPBasicAuth(st.session_state.api_key, ''))
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
                            output_pdf_name = file_name.replace('.eml', '.pdf')
                            output_pdf_path = os.path.join(st.session_state.temp_dir.name, output_pdf_name)  # Save in temporary folder

                            response = requests.get(download_endpoint, stream=True, auth=HTTPBasicAuth(st.session_state.api_key, ''))
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
                                eml_files_processed += 1
                                progress = eml_files_processed / num_eml_files
                                progress_bar.progress(progress)
                                progress_text.text(f"Converted {eml_files_processed} out of {num_eml_files} .eml files")
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
                    st.download_button(
                        label="Download Combined PDF",
                        data=f,
                        file_name=combined_pdf_name,
                        mime="application/pdf"
                    )

                # Prompt the user if they want to combine more files
                st.write("Do you want to combine more files?")
                with combine_more:
                    combine_more_clicked = st.button('Yes')
                with exit_app:
                    exit_app_clicked = st.button('No')

                if combine_more_clicked:
                    # Update the dataframe to replace '.eml' files with their converted '.pdf' versions
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

                    # Recreate the DataFrame with updated file types
                    df = pd.DataFrame(st.session_state.file_details)
                    df.index = df.index + 1  # Start index at 1

                    # Clear the previous user input and selected files
                    st.session_state.user_input = ''

                    # Clear previous selected files output
                    st.experimental_rerun()
                elif exit_app_clicked:
                    st.write("Thank you for using the app!")
                    # Clear session_state to start over
                    st.session_state.clear()
                    st.stop()
        else:
            st.info("Please enter the indices of the files you want to combine.")
    else:
        st.info("Please upload your EML and PDF files to proceed.")

if __name__ == '__main__':
    main()
