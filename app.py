import streamlit as st
import os
import pandas as pd
from datetime import datetime
import requests
from requests.auth import HTTPBasicAuth
import time
from PyPDF2 import PdfMerger, PdfReader
import tempfile

def main():
    # Set the title of the app
    st.title('EML Conversion and PDF Combination App')

    # Prompt the user for the folder path
    folder_path = st.text_input('Enter the folder path where your files are located:')

    if folder_path:
        # Check if the folder exists
        if os.path.isdir(folder_path):
            # Get all the file names in the folder along with their details
            file_details = []
            for file_name in os.listdir(folder_path):
                file_path = os.path.join(folder_path, file_name)
                if os.path.isfile(file_path):
                    file_date = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
                    file_type = os.path.splitext(file_name)[1]
                    file_details.append([file_name, file_date, file_type])

            # Create a DataFrame from the file details
            df = pd.DataFrame(file_details, columns=['File Name', 'Date Modified', 'File Type'])
            df.index = df.index + 1  # Start index at 1

            # Display the DataFrame
            st.write("Files in the folder:")
            st.dataframe(df)

            # Prompt the user to input the indices
            user_input = st.text_input("Choose the files you want to convert (e.g., '1, 2-4'): ")

            if user_input:
                # Parse the user input
                selected_indices = set()
                for part in user_input.split(','):
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
                st.dataframe(selected_df)

                # Prompt for Zamzar API key (required for .eml to .pdf conversion)
                api_key = st.text_input('Enter your Zamzar API key (required for .eml to .pdf conversion):', type='password')

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
                        file_path = os.path.join(folder_path, file_name)

                        if file_type == '.eml':
                            if api_key == '':
                                st.error('API key is required for .eml to .pdf conversion.')
                                st.stop()

                            st.write(f"Converting {file_name} to PDF...")
                            # Zamzar API conversion
                            endpoint = "https://sandbox.zamzar.com/v1/jobs"
                            target_format = "pdf"

                            with open(file_path, 'rb') as f:
                                file_content = {'source_file': f}
                                data_content = {'target_format': target_format}
                                response = requests.post(endpoint, data=data_content, files=file_content, auth=HTTPBasicAuth(api_key, ''))

                            if response.status_code != 201:
                                st.error(f"Error starting conversion job for {file_name}.")
                                st.stop()

                            job = response.json()
                            job_id = job['id']

                            # Poll the job status
                            status_endpoint = f"https://sandbox.zamzar.com/v1/jobs/{job_id}"
                            with st.spinner(f"Converting {file_name}..."):
                                while True:
                                    response = requests.get(status_endpoint, auth=HTTPBasicAuth(api_key, ''))
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
                            output_pdf_path = os.path.join(folder_path, output_pdf_name)  # Save in original folder

                            response = requests.get(download_endpoint, stream=True, auth=HTTPBasicAuth(api_key, ''))
                            if response.status_code == 200:
                                with open(output_pdf_path, 'wb') as pdf_file:
                                    for chunk in response.iter_content(chunk_size=1024):
                                        if chunk:
                                            pdf_file.write(chunk)
                                st.write(f"PDF created: {output_pdf_name} in original folder")

                                # Delete the original .eml file
                                os.remove(file_path)
                                st.write(f"Deleted original EML file: {file_name}")

                                # Update the DataFrame
                                processed_df.at[index, 'File Name'] = output_pdf_name
                                processed_df.at[index, 'File Type'] = '.pdf'
                                processed_df.at[index, 'Date Modified'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                                # Update progress bar
                                eml_files_processed += 1
                                progress = eml_files_processed / num_eml_files
                                progress_bar.progress(progress)
                                progress_text.text(f"Converted {eml_files_processed} out of {num_eml_files} .eml files")
                            else:
                                st.error(f"Failed to download the converted PDF for {file_name}.")
                                st.stop()
                        elif file_type == '.pdf':
                            # PDF files remain in the original folder
                            pass  # No action needed
                        else:
                            st.error(f"Unsupported file type: {file_type}")
                            st.stop()

                    # Combine the PDF files
                    st.write("Combining PDF files...")
                    pdf_merger = PdfMerger()

                    for index, row in processed_df.iterrows():
                        if row['File Type'] == '.pdf':
                            pdf_file_path = os.path.join(folder_path, row['File Name'])
                            # Open the PDF file in binary mode
                            try:
                                with open(pdf_file_path, 'rb') as pdf_file:
                                    pdf_reader = PdfReader(pdf_file)
                                    pdf_merger.append(pdf_reader)
                            except Exception as e:
                                st.error(f"Error reading PDF file {row['File Name']}: {e}")
                                st.stop()

                    # Write the combined PDF
                    combined_pdf_path = os.path.join(folder_path, 'combined_output.pdf')
                    with open(combined_pdf_path, 'wb') as combined_pdf_file:
                        pdf_merger.write(combined_pdf_file)
                    pdf_merger.close()

                    st.success(f"Combined PDF created at: {combined_pdf_path}")
                    # Provide a download link
                    with open(combined_pdf_path, 'rb') as f:
                        btn = st.download_button(
                            label="Download Combined PDF",
                            data=f,
                            file_name="combined_output.pdf",
                            mime="application/pdf"
                        )
            else:
                st.info("Please enter the indices of the files you want to combine.")
        else:
            st.error("The provided folder path does not exist.")

if __name__ == '__main__':
    main()
