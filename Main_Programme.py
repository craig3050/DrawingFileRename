import streamlit as st
import pandas as pd
import io
import os
import zipfile
import shutil


def create_excel_download_link(df, original_title_col):
    """
    Create an Excel file in memory and return a link for downloading it.

    :param df: DataFrame containing file names.
    :param original_title_col: Column name in df that contains the original file names.
    :return: A BytesIO object containing the Excel file, or None if DataFrame is empty.
    """
    if df.empty:
        return None

    if original_title_col not in df.columns:
        raise ValueError(f"Column '{original_title_col}' not found in DataFrame")

    export_df = pd.DataFrame({
        'Original Title': df[original_title_col],
        'Renamed Title': ['' for _ in range(len(df))]
    })

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False)

    output.seek(0)
    return output


def save_uploaded_files(uploaded_files):
    temp_dir = "temp_uploaded_files"
    os.makedirs(temp_dir, exist_ok=True)

    for file in uploaded_files:
        file_path = os.path.join(temp_dir, file.name)
        with open(file_path, "wb") as f:
            f.write(file.getbuffer())

    return temp_dir
def process_and_zip_files(drawings_df, temp_dir):
    renamed_dir = "temp_renamed_files"
    os.makedirs(renamed_dir, exist_ok=True)

    for index, row in drawings_df.iterrows():
        original_name = row['Original Name']
        new_name_base = str(row['New Name']) if pd.notna(row['New Name']) and row['New Name'] != '' else original_name
        _, extension = os.path.splitext(original_name)
        new_name = new_name_base + extension  # Append the original file extension

        original_path = os.path.join(temp_dir, original_name)
        new_path = os.path.join(renamed_dir, new_name)

        if os.path.exists(original_path):
            shutil.copy(original_path, new_path)
            print(f"File '{original_name}' renamed to '{new_name}'")  # Diagnostic print


    # Create a zip file
    zip_name = "renamed_files.zip"
    with zipfile.ZipFile(zip_name, 'w') as zipf:
        for file in os.listdir(renamed_dir):
            zipf.write(os.path.join(renamed_dir, file), file)

    # Clean up the temporary directories
    shutil.rmtree(temp_dir)
    shutil.rmtree(renamed_dir)

    return zip_name
def main():
    st.title('Drawing Renamer Tool')

    # Initialize a session state to store the DataFrame
    if 'drawings_df' not in st.session_state:
        st.session_state['drawings_df'] = pd.DataFrame(columns=['Original Name', 'New Name'])

    uploaded_files = st.file_uploader("Upload Drawings", accept_multiple_files=True, type=['png', 'jpg', 'jpeg', 'pdf'])

    # Save uploaded files immediately after they are uploaded
    if uploaded_files:
        # Check if the DataFrame already exists and has updated 'New Name' data
        if 'drawings_df' in st.session_state and not st.session_state['drawings_df']['New Name'].isnull().all():
            # Preserve the existing DataFrame with updated names
            existing_names = st.session_state['drawings_df'].set_index('Original Name')['New Name']
            file_names = [file.name for file in uploaded_files]
            new_df = pd.DataFrame(file_names, columns=['Original Name'])
            new_df['New Name'] = new_df['Original Name'].map(existing_names).fillna('')
            st.session_state['drawings_df'] = new_df
        else:
            # Initialize a new DataFrame
            temp_dir_path = save_uploaded_files(uploaded_files)
            st.session_state['temp_dir_path'] = temp_dir_path
            file_names = [file.name for file in uploaded_files]
            st.session_state['drawings_df'] = pd.DataFrame(file_names, columns=['Original Name'])
            st.session_state['drawings_df']['New Name'] = ''

    # Button to toggle rename options
    rename_option = st.radio(
        "Select Rename Option",
        ('Rename from title block', 'Rename from Excel sheet')
    )
    st.write("Selected Option:", rename_option)

    if rename_option == 'Rename from title block':
        # Show 'Search for names in title block' button
        if st.button('Search for names in title block'):
            # Your logic here
            pass

    elif rename_option == 'Rename from Excel sheet':
        # Show 'Export to Excel' and 'Import Excel file' options
        if st.button('Export to Excel'):
            # Check if the DataFrame is not empty
            if not st.session_state['drawings_df'].empty:
                excel_file = create_excel_download_link(st.session_state['drawings_df'], 'Original Name')
                st.markdown("## Download Excel File")
                st.markdown("Click the button below to download the Excel file with the drawing names.")
                st.download_button(
                    label="Download Excel file",
                    data=excel_file,
                    file_name="exported_file_names.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No file names to export.")

        # File uploader for Excel file
        uploaded_excel = st.file_uploader("Upload Excel File for Renaming", type=['xlsx'], key='excel_uploader')

        if st.button('Import Excel File'):
            if uploaded_excel is not None:
                try:
                    imported_df = pd.read_excel(uploaded_excel, engine='openpyxl')
                except Exception as e:
                    st.error(f"Error reading the Excel file: {e}")
                    return

                if 'Renamed Title' in imported_df.columns and 'Original Title' in imported_df.columns:
                    name_map = dict(zip(imported_df['Original Title'], imported_df['Renamed Title']))

                    # Create a local copy of the DataFrame for updating
                    updated_df = st.session_state['drawings_df'].copy()

                    for original, new_name in name_map.items():
                        matched_rows = updated_df['Original Name'] == original
                        if matched_rows.any():
                            updated_df.loc[matched_rows, 'New Name'] = new_name
                            print(f"Updated '{original}' to '{new_name}'")  # Diagnostic print statement

                    # Assign the updated DataFrame back to the session state
                    st.session_state['drawings_df'] = updated_df
                    print("DataFrame updated in session state:")  # Diagnostic print
                    print(st.session_state['drawings_df'])  # Diagnostic print to confirm update
                    st.success("File names updated from the Excel file.")
                else:
                    st.error("The uploaded Excel file does not have the required columns.")
            else:
                st.error("No Excel file uploaded. Please upload a file to process.")

    st.write('Uploaded Drawings:')
    st.dataframe(st.session_state['drawings_df'], use_container_width=True)

    # Process Rename button
    if st.button('Process Rename', key='process_rename_button', use_container_width=True):
        if 'temp_dir_path' in st.session_state and os.path.exists(st.session_state['temp_dir_path']):
            zip_file = process_and_zip_files(st.session_state['drawings_df'], st.session_state['temp_dir_path'])
            with open(zip_file, "rb") as f:
                st.download_button("Download Renamed Files", f, file_name=zip_file, mime="application/zip")
            os.remove(zip_file)
        else:
            st.error("No files uploaded or temporary directory missing.")


if __name__ == "__main__":
    main()
