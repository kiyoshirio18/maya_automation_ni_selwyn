import streamlit as st
from io import BytesIO
from resources.excel_tools import xlsx_to_xls

st.header("Convert XLSX to XLS")
st.write(f''':green[:green-background[XLSX]]''')

# Create form for file upload
with st.form("convert"):
    files = st.file_uploader("Upload XLSX files to convert", type="xlsx", accept_multiple_files=True)

    # Process conversion when the form is submitted
    if st.form_submit_button("Convert", use_container_width=True):
        if files:
            converted_files = []
            for file in files:
                try:
                    # Read the uploaded file's bytes
                    input_bytes = file.read()

                    # Call the xlsx_to_xls function
                    xls_data = xlsx_to_xls(input_bytes)

                    if xls_data:
                        converted_files.append((file.name, xls_data))
                    else:
                        st.error(f"Conversion failed for {file.name}")
                
                except Exception as e:
                    st.error(f"Error processing {file.name}: {e}")
            
            if converted_files:
                st.success(f"Successfully Converted {len(converted_files)} files.")
                # Store converted files outside of the form
                st.session_state.converted_files = converted_files
            else:
                st.warning("No files were successfully converted.")
        else:
            st.warning("Please upload at least one file.")

# Provide download buttons for the converted files
if 'converted_files' in st.session_state:
    for name, xls_data in st.session_state.converted_files:
        st.download_button(
            label=f"Download {name.replace('.xlsx', '.xls')}",
            data=xls_data,
            file_name=name.replace('.xlsx', '.xls'),
            mime="application/vnd.ms-excel",
            use_container_width=True
        )