import streamlit as st
import os
from file_audit import count_files_and_sizes, convert_xls_to_xlsx, analyze_excel_sheets, find_duplicate_files, is_file_corrupted
# from file_audit import analyze_pdf_files
import pandas as pd
import subprocess
import time
import seaborn as sns
import pyautogui
import folium
from streamlit_folium import st_folium
import re

import argparse




def main():

    # Use the full page instead of a narrow central column
    st.set_page_config(layout="wide")

    st.title("DrillParse")
    st.write("Automate and customized pre-processing of Unstructured Data of Daily Drilling Report (DDR) in a general-manner of extraction from various DDR templates.")

    # st.caption("Version Note")
    st.caption("DrillParse current version only accepts new .xlsx format only, not legacy .xls format. Please convert the legacy .xls format to .xlsx format first to prevent parsing error.")
    st.caption("Assumption of all DDRs templates across multiple sheets and xlsx files are consistent.")
    st.divider()

    st.subheader("Drilling Report Uploader",divider=True)
    selected_option = st.radio("Select Parser Pipeline", ["ExcelParser","PDFParser"])

    # Upload files
    uploaded_files = st.file_uploader("Upload files", accept_multiple_files=True)

    if st.button("Reset Process"):
                pyautogui.hotkey("ctrl","F5")

    

    # Divide into two columns
    col11, col22 = st.columns(2, gap="medium")

    if uploaded_files: 

        with col11:


            st.subheader("File Audit",divider=True)
            st.write("Audit for file sheet counts, file sizes, duplicated files, corrupted files and file extension types for resource estimation")



            if uploaded_files:
                # Create a temporary directory to store the uploaded files
                temp_directory = "temp_files"
                os.makedirs(temp_directory, exist_ok=True)

                # Save the uploaded files to the temporary directory
                for uploaded_file in uploaded_files:
                    file_path = os.path.join(temp_directory, uploaded_file.name)
                    with open(file_path, "wb") as file:
                        file.write(uploaded_file.getvalue())
            
                    
                # Run file audit workflow
                file_types, file_sizes, file_info = count_files_and_sizes(temp_directory)
                convert_xls_to_xlsx(temp_directory)

                                # Check for duplicated files
                duplicated_files = find_duplicate_files(temp_directory)

                # Display audit statistics
                for file_type, count in file_types.items():
                    if count != 0:
                        
                        non_zero_file_types = {key: value for key, value in file_types.items() if value != 0}
                        non_zero_file_sizes = {key: value for key, value in file_sizes.items() if key == file_type}

                        
                        if non_zero_file_types:

                            total_sheets_xlsx, max_sheets_xlsx, min_sheets_xlsx, average_sheets_xlsx, total_sheets_xls, max_sheets_xls, min_sheets_xls, average_sheets_xls, max_sheets_xlsx_file, min_sheets_xlsx_file, max_sheets_xls_file, min_sheets_xls_file,data = analyze_excel_sheets(temp_directory)


                            

                            # Display the DataFrame as a table
                            st.table(data)

                            if duplicated_files:
                                st.write(f"{len(duplicated_files)} duplicated files found")

                                st.subheader("Duplicated Files:")
                                for key, file_paths in duplicated_files.items():
                                    st.write(f"File Information: {key}")
                                    st.write(f"File Paths: {file_paths}")
                            else:
                                st.write("No duplicated files found.")

                            # Check for corrupted files
                            corrupted_files_count, corrupted_files_info = is_file_corrupted(temp_directory)

                            if corrupted_files_count > 0:
                                st.write(f"{corrupted_files_count} corrupted files found")

                                st.subheader("Corrupted Files:")
                                for file_info, file_paths in corrupted_files_info.items():
                                    st.write(f"File Information: {file_info}")
                                    st.write(f"File Paths: {file_paths}")
                            else:
                                st.write("No corrupted files found.")                          
                            
                        

                            time.sleep(10)

                            
                            
                            # Remove temporary directory and files
                            for file_name in os.listdir(temp_directory):
                                file_path = os.path.join(temp_directory, file_name)
                                
                                try:
                                    os.remove(file_path)
                                except PermissionError:
                                    print(f"PermissionError: Could not remove {file_path}. It might be in use by another process.")

                            # Now, remove the empty directory
                            try:
                                os.rmdir(temp_directory)
                            except OSError:
                                print(f"OSError: Could not remove {temp_directory}. It might not be empty.")

                            

                            


                            # Add a button to trigger data_ingestion_DDR.py
                            # output_directory = st.text_input("Enter output path", value="C:\\Users\\User\\Downloads\\")
        


            

        with col22:

            # Create a DataFrame for Excel Sheet Statistics
            df_excel_stats = pd.DataFrame({
                "Metric": ["Total Sheets (.xlsx)", "Max Sheets (.xlsx)", "Min Sheets (.xlsx)", "Average Sheets (.xlsx)"],
                "Value": [int(total_sheets_xlsx), f"{max_sheets_xlsx} (File: {max_sheets_xlsx_file})", f"{min_sheets_xlsx} (File: {min_sheets_xlsx_file})", average_sheets_xlsx]
            })

            st.subheader("File Stats", divider=True)
            st.write("Overall file statistics")
            st.table(df_excel_stats)
            
            
            

            



            

        # Divide into two columns
        col33, col44 = st.columns(2, gap="medium")

        

        
        with col33:
            
            st.subheader("General Well Info",divider=True)
            st.write("Note: This parser is limited per processing per well only. Key in a single well detail at a time")
            wguid = st.text_input("Well GUID")
            wname = st.text_input("Well Name")
            wbname = st.text_input("Wellbore Name")

            st.write("Select your expected .xlsx output template")

            selected_option = st.radio("Select Extraction Template", ["DrillPlan","Performance Insight"])

        with col44:

            st.subheader("Coordinate Check",divider=True)
            st.write("Pre-check of Well Coordinate displayed on map")
            cood1 = st.text_input("Latitude N/S ('-' for S)" )
            cood2 = st.text_input("Longitude W/E ('-' for W)")

                     

            if cood1 and cood2:

                # center on Liberty Bell, add marker
                m = folium.Map(location=[cood1, cood2], zoom_start=5)
                folium.CircleMarker([cood1, cood2], popup=wname, tooltip=wname
                ).add_to(m)

                # call to render Folium map in Streamlit
                st_data = st_folium(m, width=1000)
            
            

        
                        
    

        
        st.subheader("Parser Configuration",divider=True)
        st.write("Configure your DDR attributes and indexes to aid the extraction pipeline. ")

        # Divide into two columns
        col10, col20 = st.columns((1,1))

        with col10:
            st.markdown("""
            **Definitions**
            
            - **TimeStart Marker:** Keyword of Start Time Column
            - **TimeStart Index:** Start Time Column number
            - **TimeEnd Marker:** Keyword of End Time Column
            - **TimeEnd Index:** End Time Column Number
            - **NPT Code Marker:** Keyword of NPT Code Column (if any)
            - **NPT Code Index:** NPT Code Column Number (if any)

            """)

        with col20:
            st.markdown("""
            **Definitions**
            
            - **ActivityStart Marker:** Keyword of Activity Header
            - **ActivityEnd Marker:** Keyword (end-limit) of Activity Column
            - **Activity Index:** Activity Column Number
            - **Date Marker:** Header box for Date Cell (if any)
            - **Activity Code Marker:** Keyword of Activity Code Column
            - **Acitivity Code Index:** Activity Code Column Number

            """)

        

        # Divide into two columns
        col1, col2, col3 = st.columns((5,1,1))

        with col1:
            # Insert an image below the parser configuration
            st.image(r"config_example.png", use_column_width=True)
            st.image(r"config_example2.png", use_column_width=True)

        # First Column
        with col2:
            time_start_marker = st.text_input("TimeStart Marker")
            time_start_index = st.text_input("TimeStart Index")
            time_end_marker = st.text_input("TimeEnd Marker")
            time_end_index = st.text_input("TimeEnd Index")
            NPT_code_marker = st.text_input("NPT Code Marker")
            NPT_code_index = st.text_input("NPT Code Index")


        # Second Column
        with col3:
            activity_start_marker = st.text_input("ActivityStart Marker")
            activity_end_marker = st.text_input("ActivityEnd Marker")
            activity_index = st.text_input("Activity Index")
            date_marker = st.text_input("Date Marker")
            activity_code_marker = st.text_input("Activity Code Marker")
            activity_code_index = st.text_input("Activity Code Index")

        if st.button("Start Parse"):
            
            start_time = time.time()  # Record the start time

            # Display a progress bar
            progress_bar = st.progress(0)

            # Call data_ingestion_DDR.py using subprocess
            # subprocess.run(["python", "data_ingestion_DDR.py", temp_directory, wguid, wname, wbname, time_start_marker,time_start_index,time_end_marker,time_end_index,activity_start_marker,activity_end_marker,date_marker,activity_index,
            #                 NPT_code_marker,NPT_code_index,activity_code_marker,activity_code_index])
            
            subprocess.run(["python", "data_ingestion_DDR.py", temp_directory, wguid, wname, wbname, time_start_marker,time_start_index,time_end_marker,time_end_index,
                            activity_start_marker,activity_end_marker,date_marker,activity_index])

            # Update progress bar
            for percent_complete in range(0, 101, 10):
                time.sleep(0.1)
                progress_bar.progress(percent_complete)

            # Inform the user that processing is complete
            # st.success("Processing complete! View your extracted file in temp_files/output/")

            if os.path.exists(temp_directory + "/output/DDR_extraction_output.xlsx"):
                st.success("Parsing Completed! Excel file exists at temp_files/output/")
            else:
                st.error("Parsing Error : Please Troubleshoot")

            end_time = time.time()  # Record the end time
            elapsed_time = end_time - start_time  # Calculate the elapsed time in seconds

            # Convert seconds to minutes for better readability
            elapsed_time_minutes = elapsed_time / 60

            # Display the time taken
            st.write(f"Total parsing time approximately: {elapsed_time_minutes:.2f} minutes for a total of {total_sheets_xlsx} excel sheets / DDRs")



   
if __name__ == "__main__":
    main()
