import streamlit as st
import json
import subprocess

# Use the full page instead of a narrow central column
st.set_page_config(layout="wide")

# st.write(st.session_state["shared"])

st.title("DrillParse - PDF")
st.write("Automate and customized pre-processing of Unstructured Data of Daily Drilling Report (DDR) in a general-manner of extraction from various DDR templates.")

# st.caption("Version Note")
st.caption("Assumption of all DDRs templates across multiple sheets and xlsx files are consistent.")
st.divider()

st.subheader("Drilling Report Uploader",divider=True)



# Upload files
uploaded_files = st.file_uploader("Upload PDF files", accept_multiple_files=True)




if uploaded_files:
        
        if st.button("Reset Process"):
            pyautogui.hotkey("ctrl","F5")

        st.divider()

        st.subheader("Parser Configuration",divider=True)
        
        col1, col2, col3 = st.columns(3, gap="medium")

        with col1:
        
            # Start Date
            startmarker_date = st.text_input("Start Date Marker")
            date_finder = st.selectbox("Date Search Method",("CellWithMarker","CellBelowMarker","CellBelowMarkerIgnoreEmpty","FirstNonEmptyCellAfterMarker","CellWithMarkerOrFirstNonEmptyCellAfterMarker","RowsBetweenMarkers"))
            date_format = st.text_input("Date Format")

            # Borehole
            startmarker_borehole = st.text_input("Borehole Marker")
            borehole_finder = st.text_input("Borehole Finder",value="CellWithMarker")

        with col2:

            # DDR Table
            startmarker_table = st.text_input("DDR Table Start Marker")
            endmarker_table = st.text_input("DDR Table End Marker")
            table_finder = st.text_input("Table Finder",value="RowsBetweenMarkers")



            ### Table columns detail

            # time column
            startmarker_time = st.text_input("Start Time Marker")
            time_regex = st.text_input("Start Time Regular Expression (if required)")
            time_format = st.text_input("Time Format")

        with col3:

            # time column
            startmarker_duration = st.text_input("Duration Marker")
            duration_sep = st.text_input("Duration Seperator")

        

            # Phase column
            startmarker_phase = st.text_input("Phase Marker")

            # Task column
            startmarker_task = st.text_input("Task Marker")

            # Activity column
            startmarker_activity = st.text_input("Activity Marker")

            # Code column
            startmarker_code = st.text_input("Code Marker")

            # Multi-line comments column
            startmarker_comment = st.text_input("Comment Marker")

        st.divider()



        if st.button("Parse PDF"):
            # Create a dictionary with user inputs
            json_data = {
                            "InputBlocks": [
                                {
                                "Type": "Date",
                                "Name": "ReportDate",
                                "StartMarker": startmarker_date,
                                "SearchMethod": date_finder,
                                "Formats": date_format
                                },
                                {
                                "Type": "Text",
                                "Name": "BoreholeName",
                                "StartMarker": startmarker_borehole,
                                "SearchMethod": borehole_finder
                                },
                                {
                                "Type": "Table",
                                "Name": "DDR",
                                "StartMarker": startmarker_table,
                                "EndMarker": endmarker_table,
                                "SearchMethod": table_finder,

                                "TableColumns": [
                                    {
                                    "Name": "StartTime",
                                    "Type": "Time",
                                    "IsKeyColumn": True,
                                    "ReplaceValues": {
                                        time_regex
                                    },
                                    "DefaultValue": "",	
                                    "Formats": time_format,
                                    "StartMarker": startmarker_time,
                                    "AddDayIfZero":True,
                                    "IsObligatory": True
                                    },
                                    {
                                    "Name": "Duration",
                                    "Type": "Hours",
                                    "StartMarker": startmarker_duration,
                                    "DecimalSeparator": duration_sep,
                                    "IsObligatory": True
                                    },
                                    {
                                    "Name": "Phase",
                                    "Type": "Text",
                                    "StartMarker": startmarker_phase,
                                    "IsObligatory": True
                                    },
                                    {
                                    "Name": "Task",
                                    "Type": "Text",
                                    "StartMarker": startmarker_task,
                                    "IsObligatory": True
                                    },
                                    {
                                    "Name": "Activity",
                                    "Type": "Text",
                                    "StartMarker": startmarker_activity,
                                    "IsObligatory": True
                                    },
                                    {
                                    "Name": "Code",
                                    "Type": "Text",
                                    "StartMarker": startmarker_code,
                                    "IsObligatory": True
                                    },
                                    {
                                    "Name": "Comments",
                                    "Type": "Multiline",
                                    "StartMarker": startmarker_comment,
                                    "IsObligatory": True
                                    }
                                ]
                                }
                            ],
                            "OutputTable": {
                                "Columns": [
                                {
                                    "Name": "BoreholeName",
                                    "HeaderText": "Borehole",
                                    "SourceBlocks": ["BoreholeName"]
                                },
                                {
                                    "Name": "DailyCost",
                                    "HeaderText": "Daily Cost",
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "Phase",
                                    "HeaderText": "Phase",
                                    "SourceBlocks": ["DDR.Phase"],
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "BoreholeSection",
                                    "HeaderText": "Borehole Section",
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "StartDate",
                                    "HeaderText": "Start Time",
                                    "SourceBlocks": ["ReportDate", "DDR.StartTime"],
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "EndDate",
                                    "HeaderText": "End Time",
                                    "Format": None,
                                    "SourceBlocks": ["ReportDate", "DDR.StartTime", "DDR.Duration"],
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "Duration",
                                    "HeaderText": "Duration(h)",
                                    "SourceBlocks": ["DDR.Duration"],
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "StartDepth",
                                    "HeaderText": "Start Depth(m)",
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "EndDepth",
                                    "HeaderText": "End Depth(m)",
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "Code",
                                    "HeaderText": "Operation Code",
                                    "Mode": "Cell"
                                },
                                        {
                                    "Name": "ActivityCode",
                                    "HeaderText": "Activity Code",
                                    "SourceBlocks": ["DDR.Activity"],
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "SubCode",
                                    "HeaderText": "Sub Code",
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "Planned",
                                    "HeaderText": "Planned",
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "TimeType",
                                    "HeaderText": "Time Classification",
                                    "Mode": "Cell",
                                    "SourceBlocks": ["DDR.Code"],
                                    "ReplaceValues": {
                                    "^P.*$": "Productive"
                                    },
                                    "DefaultValue": "Non-Productive"		
                                },
                                {
                                    "Name": "NPTReason",
                                    "HeaderText": "NPT Reason",
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "NPTVendor",
                                    "HeaderText": "NPT Vendor",
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "Comments",
                                    "HeaderText": "Comments",
                                    "SourceBlocks": ["DDR.Comments"],
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "TranslatedComments",
                                    "HeaderText": "Translated Comments",
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "Perforation",
                                    "HeaderText": "Perforation",
                                    "Mode": "Cell"
                                },
                                {
                                    "Name": "AdditionalCode",
                                    "HeaderText": "Additional Code",
                                    "Mode": "Cell"
                                }
                                ]
                            }
                            }


            # Convert sets to lists in the JSON data
            def convert_sets_to_lists(obj):
                if isinstance(obj, set):
                    return list(obj)
                elif isinstance(obj, dict):
                    return {key: convert_sets_to_lists(value) for key, value in obj.items()}
                elif isinstance(obj, list):
                    return [convert_sets_to_lists(item) for item in obj]
                else:
                    return obj

            json_data_serializable = convert_sets_to_lists(json_data)

            # Write to JSON file
            with open("BPOman1.json", "w") as json_file:
                json.dump(json_data_serializable, json_file, indent=2)
            
            subprocess.run([r"pdfparser/PdfParser/Parser/BPOman1.bat"])

            st.success("JSON file created: output.json")






        



