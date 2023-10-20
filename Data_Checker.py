import PySimpleGUI as sg
import pandas as pd
#import os
import math
from docx import *
# Program: Data_Checker.py
# Version: 0.5
# Description: This program is used in order to display data onto a GUI inorder for the data to be compared to see if they match or not.
# Functions:
#   (1) To read Excel sheets and Microsoft Documents.
#   (2) Use parsed data in two tables that will be displayed on a GUI.

# Theme
sg.theme('DarkBlack')

# Settings for the window.
window = sg.FlexForm('Data Checker', default_button_element_size = (5,2), auto_size_buttons=False, grab_anywhere=False, resizable=False)

# The layout of the window.
layout = [ # The browser for choosing an excel file to be parsed.
           [sg.Text("Choose an excel file:",
                   size=(20,1),
                   font=('Times New Roman', 12)),
           sg.Input(key="-data_1-",
                    size=(100,1)),
           sg.FileBrowse("Browse",
                         size=(10, 1),
                         tooltip="Choose the desired xlsx file that has the data.")],
           
          # Once an excel is loaded, pick a sheet (technology) that will be parsed.
          [sg.Text("Pick a Technology\nto load from the Excel Sheet:", 
                    size=(20,2), 
                    font=('Times New Roman', 12)),
           sg.Combo(values="", 
                    key="-tech_1-", 
                    size=(25,1)),
           sg.Button("Load",
                     size=(10,1),
                     tooltip="Refreshes technology list.\nHit this after loading a new excel file."),
           sg.Text("",
                   key="-Error_Technology-",
                   size=(20,1),
                   font=('Times New Roman', 12))],
                    
          # The browser for choosing a docx file to be parsed.
          [sg.Text("Choose a docx file:",
                   size=(20,1),
                   font=('Times New Roman', 12)),
           sg.Input(key="-data_2-", 
                    size=(100,1)),
           sg.FileBrowse("Browse",
                         size=(10, 1),
                         tooltip="Choose the desired docx file that has the data.")],
          
          # Input on whether extremity is used or not. (NOTE: Currently this is hidden, because we don't delete rows in the excels normally, so the amount of rows in the excel that is being parsed should not change)
          [sg.Text("Extremity?:",
                   size=(20,1),
                   font=('Times New Roman', 12),
                   visible=False),
           sg.Combo(values=["Yes", "No"],
                    default_value="Yes", 
                    key="-confirm_extremity_1-", 
                    size=(25,1),
                    disabled=True,
                    visible=False)], 
          
        ##  This commented block was made to account for if a position is excluded. However, I forgot that we just hide rows that we will not use.
        ##  So, this commented block could potentially be used in the future.
        #   [sg.Text()],
 
        #   [sg.Text("The below options are used for when a test position is excluded. (I.e. Edge Left is excluded because it is >= 25mm from the current ANT).",
        #            font=('Times New Roman', 12))],
          
        #   [sg.HorizontalSeparator()],
 
        #   [sg.Text("(4) Head or Body and how many positions are excluded?:",
        #            font=('Times New Roman', 12)),
        #    sg.Combo(values=["Head", "Hotspot"],
        #             key="-excluded_exposure_1-",
        #             size=(25,1)),
        #    sg.Input("",
        #             key="-excluded_number_1-",
        #             size=(25,1))],  
          
        #   [sg.HorizontalSeparator()],
          
          [sg.HorizontalSeparator()],
          
          # Table for showing the data from the excel.
          [sg.VerticalSeparator(),
           sg.Text("Excel\nData:",
                   size=(7,2),
                   font=("Times New Roman", 16, "bold"),
                   text_color="green"),
           sg.Table(values="", 
                    headings=["Plot #", "Test Position", "Ch #.", "Freq. (MHz)", "RB Allocation", "RB Offset", "1-g Meas. (W/kg)", "10-g Meas. (W/kg)"],  
                    key="-data_table_1-",
                    justification='center',
                    def_col_width=15,
                    num_rows=18,
                    auto_size_columns=False,
                    enable_events=True)],
          
          [sg.HorizontalSeparator()],
          
          # Table for showing the data from the plots.
          [sg.VerticalSeparator(),
           sg.Text("Plot\nData:",
                   size=(7,2),
                   font=("Times New Roman", 16, "bold"),
                   text_color="blue"),
           sg.Table(values="", 
                    headings=["Plot #", "Test Position", "Ch #.", "Freq. (MHz)", "RB Allocation", "RB Offset", "1-g Meas. (W/kg)", "10-g Meas. (W/kg)"],  
                    key="-data_table_2-",
                    justification='center',
                    def_col_width=15,
                    num_rows=18,
                    auto_size_columns=False,
                    enable_events=True)],          
          
          [sg.HorizontalSeparator()],
          
          # The buttons for: loading an excel, loading a docx, and compare results window.
          [sg.Button("Load Excel",
                     size=(10,1),
                     font=("Times New Roman", 12, "bold"),
                     button_color="white"),
           sg.Button("Load Docx",
                     size=(10,1),
                     font=("Times New Roman", 12, "bold"),
                     button_color="white"),
           sg.Button("Compare",
                     key="-compare-",
                     size=(10,1),
                     font=("Times New Roman", 12, "bold"),
                     button_color="white"),
           sg.Text("",
                   key="-FuckedUp-",
                   size=(20,1),
                   font=('Times New Roman', 12, "bold"))]  
]

window.Layout(layout)   # Display the window.
xl = path = tech = ""
data_excel = []
data_plot = []
skip_rows, num_rows = 0, 0

rb_positions_lte = {
    "1400000": {
        "1 RB": {
            "Low": "0",
            "Mid": "3",
            "High": "5"
        },
        "50% RB": {
            "Low": "0",
            "Mid": "1",
            "High": "3"
        }
    },
    "3000000": {
        "1 RB": {
            "Low": "0",
            "Mid": "8",
            "High": "14"
        },
        "50% RB": {
            "Low": "0",
            "Mid": "4",
            "High": "7"
        }
    },
    "5000000": {
        "1 RB": {
            "Low": "0",
            "Mid": "12",
            "High": "24"
        },
        "50% RB": {
            "Low": "0",
            "Mid": "7",
            "High": "13"
        }
    },
    "10000000": {
        "1 RB": {
            "Low": "0",
            "Mid": "25",
            "High": "49"
        },
        "50% RB": {
            "Low": "0",
            "Mid": "12",
            "High": "25"
        }
    },
    "15000000": {
        "1 RB": {
            "Low": "0",
            "Mid": "37",
            "High": "74"
        },
        "50% RB": {
            "Low": "0",
            "Mid": "20",
            "High": "39"
        }
    },
    "20000000": {
        "1 RB": {
            "Low": "0",
            "Mid": "49",
            "High": "99"
        },
        "50% RB": {
            "Low": "0",
            "Mid": "24",
            "High": "50"
        }
    }
}

rb_positions_nr = {
    "1400000": {
        "1 RB": {
            "Low": "0",
            "Mid": "3",
            "High": "5"
        },
        "50% RB": {
            "Low": "0",
            "Mid": "1",
            "High": "3"
        }
    },
    "3000000": {
        "1 RB": {
            "Low": "0",
            "Mid": "8",
            "High": "14"
        },
        "50% RB": {
            "Low": "0",
            "Mid": "4",
            "High": "7"
        }
    },
    "5000000": {
        "1 RB": {
            "Low": "0",
            "Mid": "12",
            "High": "24"
        },
        "50% RB": {
            "Low": "0",
            "Mid": "7",
            "High": "13"
        }
    },
    "10000000": {
        "1 RB": {
            "Low": "0",
            "Mid": "25",
            "High": "49"
        },
        "50% RB": {
            "Low": "0",
            "Mid": "12",
            "High": "25"
        }
    },
    "15000000": {
        "1 RB": {
            "Low": "0",
            "Mid": "37",
            "High": "74"
        },
        "50% RB": {
            "Low": "0",
            "Mid": "20",
            "High": "39"
        }
    },
    "20000000": {
        "1 RB": {
            "Low": "0",
            "Mid": "49",
            "High": "99"
        },
        "50% RB": {
            "Low": "0",
            "Mid": "24",
            "High": "50"
        }
    }
}

while True:
    event, values = window.read()
    
    # Break out of loop which closes the window.
    if event == sg.WIN_CLOSED:
        break
    
    window["-FuckedUp-"].update("")
    path_excel = values["-data_1-"]                         # This is the path used to get the excel.
    path_docx = values["-data_2-"]                          # This is the path used to get the microsoft document.
    tech = values["-tech_1-"]                               # This is the technology/band that will be selected.
    extremity_confirm = values["-confirm_extremity_1-"]     # This is a confirmation for whether you need extremity or not.
    # excluded_exposure =  values["-excluded_exposure_1-"]  # This is the selection of body exposure condition. (Only Head and Body).
    # excluded_number = values["-excluded_number_1-"]       # This is the number of excluded positions due to distance being to far from the antenna.
        
    print("Event:", event)
    print("Technology:", tech)
    # When the "Load" button is pressed:
    # (1) Sheet names will be taken from the xlsx file.
    # (2) Remove unnecessary sheets from the sheet names list.
    # (3) Update the combo element that holds the sheet names.
    if event == "Load" and path_excel != "":
        
        # Clear window.
        window["-Error_Technology-"].update("")
        window["-FuckedUp-"].update("")
        
        try:
            # Put sheet names from Excel in list.
            xl = pd.ExcelFile(path_excel).sheet_names
        except:
            window["-Error_Technology-"].update("Please use an Excel file!")
            window["-FuckedUp-"].update("Dx")
            
        for sheet_name in ['How to Use this Workbook', 'Inter-Band CA Exclusion', 'Settings', 'Sum of SAR',
                           'Repeated', 'Data', 'Master List (WWAN)', 'List Variables', 'ISED Extremity']:
            if sheet_name in xl:
                xl.remove(sheet_name)

        window["-tech_1-"].update(values = xl)
    # When the "Insert Data" button is pressed:
    # (1) Choose the correct section for the selected technology.
    # (2) Once (1) is chosen, the columns, skipped rows, and number rows used for reading the excel.
    # (3) Based on the selected excel sheet and previous parameters from (1) & (2), a dataframe is created.
    # (4) This is then filtered for plot number, test position, channel, frequency, 1-g measured, and 10-g measured.
    # (5) Update the table element that holds all the filtered data from (4).
    elif event == "Load Excel" and tech in xl and tech != "":
        
        ref_df = pd.read_excel(path_excel, sheet_name=tech, index_col=None, na_values=['N/A'])
        ref_data = ref_df.values.tolist()
        
        # Get the parameters on where the parse the data in the excel.
        count, skip_rows, num_rows = 0, 0, 0
        for ref_row in ref_data:
            if ref_row[0] == "System Check Date" and skip_rows == 0:
                print("Found")
                skip_rows = count + 1
                count = 0
            elif ref_row[0] == "Repeated" and skip_rows != 0:
                print("Founded")
                num_rows = count - 3
            print(skip_rows, count)
            count += 1
        print("Skip Rows: {}".format(skip_rows), "Num rows: {}".format(num_rows))
        
        del ref_df
        del ref_data
        
        # 3 Channels for GSM and W-CDMA or 1 channel for FR1 and LTE.
        if any(technology in tech for technology in ["GSM", "PCS", "W-CDMA"]):
            cols = "M, U:V, Y, AA"  # Sets the columns in excel that are parsed.
        elif "LTE" in tech or "FR1" in tech:
            cols = "M, W:Z, AC, AE" # Sets the columns in excel that are parsed.
                
        df = pd.read_excel(path_excel, sheet_name=tech, index_col=None, na_values=['N/A'], usecols="{}".format(cols), skiprows=skip_rows, nrows=num_rows) # Create a dataframe from the excel on the selected rows and columns.
        data = df.values.tolist()   # Insert all the values of the dataframe into a list.        

        len_data = len(data)
        for position in range(0, len_data):
            # Insert index column
            # data[position].insert(0, position+1)

            merge_index = 0
            frequency_num = 1
            channel_num = 2
            if any(technology in tech for technology in ["LTE", "FR1"]):
                num_rb, offset_rb, meas_1g, meas_10g = 3, 4, 5, 6
                ch_nrb_orb = [channel_num, frequency_num, num_rb, offset_rb] # Holds Ch #, Freq, Number RB, and Offset RB.
                meas_values = [meas_1g, meas_10g] # Holds 1-g/10-g measured.
            elif any(technology in tech for technology in ["GSM", "PCS", "W-CDMA"]):
                meas_1g, meas_10g = 3, 4
                ch_nrb_orb = [channel_num, frequency_num] # Holds Ch # and Freq.
                meas_values = [3, 4] # Holds 1-g/10-g measured.
                
            for sublist_number in range(0, len(data[position])):                
                # Because of merged cells in the xlsx, the cells that are not at the top
                # of the merged cell are considered "NaN". This solves that by replacing
                # the index with a "NaN" with the previous index, which is the test position.
                if sublist_number == merge_index and pd.isna(data[position][sublist_number]):
                    data[position][sublist_number] = data[position-1][sublist_number]
                # Convert the "Ch. #" to an integer and add ""
                elif sublist_number in ch_nrb_orb:
                    data[position][sublist_number] = "{}".format(data[position][sublist_number]) if pd.isna(data[position][sublist_number]) else "{}".format(float(data[position][sublist_number]) if sublist_number == 2 else int(data[position][sublist_number]))
                # Round the "1-g Meas. (W/kg)" and "10-g Meas. (W/kg)" columns.
                elif sublist_number in meas_values:
                    data[position][sublist_number] = "{:.3f}".format(round(data[position][sublist_number], 3))
                    
            # If not LTE. Resource Blocks (RBs) are not used for the other technologies.
            if "LTE" not in tech and "FR1" not in tech:
                data[position].insert(3, "N/A")
                data[position].insert(4, "N/A")

        # This section detects if the current sublist has no number in the 1-g Meas or 10-g Meas (Or 'nan').
        len_data = len(data)
        plot_number_tracker = 0 # Used to update the index after a pop has happened.
        for index in range(0, len_data):
            index += plot_number_tracker
            if data[index][5] == "nan" or data[index][6] == "nan":  # Reason I didn't use pd.isna() is because the 1-g and 10-g are added as strings instead of None.
                data.pop(index) # Remove current index that contains 'nan' on 1-g Meas or 10-g Meas.
                len_data -= 1
                plot_number_tracker -= 1
            else:
                data[index].insert(0, index+1)
        
        df_excel = data       
        data_excel = df_excel # Used for comparison purposes.
        
        window["-data_table_1-"].update(values = data)
        
    elif event == "Load Docx" and path_docx != "":

        window["-Error_Technology-"].update("")
        window["-FuckedUp-"].update("") 

        try:
            docx = Document(path_docx)  # Open docx.
        except:
            window["-Error_Technology-"].update("Please use a Docx file!")
            window["-FuckedUp-"].update("Dx")            
        docx_paragraphs = [para.text for para in docx.paragraphs if ((para.text).strip() not in ["", "\n"])] # Remove random whitespace and newlines from the paragraph list. # Load all paragraphs in the docx.
        docx_tables = docx.tables   # Load all tables in the docx.
        table_1 = []    # Initialize the list that will hold all the data.                   
                        
        start = 0 # Determines the table that is having its data parsed.
        sublist_start = 0 # Determines which index of the list holding the data (table_1) to extend it with more data.
        while start < 4:
            for table_num in range(start, len(docx_tables), 4):
                if start == 2: # 'start = 2' is the "Scan Setup" table on the plot. (Unused because it mostly does not have any useful info to parse).
                    break
                elif start == 0: # 'start = 0' is the "Exposure Conditions" table on the plot.
                    split_freqch = (docx_tables[table_num].rows[1].cells[1].text).split()
                    frequency, channel = '{:.1f}'.format(float(split_freqch[0])), split_freqch[2]
                    #rpermittivity = docx_tables[table_num].rows[0].cells[3].text
                    #conductivity = docx_tables[table_num].rows[1].cells[3].text
                    
                    table_1.append([channel,    # Get channels.
                                    frequency]) # Get frequencies.
                                                # NOTE: IF YOU WANT THE RELATIVE PERMITTIVITY AND CONDUCTIVITY, ADD 'rpermittivity' AND 'conductivity' HERE AND UNCOMMENT!
                # NOTE: IF YOU WANT TO ADD THE HARDWARE (DAE/PROBE) UNCOMMENT THIS SECTION!
                # elif start == 1: # 'start = 1' is the "Hardware Setup" table on the plot.
                #     split_probedate = (docx_tables[table_num].rows[0].cells[1].text).split()
                #     split_daedate = (docx_tables[table_num].rows[1].cells[1].text).split()
                #     probe_sn, probe_caldate = split_probedate[2], split_probedate[4]
                #     dae_sn, dae_caldate = split_daedate[1], split_daedate[3]
                    
                #     # Extend current sublist of 'table_1' with probe sn/calibration date and DAE sn/calibration date.
                #     if sublist_start < len(table_1):
                #         table_1[sublist_start].extend([probe_sn,        # Get probe sn.
                #                                        probe_caldate,   # Get probe calibration due date.
                #                                        dae_sn,          # Get dae sn.
                #                                        dae_caldate])    # Get dae calibration due date.
                #     sublist_start += 1
                elif start == 3: # 'start = 3' is the "Measurement Results" table on the plot.
                    if len(docx_tables[table_num].rows[0].cells) == 2:
                        zoom_meas_1g = "0.000"
                        zoom_meas_10g = "0.000"
                    elif len(docx_tables[table_num].rows[0].cells) == 3:
                        zoom_meas_1g = docx_tables[table_num].rows[1].cells[2].text    # Get zoom scans measured 1-g (W/kg).
                        zoom_meas_10g = docx_tables[table_num].rows[2].cells[2].text   # Get zoom scans measured 10-g (W/kg).
                        #power_drift = docx_tables[table_num].rows[3].cells[2].text     # Get power drift (dB).
                    elif len(docx_tables[table_num].rows[0].cells) == 4:
                        first_zoom_meas_1g = docx_tables[table_num].rows[1].cells[2].text    # Get first zoom scans measured 1-g (W/kg).
                        first_zoom_meas_10g = docx_tables[table_num].rows[2].cells[2].text   # Get first zoom scans measured 10-g (W/kg).
                        second_zoom_meas_1g = docx_tables[table_num].rows[1].cells[3].text   # Get second zoom scans measured 1-g (W/kg).
                        second_zoom_meas_10g = docx_tables[table_num].rows[2].cells[3].text  # Get second zoom scans measured 10-g (W/kg).
                        #first_power_drift = docx_tables[table_num].rows[3].cells[2].text     # Get first power drift.
                        #second_power_drift = docx_tables[table_num].rows[3].cells[3].text    # Get second power drift.
                        
                        zoom_meas_1g = first_zoom_meas_1g if first_zoom_meas_1g > second_zoom_meas_1g else second_zoom_meas_1g      # Determines which 1-g measured zoom scan to use.
                        zoom_meas_10g = first_zoom_meas_10g if first_zoom_meas_10g > second_zoom_meas_10g else second_zoom_meas_10g # Determines which 10-g measured zoom scan to use.
                        #power_drift = first_power_drift if first_zoom_meas_1g > second_zoom_meas_1g else second_power_drift         # Determines which power drift to use.
                        
                    # Extend current sublist of 'table_1' with 1-g measured, 10-g measured, and power drift.
                    if sublist_start < len(table_1):
                        table_1[sublist_start].extend(["{:.3f}".format(round(float(zoom_meas_1g), 3)),
                                                       "{:.3f}".format(round(float(zoom_meas_10g), 3))]) # NOTE: IF YOU WANT TO ADD POWER DRIFT, ADD 'power_drift' HERE.
                    sublist_start += 1

            sublist_start = 0
            start += 1
               
        plot_data = []
        start = 0
        index = 0
        plot_num = 0
        for paragraph_num in range(0, len(docx_paragraphs)):
            para_index = docx_paragraphs[paragraph_num] # Holds the current string in the paragraph.
            
            if paragraph_num % 6 == 0 or paragraph_num == 0:
                plot_num += 1
                plot_data.append([plot_num])
            
            # Logic to get position.
            split_head = (docx_tables[start].rows[2].cells[3].text).split()
            if any([pos in para_index for pos in ["CHEEK", "TILT", "BACK", "FRONT", "EDGE TOP", "EDGE RIGHT", "EDGE BOTTOM", "EDGE LEFT"]]):
                for pos in ["CHEEK", "TILT", "BACK", "FRONT", "EDGE TOP", "EDGE RIGHT", "EDGE BOTTOM", "EDGE LEFT"]:
                    position = para_index[para_index.find("{}".format(pos)):].strip()
                    if position in ["CHEEK", "TILT"]:
                        side = "Left" if "LeftHead" in split_head else "Right"
                        cheek_tilt = position[0] + position[1:].lower()
                        position = side + " " + cheek_tilt
                        break
                    elif position in ["BACK", "FRONT", "EDGE TOP", "EDGE RIGHT", "EDGE BOTTOM", "EDGE LEFT"]:
                        if position in ["BACK", "FRONT"]:
                            position = position[0] + position[1:].lower()
                        elif position in ["EDGE TOP", "EDGE RIGHT", "EDGE BOTTOM", "EDGE LEFT"]:
                            edge_split = position.split()    # Split 'EDGE' and side into list.
                            edge = edge_split[0][0] + edge_split[0][1:].lower() # Logic to get 'EDGE' and lowercase everything after the first character.
                            side = edge_split[1][0] + edge_split[1][1:].lower() # Logic to get the side and lowercase everything after the first character.
                            position = ' '.join([edge, side])
                        break
                plot_data[index].extend([position])
                if start < len(docx_tables)-4:
                    start += 4

            # Logic to get SAR lab and date/time tested.
            if "SAR Lab" and "Date/Time:" in para_index:
                date = para_index[para_index.find("Date/Time:")+11:para_index.find("Date/Time:")+21].strip()
                lab = para_index[para_index.find("SAR Lab"):para_index.find("SAR Lab")+11].strip()
                        
            # Logic to get GSM/PCS and WCDMA.
            if ("GSM" in para_index or "PCS" in para_index) or "WCDMA" in para_index or "WiFi" in para_index:
                gsm_pcs_tech = "GSM" if para_index.find("GSM") != -1 else "PCS"
                if "GSM" in para_index:
                    find_colon_gsm_850_900 = para_index[para_index.find("GSM")+7]
                    if find_colon_gsm_850_900 == ":":
                        technology = para_index[para_index.find("GSM"):para_index.find("GSM")+7]
                    else:
                        technology = para_index[para_index.find("GSM"):para_index.find("GSM")+8]
                elif "PCS" in para_index:
                    find_colon_gsm_850_900 = para_index[para_index.find("PCS")+7]
                    if find_colon_gsm_850_900 == ":":
                        technology = para_index[para_index.find("PCS"):para_index.find("PCS")+7]
                    else:
                        technology = para_index[para_index.find("PCS"):para_index.find("PCS")+8]
                elif "WCDMA" in para_index:
                    wcdma, band = para_index[para_index.find("WCDMA"):para_index.find("WCDMA")+5], para_index[para_index.find("Band"):para_index.find("Band")+6]
                    technology =  wcdma + " " + band
                # GSM / WCDMA / Wi-Fi don't have RBs
                table_1[index].insert(2, "N/A") # Fill table with 'N/A' for RB Allocation.
                table_1[index].insert(3, "N/A") # Fill table with 'N/A' for RB Offset.
            # Logic to get LTE.
            elif "LTE" in para_index or "5G NR" in para_index:
                if "LTE" in para_index:
                    lte = para_index[para_index.find("LTE"):para_index.find("LTE")+3]
                else:
                    fr1 = "FR1"
                    
                band = para_index[para_index.find("Band"):para_index.find("Band")+7] if para_index[para_index.find("Band")+7].isdigit() else para_index[para_index.find("Band"):para_index.find("Band")+6]
                technology = lte + " " + band if "LTE" in para_index else fr1 + " " + band
                #plot_data[index].extend([technology]) # LTE in plot_data.

                if para_index.find("Low") != -1:
                    rb_position = para_index[para_index.find("Low"):para_index.find("Low")+3]
                elif para_index.find("Mid") != -1:
                    rb_position = para_index[para_index.find("Mid"):para_index.find("Mid")+3]
                elif para_index.find("High") != -1:
                    rb_position = para_index[para_index.find("High"):para_index.find("High")+4]
                    
                if para_index.find("1 RB") != -1:
                    num_rb = para_index[para_index.find("1 RB"):para_index.find("1 RB")+4]
                elif para_index.find("50% RB") != -1:
                    num_rb = para_index[para_index.find("50% RB"):para_index.find("50% RB")+6]
                elif para_index.find("100% RB") != -1:
                    num_rb = para_index[para_index.find("100% RB"):para_index.find("100% RB")+7]
                
                # This section is the logic to get the NRB (Number of Resource Blocks) for LTE
                check_half_rb = para_index.find("50%") != -1     # Check if the current plot is for 50% RB
                check_full_rb = para_index.find("100%") != -1    # Check if the current plot is for 100% RB
                find_bw = str(para_index[para_index.find("RB,")+3:para_index.find("MHz")-1]).strip() # Get the bandwidth number.
                    
                if "{} MHz".format(str(find_bw)) in para_index and ("LTE" in para_index):
                    # When the current bandwidth for the band is less than 3 MHz.
                    bw_hz = float(find_bw) * pow(10, 6)                                              # Bandwidth in Hz
                    if bw_hz / pow(10,6) < 3:
                        size_guard_band = bw_hz * 0.001                                              # Guardband = 10% of BW (In Hz)
                        single_slot = 12 * 15                                                        # 12 subcarriers * 15 kHz subcarrier spacing = 180 kHz size of 1 slot.
                        nRB_ref = math.floor((size_guard_band)/(single_slot))                        # Reference NRB which is the BW / slot size. I.e. for 1.4 MHz 1400 kHz / 180 kHz = floor(7.77) = 7.
                        guard_band = (size_guard_band - ((nRB_ref * single_slot) - single_slot)) / 2 # Calculate guardband in kHz.
                        used_bw = size_guard_band - (guard_band * 2)                                 # Calculate the usable bandwidth, (BW w/guard - (guard band size * 2)) = usable BW (kHz)
                        full_rb = math.floor(used_bw / single_slot)                                  # Calculate the full number of RBs that users can allocate.
                    # When the current bandwidth is in [3, 5, 10, 15, 20] (MHz)
                    else:
                        size_guard_band = bw_hz * 0.1
                        single_slot = 12 * 15 * pow(10, 3)
                        used_bw = bw_hz - size_guard_band
                        full_rb = math.floor(used_bw / single_slot)

                    mhz = pow(10,6) # 10^6 for MHz convertion.
                    
                    # Get the RB allocation and offset. This is dependent on the RB position and what percentage of the RB is being allocated.
                    if check_half_rb:
                        rb_allocation = str(math.floor(full_rb / 2)) # NOTE: PLACEHOLDER FORMULA
                        if any(position in para_index for position in ["RBPosition:Low", "RBPosition:Mid"]):
                            rb_offset = rb_positions_lte[str(int(bw_hz))][num_rb][rb_position]
                        else:
                            rb_offset = str(full_rb - 1)
                    elif check_full_rb:
                        rb_allocation = str(full_rb)
                        rb_offset = "0"
                    else:
                        rb_allocation = "1"
                        if any(position in para_index for position in ["RBPosition:Low", "RBPosition:Mid"]):
                            rb_offset = rb_positions_lte[str(int(bw_hz))][num_rb][rb_position]
                        else:
                            rb_offset = str(full_rb - 1)
                    table_1[index].insert(2, rb_allocation) # Insert number of allocated RBs.
                    table_1[index].insert(3, rb_offset)     # Insert offset for RBs.
                    print(rb_allocation, rb_offset, full_rb)

            if len(plot_data[index]) == 2 and paragraph_num % 6 == 0:
                index += 1
        
        df_plot = pd.DataFrame(table_1)    
        df_plot = df_plot.values.tolist()          
        
        for index in range(0, len(df_plot)):
            df_plot[index] = plot_data[index] + df_plot[index]
                
        data_plot = df_plot # Used for comparison purposes.
        
        window["-data_table_2-"].update(values = df_plot)
    elif event == "-compare-":
        sg.theme('DarkBlack')
                
        layout_match = [[sg.Table(values="", 
                                    headings=["Plot #", "Match?"],  
                                    key="-data_table_3-",
                                    justification='center',
                                    def_col_width=15,
                                    num_rows=18,
                                    auto_size_columns=False,
                                    enable_events=True)],
                        [sg.Button("Prepare to Be Sad",
                                   size=(30,1),
                                   font=("Times New Roman", 12, "bold"))]]
        window_match = sg.Window("Sadness?", layout_match, auto_size_buttons=True, auto_size_text=True, modal=True)
        while True:
            event_match, values_match = window_match.read()

            if event_match == None:
                break
            elif event_match == "Prepare to Be Sad":
                match_xlsx_docx = []
                for data_1, data_2, index in zip(data_excel, data_plot, range(0, len(data_excel))):
                    print(data_1)
                    print(data_2)
                    match_xlsx_docx.append(["{}".format(index + 1), "Yes" if data_1 == data_2 else "No"])
                window_match["-data_table_3-"].update(values = match_xlsx_docx)
    else:
        if event == "Load" or (event == "Load Excel" and path == ""):
            error = "Please load an excel file."
        elif event == "Load Excel" and tech == "":
            error = "Please load a technology."
        elif event == "Load Docx" and path_docx == "":
            error = "Please load a docx file."
        else:
            error = ""
        window["-FuckedUp-"].update(error)
window.close()