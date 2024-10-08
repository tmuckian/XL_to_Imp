##############################################################################
# REQUIRED MODULES
##############################################################################
import os
import csv
import pandas as pd


##############################################################################
# MODULE DOCUMENTATION
############################################################################### Example data structure:
#
# Converts a formatted Excel file and generates an Ovation import file from it.
#
##############################################################################
# FUNCTIONS
##############################################################################
def defaults(default_file):
    """Creates dictionarys for the different data types
    Parameters
    -----------
    input_file: str
        The Excel file to read
    output_file: str
        The Ovation import file (.imp) to write
    """
    s_type = None
    if os.path.isfile(default_file):
        with open(default_file, mode = 'r') as file:
            data = csv.reader(file)
            for row in data:
                if "Type" in row[0]:
                    s_type = row[1]
                    datatypes[row[1]] = {}
                else:
                    #if len(datatypes[s_type]) == 1:
                    #   print("new data type")
                    if row[0] != '' and row[0] != 'Field':
                        datatypes[s_type][row[0]] = [row[1], row[2], ]
    else:
        print("Failed to find %s" % default_file)

def run(input_file, output_file):
    """Create an Ovation Import file from a formatted Excel Sheet
    Parameters
    -----------
    input_file: str
        The Excel file to read
    output_file: str
        The Ovation import file (.imp) to write
    """
    if os.path.isfile(input_file):
        xl = pd.ExcelFile(input_file) #read the excel files
        xl.sheet_names #get the sheet names that are the data types

        s_type = None #initialize datatype string
        
        for name in xl.sheet_names:  #cycle through the different sheets of the Excel File
            if name in datatypes.keys(): #check to see if this datatype is in the defaults file
                s_type = name
                sheet_data = pd.read_excel(xl, name)
                data = sheet_data.to_numpy()
                for rows in data:
                    p_name = rows[1]
                    dict_output[p_name]={'POINT_NAME': p_name}
                    col = 0            
                    for columns in sheet_data:
                        if pd.isna(rows[col]):
                            dict_output[p_name][columns] = ''
                        else:
                            dict_output[p_name][columns] = rows[col]
                        col += 1
            else:  #Error message if the datatype is missing from the default file
                print("Missing Data Type: %s from Defaults File" %name)        
    else:
        print("Failed to find %s" % input_file)
        # Fingers crossed that crashing doesn't corrupt

    #Runs the clean up program
    clean_up() 
    #creates the output file
    create_output(output_file) 

def clean_up():
    """Compares the dict_output generated previously against the defaults and deletes unused parameters, checks for required paraments, fills in defaults
    Parameters
    -----------
    none <uses the global variables>
    """
    for p_name in dict_output: #loop through the point names in the output dictionary
        s_type = dict_output[p_name]['RECORD_TYPE'] #find the record type of the current point name
        for s_field in datatypes[s_type]: #loop through all of the default fields for this point type
            [s_req, s_default] = datatypes[s_type][s_field] #get the value of whether the field is required and the default value
                        
            s_value = dict_output[p_name].get(s_field, '[unknown]') #gets the value for the field in the output file

            #checks to see if any required fields are missing
            if s_value =='' and s_req == 'x':
                global error_msg
                error_msg = "Field: " + str(s_field) + " is required for point: " + str(p_name) + " on sheet " + str(s_type)
                print(error_msg)
                return
            #fills in default data if it's not found and it's configured in the defaults
            if s_value == '[unknown]':
                if s_default != '':
                    #print('added default ' + str(p_name) + ' ' + str(s_field) + ' '  + str(s_default))
                    dict_output[p_name][s_field] = s_default

def create_output(output_file):
    """creates the output file using the output dictionary
    Parameters
    -----------
    output_file: file name for the output
    """    
    o_file = open(output_file, mode='w')
    if len(error_msg) > 1:
        o_file.write(error_msg)
    else:
        for p_name in dict_output:
            o_string = 'OBJECT="POINT" ACTION="INSERT" POINT_NAME="' + p_name + '"\n'
            o_file.write(o_string)
            for s_field in dict_output[p_name]:
                if s_field != 'POINT_NAME' and s_field != 'INDEX':
                    o_string = '  ' + str(s_field) + ' = "' + str(dict_output[p_name][s_field]) + '"\n'
                    o_file.write(o_string)
            o_file.write('\n')
    o_file.close


##############################################################################
# MAIN
##############################################################################
if __name__ == "__main__":
    print("start") #debug starting because why not
    
    #Global Vairables
    #These can be prompts in the future
    Net = 0
    Unit = 0
    datatypes = {}
    dict_output = {}
    error_msg = ''
    in_def = "Defaults.csv"
    in_name = "Points.xlsx"
    out_name = "output.imp"
    defaults(in_def)
    run(in_name, out_name)
    #print(dict_output)
    print("end")


##############################################################################
#  Example Import File
##############################################################################
#
#OBJECT="POINT" ACTION="INSERT" POINT_NAME="D004P1B1L1"
#  RECORD_TYPE="RM"
#  NETWORK_ID="0"
#  UNIT_ID="1"
#  DROP_ID="4"
#  BROADCAST_FREQUENCY="S"
#  OPP_RATE="S"
#  CHARACTERISTICS="R-------"
#  SECURITY_GROUP_1="1"
#  COLLECT_ENABLED="0"
#  ALARM_PRIORITY="1"
#  IO_TASK_INDEX="2"
#  DISABLE_ALARM_CHEK_REM="0"
#  HIGHLY_MANAGED_ALARM="0"
#  SPECIAL_SENSOR_ALARM_PROCESSING="0"
#  SENSOR_CHARACTERISTICS="--------"
#  SENSOR_ALARM_PRIORITY="1"
#  HIGH_INTEGRITY="0"
#  SUMMARY_DIAGRAM_PT_GRP="0"
#  RM_DISABLE_DROP_ALARM="0"
#  IO_LOCATION="1.1.1"
#  SECURITY_GROUP_4="1"

#OBJECT="POINT" ACTION="INSERT" POINT_NAME="HS500KGV501E"
#  RECORD_TYPE="LD"
#  NETWORK_ID="0"
#  UNIT_ID="1"
#  DROP_ID="1"
#  DESCRIPTION="PUMP 5 PRIME DISCH KGV REMOTE"
#  BROADCAST_FREQUENCY="S"
#  OPP_RATE="S"
#  CHARACTERISTICS="P-------"
#  SECURITY_GROUP_1="1"
#  SECURITY_GROUP_2="1"
#  SECURITY_GROUP_3="1"
#  SECURITY_GROUP_4="1"
#  SECURITY_GROUP_16="1"
#  COLLECT_ENABLED="0"
#  SOE_ENABLED="0"
#  PERIODIC_SAVE="0"
#  TAGOUT="0"
#  UNCOMMISSIONED="0"
#  INITIAL_VALUE="0"
#  INVERTED="0"
#  RESET_SUM="0"
#  SOE_POINT="0"
#  SOE_1_SHOT_ALGORITHM="0"
#  SOE_REPORTING_OPTION="0"
#  STATUS_CHECKING_TYPE="N"
#  AUTO_RESET="1"
#  AUTO_ACKNOWLEDGE="0"
#  POWER_CHECK_ENABLE="0"
#  POWER_CHECK_CHANNEL="16"
#  ALARM_PRIORITY="1"
#  SET_DESCRIPTION="REMOTE"
#  RESET_DESCRIPTION="LOCAL"
#  IO_TASK_INDEX="2"
#  DISABLE_ALARM_CHEK_REM="0"
#  SUMMARY_ALARM_POINT="0"
#  HIGHLY_MANAGED_ALARM="0"
#  SPECIAL_SENSOR_ALARM_PROCESSING="0"
#  SENSOR_CHARACTERISTICS="--------"
#  SENSOR_ALARM_PRIORITY="1"
#  HIGH_INTEGRITY="0"
#  SUMMARY_DIAGRAM_PT_GRP="0"
#  TERMINAL_1="B2 A10"
#  TERMINAL_2="B2 B10"
#  IO_TYPE="R"
#  IO_LOCATION="1.2.8"
#  IO_CHANNEL="9"

#OBJECT="POINT" ACTION="INSERT" POINT_NAME="LIT500001A"
#  RECORD_TYPE="LA"
#  NETWORK_ID="0"
#  UNIT_ID="1"
#  DROP_ID="2"
#  DESCRIPTION="MPS OVERFLOW CH LVL"
#  BROADCAST_FREQUENCY="S"
#  OPP_RATE="S"
#  CHARACTERISTICS="P-------"
#  PERIODIC_SAVE="0"
#  TAGOUT="0"
#  UNCOMMISSIONED="0"
#  INITIAL_VALUE="0"
#  AUTO_RESET="1"
#  AUTO_ACKNOWLEDGE="0"
#  LOW_ALARM_PRIORITY_1="1"
#  LOW_ALARM_PRIORITY_2="1"
#  LOW_ALARM_PRIORITY_3="1"
#  LOW_ALARM_PRIORITY_4="1"
#  LOW_ALARM_PRIORITY_USER="1"
#  HIGH_ALARM_PRIORITY_1="1"
#  HIGH_ALARM_PRIORITY_2="1"
#  HIGH_ALARM_PRIORITY_3="1"
#  HIGH_ALARM_PRIORITY_4="1"
#  HIGH_ALARM_PRIORITY_USER="1"
#  THERMOCOUPLE_UNITS="F"
#  CONVERSION_TYPE="1"
#  CJC_TEMPERATURE_UNITS="F"
#  DISPLAY_TYPE="S"
#  SIGNIFICANT_DIGITS="2"
#  IO_TASK_INDEX="2"
#  SECURITY_GROUP_1="1"
#  SECURITY_GROUP_4="1"
#  DEADBAND_ALGORITHM="STANDARD"
#  COLLECT_ENABLED="0"
#  DISABLE_ALARM_CHEK_LIMIT_CHECK_REM="0"
#  HIGHLY_MANAGED_ALARM="0"
#  SPECIAL_SENSOR_ALARM_PROCESSING="0"
#  SENSOR_CHARACTERISTICS="--------"
#  SENSOR_ALARM_PRIORITY="1"
#  HIGH_INTEGRITY="0"
#  SUMMARY_DIAGRAM_PT_GRP="0"
#  ENGINEERING_UNITS="feet"
#  MAXIMUM_SCALE="25"
#  MINIMUM_SCALE="0"
#  TOP_OUTPUT_SCALE="25"
#  BOTTOM_OUTPUT_SCALE="0"
#  IO_TYPE="R"
#  IO_LOCATION="1.2.6"
#  IO_CHANNEL="3"
#  TERMINAL_1="B1 A7"
#  TERMINAL_2="B1 A9"
#  TERMINAL_3="B1 A8"
#############################################################################