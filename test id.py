import pandas as pd
import logging 
import openpyxl
import sys
import warnings




from ID_PY import *

sys.path.append(r"C:\\testid\\NGATestlines_pve_lcd.xlsx")

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
logging.basicConfig(level=logging.INFO,  # Set the logging level
format='%(asctime)s - %(levelname)s - %(message)s',filename='app.log',  # Log messages will be saved to this file
    filemode='a')
def process_excel(file_path, GOAL_column,TESTstep):
    try:
        # Use a correct format for sheet_name
        sheets = pd.read_excel(file_path, sheet_name="TestLines")  # Load the "TestLines" sheet correctly as DataFrame
        df1 = sheets  # Directly use the DataFrame returned

        # Drop rows where the GOAL_column is NaN
        df1_1 = df1.dropna(subset=[GOAL_column,TESTstep])

        if GOAL_column not in df1_1.columns and TESTstep not in df1_1.columns:
            logging.info(f"Column '{GOAL_column}' and test {TESTstep} not found in the Excel file.")
            return None
        else:
            testid = df1_1[GOAL_column]
            filtered =df1_1[TESTstep]
            logging.info(f"Found test IDs: {testid.tolist()}")
            logging.info(f'Found test Rocket or pysces{filtered.tolist()}')
        
    except Exception as e:
        logging.error(f"Error reading Excel file: {e}")
        return None

    return testid,filtered

def main_function():
    
        test__id,fil=process_excel("NGATestlines_pve_lcd.xlsx","GoalName","TestStep5")
        #logging.info(f"type of the cons :{type(test__id)}")
        data =  pd.Series(test__id)
        data_list = data.tolist()
        step =  pd.Series(fil)
        step_list = step.tolist()
        #logging.info(f" list of data to be printer {data_list}")
        id_processing("NGATestlines_pve_lcd.xlsx","Name","Command",data_list,step_list)     
        #logging.info(f"data type of list value {type(cons)}")                       
        #finding_test("NGATestlines_pve_lcd.xlsx","Name","Command",cons)
        
if __name__  == "__main__" :
    logging.info("programming start...........")
    main_function()
    logging.info("********programming***END**")