import pandas as pd
import logging

def finding_test(filepath, id_name, COMM, con):
    ''' 
    Process the Excel file to find a specific string or value in the specified columns.
    
    Args:
        filepath (str): Path to the Excel file.
        id_name (str): Name of the column to filter non-null values.
        COMM (str): Column to extract data from after filtering.
        con (dict): A dictionary containing filtering conditions (e.g., {"name": "value to match"}).

    Returns:
        pd.Series or None: Extracted data from the `COMM` column if conditions are met, otherwise None.
    ''' 
    try:
        # Load the "Steps" sheet
        sheets = pd.read_excel(filepath, sheet_name="Steps")
        df = sheets  # Assign the DataFrame directly

        # Check if required columns exist
        if id_name not in df.columns or COMM not in df.columns:
            logging.error(f"One or more specified columns ('{id_name}', '{COMM}') not found in the Excel file.")
            return None

        logging.info(f"Type of 'con' variable: {type(con)}")

        # Validate 'con' and extract filter value
        if not isinstance(con, dict) or "name" not in con:
            logging.error("Invalid 'con' parameter. It must be a dictionary with a 'name' key.")
            return None

        acc_string = con["name"]
        logging.info(f"String to match: {acc_string}")

        # Filter the DataFrame based on the condition
        filtered_df = df[df[id_name].astype(str).str.contains(acc_string, case=False, na=False)]

        logging.info(f"Filtered DataFrame:\n{filtered_df}")

        # Extract and return data from the 'COMM' column
        if not filtered_df.empty:
            result = filtered_df[COMM]
            logging.info(f"Extracted data from '{COMM}' column:\n{result}")
            return result

        logging.info("No matching rows found.")
        return None

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return None
