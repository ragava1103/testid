import pandas as pd
import re
import logging


def id_processing(filepath, id_name, COMM, testid):
    """
    Process a list of test IDs, extract specific patterns, match with an Excel file,
    and save the results to a new Excel file.

    Args:
        filepath (str): Path to the Excel file.
        id_name (str): Column name in the Excel file to match with extracted names.
        COMM (str): Column to extract data from the Excel file.
        testid (list): List of test ID strings.

    Returns:
        pd.DataFrame or None: Extracted results if successful, None otherwise.
    """
    if not isinstance(testid, list):
        logging.error("Invalid input: testid must be a list.")
        return None

    if not testid:
        logging.info("No test IDs provided.")
        return None

    # Initialize result dictionary
    result = {"number": [], "name": []}

    # Iterate over test IDs to extract numbers and names
    for tes_id in testid:
        logging.info(f"Processing test ID: {tes_id}")
        try:
            if not isinstance(tes_id, str):
                logging.warning(f"Skipping non-string item: {tes_id}")
                continue

            # Extract number from "LCD_CXL_<number>"
            match = re.search(r"LCD_CXL_(\d+)", tes_id)
            if match:
                extracted_number = match.group(1)
                result["number"].append(extracted_number)
                logging.info(f"Extracted number: {extracted_number}")

            # Extract name after ':'
            if ':' in tes_id:
                extracted_name = tes_id.split(":", 1)[1].strip()
                result["name"].append(extracted_name)
                logging.info(f"Extracted name: {extracted_name}")

        except Exception as e:
            logging.error(f"Error processing test ID '{tes_id}': {e}")

    # Stop if no valid names are extracted
    if not result["name"]:
        logging.info("No valid names extracted from test IDs.")
        return None

    # Read the Excel file
    try:
        sheets = pd.read_excel(filepath, sheet_name="Steps")
        df = sheets

        # Check if required columns exist
        if id_name not in df.columns or COMM not in df.columns:
            logging.error(f"One or more specified columns ('{id_name}', '{COMM}') not found in the Excel file.")
            return None

        # Filter DataFrame based on extracted names
        matched_results = []
        for name, number in zip(result["name"], result["number"]):
            logging.info(f"Matching string: {name}")
            filtered_df = df[df[id_name].astype(str).str.contains(name, case=False, na=False)]
            if not filtered_df.empty:
                # Add 'number' as a new column to the filtered DataFrame
                filtered_df["Extracted_Number"] = number
                matched_results.append(filtered_df[[id_name, COMM, "Extracted_Number"]])

        # Combine results into a single DataFrame
        if matched_results:
            final_result = pd.concat(matched_results, ignore_index=True)
            logging.info(f"Final extracted data:\n{final_result}")

            # Save results to an Excel file
            output_filepath = "output_result.xlsx"
            try:
                final_result.to_excel(output_filepath, index=False)  # Save without index column
                logging.info(f"Results successfully saved to {output_filepath}")
            except Exception as e:
                logging.error(f"Error saving results to Excel: {e}")

            return final_result
        else:
            logging.info("No matches found in the DataFrame.")
            return None

    except Exception as e:
        logging.error(f"Error processing the Excel file: {e}")
        return None
