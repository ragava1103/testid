import pandas as pd
import logging
import re

def id_processing(filepath, id_name, COMM, testid, step):
    """
    Process a list of test IDs, extract specific patterns, match with an Excel file,
    and save the results to a new Excel file.

    Args:
        filepath (str): Path to the Excel file.
        id_name (str): Column name in the Excel file to match with extracted names.
        COMM (str): Column to extract data from the Excel file.
        testid (list): List of test ID strings.
        step (list): List of associated step strings.

    Returns:
        pd.DataFrame or None: Extracted results if successful, None otherwise.
    """
    if not isinstance(testid, list) or not isinstance(step, list):
        logging.error("Invalid input: testid and step must be lists.")
        return None

    if len(testid) != len(step):
        logging.error("Length mismatch: testid and step lists must have the same length.")
        return None

    if not testid:
        logging.info("No test IDs provided.")
        return None

    # Initialize result dictionary
    result = {"number": [], "name": [], "checker": []}

    # Process test IDs and step together
    for tes_id, stp in zip(testid, step):
        logging.info(f"Processing test ID: {tes_id} with step: {stp}")
        try:
            if not isinstance(tes_id, str) or not isinstance(stp, str):
                logging.warning(f"Skipping invalid pair: ({tes_id}, {stp})")
                continue

            # Extract number from "LCD_CXL_<number>"
            match = re.search(r"LCD_CXL_(\d+)", tes_id)
            if match:
                extracted_number = match.group(1)
                result["number"].append(extracted_number)
                logging.info(f"Extracted number: {extracted_number}")
            else:
                result["number"].append(None)  # Add placeholder if no match

            # Extract name after ':'
            extracted_name = tes_id.split(":", 1)[1].strip() if ':' in tes_id else None
            result["name"].append(extracted_name)
            logging.info(f"Extracted name: {extracted_name}")

            # Append step
            result["checker"].append(stp)
            logging.info(f"Extracted step: {stp}")

        except Exception as e:
            logging.error(f"Error processing test ID '{tes_id}': {e}")

    # Stop if no valid names are extracted
    if not result["name"]:
        logging.info("No valid names extracted from test IDs.")
        return None

    # Read the Excel file
    try:
        df = pd.read_excel(filepath, sheet_name="Steps")

        # Check if required columns exist
        if id_name not in df.columns or COMM not in df.columns:
            logging.error(f"One or more specified columns ('{id_name}', '{COMM}') not found in the Excel file.")
            return None

        # Filter DataFrame based on extracted names and steps
        matched_results = []
        for name, number, checker in zip(result["name"], result["number"], result["checker"]):
            if not name:
                continue

            logging.info(f"Matching string: {name}")
            filtered_df = df[df[id_name].astype(str).str.contains(name, case=False, na=False)]
            if not filtered_df.empty:
                # Add 'number' and 'checker' columns to the filtered DataFrame
                filtered_df = filtered_df[[id_name, COMM]].copy()
                filtered_df["Extracted_Number"] = number
                filtered_df["Checker"] = checker
                matched_results.append(filtered_df)

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
