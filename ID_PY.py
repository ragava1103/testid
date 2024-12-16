import re
import logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s'
                    ,filename='app.log',filemode='a')

class ID_Proc:
   
        
    def id_processing(self, testid):
        """
        Process a list of test IDs to extract specific patterns or values.

        Args:
            testid (list): List of test ID strings.

        Returns:
            dict: A dictionary with extracted values (`number` or `name`), if found.
        """
        # Validate the input type
        if not isinstance(testid, list):
            logging.error("Invalid input: testid must be a list.")
            return {"number": None, "name": None}

        # Initialize result dictionary
        result = {"number": None, "name": None}

        # Iterate over test IDs to extract the number and the name
        for tes_id in testid:
            # Ensure each item is a string
            logging.info(f"the list of test id {tes_id}")
            if not isinstance(tes_id, str):
                logging.warning(f"Skipping non-string item: {tes_id}")
                continue

            # Extract number from the pattern "LCD_CXL_<number>"
            if result["number"] is None:
                match = re.search(r"LCD_CXL_(\d+)", tes_id)
                if match:
                    result["number"] = match.group(1)
                    logging.info(f"Extracted number from test case: {result['number']}")

            # Extract name from the part after ':'
            if result["name"] is None and ':' in tes_id:
                result["name"] = tes_id.split(":", 1)[1].strip()  # Extract and trim after the first colon
                logging.info(f"Extracted test ID name: {result['name']}")

            # Stop processing further if both number and name are found
            if result["number"] and result["name"]:
                for k,v in result.items():
                    logging.info(f'key value pairs {k}:{v}')
                continue

        # Log missing values if not found
        if result["number"] is None:
            logging.info("No number found in the test IDs.")
        if result["name"] is None:
            logging.info("No test ID name containing ':' found.")

        return result
