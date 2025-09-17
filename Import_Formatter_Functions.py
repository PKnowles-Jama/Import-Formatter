import pandas as pd
import re
import os

def DefaultParseExcel(file_path: str, column_name: str):
    """
    Reads an Excel file, parses a specified column, and saves the result to a new file.
    Handles complex cases including preambles, numbered lists, and multi-row 'GIVEN' blocks.
    """
    try:
        df = pd.read_excel(file_path)
        file_name, file_ext = os.path.splitext(file_path)
    except FileNotFoundError:
        print(f"Error: The file at {file_path} was not found.")
        return

    try:
        id_column = df.columns[0]
        if column_name not in df.columns:
            print(f"Error: Column '{column_name}' not found in the Excel file.")
            return

        processed_rows = []

        for _, row in df.iterrows():
            original_id = row[id_column]
            # Ensure we're working with a string to avoid errors
            criteria_text = str(row[column_name]) if pd.notna(row[column_name]) else ""

            if not criteria_text:
                continue

            other_cols = row.drop([id_column, column_name]).to_dict()
            
            scenarios = re.split(r'(?=\d+\.\s*Scenario:|Scenario:)', criteria_text)
            
            # --- CORRECTED PREAMBLE LOGIC ---
            # Only treat the first block as a preamble if it lacks a "Scenario:" AND
            # there are other blocks that DO have one. This prevents it from
            # incorrectly consuming the entire ID-7 block.
            if len(scenarios) > 1 and 'Scenario:' not in scenarios[0]:
                preamble = scenarios.pop(0).strip()
                if preamble: # Only append if there's actual text
                    scenarios[0] = scenarios[0] + '\n' + preamble
            
            for block in scenarios:
                block = block.strip()
                if not block:
                    continue

                # --- Logic for cases with no "Scenario:" line (ID-7) ---
                if 'Scenario:' not in block:
                    given_blocks = re.split(r'(?=GIVEN:)', block)
                    for given_block in given_blocks:
                        given_block = given_block.strip()
                        if not given_block.startswith('GIVEN:'):
                            continue
                        
                        when_match = re.search(r'WHEN:\s*(.*)', given_block, re.DOTALL)
                        scenario_name = when_match.group(1).strip().split('\n')[0] if when_match else "N/A"
                        
                        processed_rows.append({
                            'Function ID': original_id,
                            'Scenario Name': scenario_name.strip(),
                            'Description': given_block,
                            **other_cols
                        })
                    continue # Skip to the next scenario block

                # --- Default processing for normal scenarios ---
                name_match = re.search(r'Scenario:\s*(.*)', block, re.DOTALL)
                scenario_name = name_match.group(1).strip().split('\n')[0] if name_match else "N/A"

                description = re.sub(r'(\d+\.\s*)?Scenario:.*?\n', '', block, count=1).strip()

                processed_rows.append({
                    'Function ID': original_id,
                    'Scenario Name': scenario_name.strip(),
                    'Description': description,
                    **other_cols
                })

        if not processed_rows:
            print("No data was processed. The output file will not be created.")
            return

        parsed_df = pd.DataFrame(processed_rows)
        output_file_path = f"{file_name}_updated{file_ext}"
        parsed_df.to_excel(output_file_path, index=False)
        print(f"🎉 Successfully created a new file 🎉 \n {output_file_path}")

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
     
def KeywordParseExcel(file_path: str, keyword: str, column_name: str):
    """
    Parses an Excel file to create new rows based on a keyword within a specified column.

    Args:
        file_path (str): The path to the input Excel file.
        keyword (str): The keyword to search for to create new rows.
        column_name (str): The name of the column to search through for the keyword.
    """
    try:
        # Check if the file exists and get the base name and extension
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Error: The file at {file_path} was not found.")
            
        file_name, file_ext = os.path.splitext(file_path)
        
        # Read the Excel file into a pandas DataFrame, works for both .xls and .xlsx
        df = pd.read_excel(file_path)

        # Initialize an empty list to store the new rows
        new_rows = []

        # Iterate through each row in the original DataFrame
        for index, row in df.iterrows():
            # Get the text from the specified column
            text = str(row[column_name])
            
            # Split the text by the keyword, keeping the keyword in the split parts
            # This uses a positive lookahead regex to split *before* the keyword
            parts = re.split(f'(?={re.escape(keyword)})', text, flags=re.IGNORECASE)

            # Clean up empty strings that may result from the split
            parts = [part.strip() for part in parts if part.strip()]

            # Keep track of the enumeration for the 'Name' column
            enum = 1
            
            # Iterate through the split parts and create new rows
            for part in parts:
                # Create a copy of the original row
                new_row = row.copy()
                
                # Update the 'Name' column with the original name and the new enumeration
                original_name = new_row['Name']
                new_row['Name'] = f"{original_name} - {enum}"

                # Update the specified column with the new split part
                new_row[column_name] = part
                
                # Append the new row to our list
                new_rows.append(new_row)
                
                # Increment the enumeration
                enum += 1

        # Create a new DataFrame from the list of new rows
        parsed_df = pd.DataFrame(new_rows)
        
        # Create the new output file name with '_updated' appended
        output_file_path = f"{file_name}_updated{file_ext}"
        
        # Save the new DataFrame to a new Excel file
        parsed_df.to_excel(output_file_path, index=False)
        
        print(f"🎉 Successfully created a new file 🎉 \n {output_file_path}")

    except KeyError:
        print(f"Error: The column name '{column_name}' was not found in the file.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")