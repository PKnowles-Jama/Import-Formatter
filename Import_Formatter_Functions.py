import pandas as pd
import re
import os

def ParseExcel(file_path: str, keyword: str, column_name: str):
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