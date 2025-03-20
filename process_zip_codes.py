import pandas as pd
import re

def process_zip_codes(input_file="sample_data.xlsx", sheet_name=0, zip_column=None):
    """
    Process ZIP codes in an Excel file by:
    1. Converting them to 5-digit zero-padded text format
    2. Truncating to exactly 5 digits if longer
    3. Replacing invalid ZIP codes (like 00000) with a default value
    4. Saving the changes back to the original file
    
    Parameters:
    - input_file: Path to the Excel file
    - sheet_name: Sheet to process (default: 0, the first sheet)
    - zip_column: Name of the ZIP code column (default: will look for a column named 'ZIP')
    """
    print(f"Processing file: {input_file}")
    
    try:
        # Read the Excel file
        xl = pd.ExcelFile(input_file)
        sheet_names = xl.sheet_names
        print(f"Found sheets: {sheet_names}")
        
        # Use the actual sheet name instead of index
        actual_sheet_name = sheet_names[sheet_name] if isinstance(sheet_name, int) else sheet_name
        print(f"Using sheet: {actual_sheet_name}")
        
        # Read the specified sheet
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        print(f"Successfully read sheet")
        print(f"Found columns: {df.columns.tolist()}")
        
        # Find the ZIP column if not specified
        if zip_column is None:
            # Look for common zip column names (case insensitive)
            zip_columns = [col for col in df.columns if 'zip' in str(col).lower()]
            if zip_columns:
                zip_column = zip_columns[0]
                print(f"Found ZIP column: {zip_column}")
            else:
                raise ValueError("ZIP column not found. Please specify the zip_column parameter.")
        
        # Ensure the specified column exists
        if zip_column not in df.columns:
            raise ValueError(f"Column '{zip_column}' not found in the Excel file.")
        
        # Print some stats about the ZIP column before processing
        print("\nBefore processing:")
        print(f"ZIP column data type: {df[zip_column].dtype}")
        print(f"Sample of original ZIP values: {df[zip_column].head(10).tolist()}")
        print(f"Unique ZIP values count: {df[zip_column].nunique()}")
        
        # Count problematic ZIP codes
        zeros_count = df[df[zip_column] == "00000"].shape[0] + df[df[zip_column] == 0].shape[0]
        nulls_count = df[zip_column].isna().sum()
        print(f"Found {zeros_count} entries with '00000' values")
        print(f"Found {nulls_count} null or missing values")
        
        # Process ZIP codes
        def format_zip(x):
            if pd.isna(x) or str(x).strip() == '':
                return '99999'  # Default value instead of blank
            
            # Try to extract digits only
            digits = re.sub(r'\D', '', str(x))
            
            # If no digits, use default instead of blank
            if not digits:
                return '99999'  # Default value
                
            # Convert to integer and back to string to remove leading zeros
            try:
                int_zip = int(digits)
                
                # If it's 0, use default instead of blank
                if int_zip == 0:
                    return '99999'  # Default value
                    
                # Format as 5-digit string with leading zeros
                formatted_zip = str(int_zip).zfill(5)[:5]
                
                # If the result is "00000", replace with our default
                if formatted_zip == "00000":
                    return '99999'
                    
                return formatted_zip
            except ValueError:
                return '99999'  # Default value for any conversion errors
        
        # Apply the ZIP code formatting
        df[zip_column] = df[zip_column].apply(format_zip)
        
        # Print some stats after processing
        print("\nAfter processing:")
        print(f"Sample of processed ZIP values: {df[zip_column].head(10).tolist()}")
        print(f"Unique ZIP values count: {df[zip_column].nunique()}")
        
        # Count empty values after processing
        empty_count = df[df[zip_column] == ''].shape[0]
        print(f"Entries with empty values after processing: {empty_count}")
        
        # Rename the column to "ZIP" if it's not already named that
        if zip_column.upper() != "ZIP":
            print(f"Renaming column '{zip_column}' to 'ZIP'")
            df = df.rename(columns={zip_column: "ZIP"})
            zip_column = "ZIP"
        
        # Create a writer with the original file
        with pd.ExcelWriter(input_file, engine='openpyxl') as writer:
            # Save the modified sheet - use the actual sheet name
            df.to_excel(writer, sheet_name=actual_sheet_name, index=False)
        
        print(f"\nZIP codes processed successfully. Changes saved to {input_file}")
        return df
    
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        raise

# Main execution
if __name__ == "__main__":
    # Process the Excel file with the specified file name
    process_zip_codes("sample_data.xlsx", zip_column="ZIP")