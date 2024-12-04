import pandas as pd
import os

def split_excel_by_column(file_path, column_name, output_dir=None):
    """
    Splits an Excel sheet into multiple Excel files based on unique values in a specified column.
 
    :param file_path: Path to the input Excel file.
    :param column_name: Column name to split the file on.
    :param output_dir: Directory to save the output files. If None, saves in the current directory.
    """
    # Read the Excel file
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading the file: {e}")
        return
    if column_name not in df.columns:
        print(f"Column '{column_name}' not found in the Excel file.")
        return
    # Set the output directory
    if output_dir is None:
        output_dir = os.getcwd()
    os.makedirs(output_dir, exist_ok=True)
    # Group by the unique values in the specified column and save each group
    unique_values = df[column_name].unique()
    for value in unique_values:
        subset = df[df[column_name] == value]
        output_file = os.path.join(output_dir, f"{value}.xlsx")
        try:
            subset.to_excel(output_file, index=False, engine='openpyxl')
            print(f"Created file: {output_file}")
        except Exception as e:
            print(f"Error saving file for value '{value}': {e}")
    print("Splitting complete!")
 
# Example usage:
if __name__ == "__main__":
    input_file = input("Enter the path to the Excel file: ").strip()
    column = input("Enter the column name to split by: ").strip()
    output_directory = input("Enter the output directory (or leave blank for current directory): ").strip()
    output_directory = output_directory if output_directory else None
    split_excel_by_column(input_file, column, output_directory)