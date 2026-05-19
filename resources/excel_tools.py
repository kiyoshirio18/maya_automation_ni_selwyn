import os
import tempfile
import win32com.client as win32
import polars as pl
import pythoncom
from io import BytesIO

def save_xlsx(_df: pl.DataFrame, formatting) -> bytes:
    output = BytesIO()
    if _df is not None:
        _df.write_excel(
            output, 
            autofit=True,
            column_formats=formatting)
    output.seek(0)
    return output.getvalue()

def xlsx_to_xls(input_bytes):
    excel = None  # Initialize excel as None
    existing_excel = False  # Flag to check if Excel was already running

    try:
        pythoncom.CoInitialize()
        # Create temporary files for input (.xlsx) and output (.xls)
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_xlsx:
            temp_xlsx_path = temp_xlsx.name
            temp_xlsx.write(input_bytes)
        
        with tempfile.NamedTemporaryFile(suffix=".xls", delete=False) as temp_xls:
            temp_xls_path = temp_xls.name

        # Try to attach to an existing Excel application if it is running
        try:
            excel = win32.GetObject(None, 'Excel.Application')
            existing_excel = True  # Excel was already running
            excel.DisplayAlerts = False  # Disable alerts like overwrite confirmation.
        except Exception:
            # Excel is not running, create a new instance
            excel = win32.Dispatch('Excel.Application')
            excel.DisplayAlerts = False
            existing_excel = False  # New instance created

        # Open the workbook (Excel file)
        workbook = excel.Workbooks.Open(temp_xlsx_path)

        # Save it as an .xls file (Excel 97-2003 Workbook format)
        workbook.SaveAs(temp_xls_path, FileFormat=56, ConflictResolution=2)  # 56 is the file format for .xls

        # Close only the workbook that was opened by this script
        workbook.Close(SaveChanges=False)

        # Read the output file into bytes
        with open(temp_xls_path, 'rb') as output_file:
            output_bytes = output_file.read()

        # Clean up temporary files
        os.remove(temp_xlsx_path)
        os.remove(temp_xls_path)

        print("Conversion successful!")
        return output_bytes
    except Exception as e:
        print(f"Conversion failed: {e}")
        return None
    finally:
        # Quit Excel only if it was not running before the script
        if excel and not existing_excel:
            try:
                excel.Quit()
            except AttributeError:
                print("Excel application not initialized properly or already closed.")
            except Exception as e:
                print(f"Failed to quit Excel properly: {e}")

def cast_columns(df: pl.DataFrame, column_types: dict) -> pl.DataFrame:
    """
    Casts columns of a Polars DataFrame to the specified data types.

    Parameters:
        df (pl.DataFrame): The input Polars DataFrame.
        column_types (dict): A dictionary where keys are column names and values are Polars data types.

    Returns:
        pl.DataFrame: The DataFrame with updated column types.
    """
    for column, dtype in column_types.items():
        if column in df.columns:
            df = df.with_columns(df[column].cast(dtype))
        else:
            print(f"Warning: Column '{column}' not found in DataFrame. Skipping.")
    return df
