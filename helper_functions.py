from datetime import datetime
import xlwings as xw
import pandas as pd


# Custom function to clean number formatting
def clean_number(value):
    try:
        # Zamiana separatorów tysięcy na pusty znak
        value = value.replace('.', '').replace(',', '.')
        return float(value)  # Konwersja na float
    except:
        return None  # W przypadku błędów (np. inne wartości) zwróć NaN


def generate_zsdkap_filename():
    today_str = datetime.today().strftime("%Y%m%d")
    filename = f"zsdkap_{today_str}_REP_LU_PPS001A"

    return filename


def update_open_mtp_excel(data_frame, file_name, sheet_name, starting_row):
    """
    Updates an open Excel file using xlwings by writing values from `orders_quantity`
    into column T based on matching `dispatch_date` in column A.

    Arguments:
    data_frame -- pandas.DataFrame: The DataFrame containing 'dispatch_date' and 'orders_quantity'.
    sheet_name -- string: The name of the worksheet to update.
    starting_row -- int: The starting row of the worksheet.

    Returns:
    None
    """
    # Ensure 'dispatch_date' column is datetime
    data_frame["dispatch_date"] = pd.to_datetime(data_frame["dispatch_date"], errors="coerce")

    # Connect to the running Excel instance
    # workbook = xw.Book(file_path)  # Open workbook from the mapped local path

    sheet = None

    app = xw.apps.active  # Check the active Excel application instance
    for workbook in app.books:
        if workbook.name == file_name:
            sheet = workbook.sheets[sheet_name]
            break

    if sheet:
        sheet.range("T36:T295").clear_contents()

        # Iterate through rows in the sheet
        row = starting_row
        while True:
            excel_date = sheet.range(f"A{row}").value  # Retrieve date from column A
            if excel_date is None:  # If column A is empty, break
                break

            try:
                # Convert Excel date to pandas datetime for comparison
                excel_date = pd.to_datetime(excel_date).date()
            except:
                row += 1
                continue

            # Check if the date exists in the DataFrame
            if excel_date in data_frame["dispatch_date"].dt.date.values:
                # Find the corresponding `orders_quantity` value
                quantity = data_frame.loc[data_frame["dispatch_date"].dt.date == excel_date, "orders_quantity"].values[0]
                # Write the value into column T of the same row
                sheet.range(f"T{row}").value = quantity

            row += 1

        print(f"Excel file — sheet: {sheet_name} — has been successfully updated.")

    else:
        print(f"Problem occurred — sheet: {sheet_name} — NOT FOUND!!.")


