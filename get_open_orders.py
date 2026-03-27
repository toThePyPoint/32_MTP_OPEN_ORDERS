import sys
import time
import traceback
from datetime import datetime
import pandas as pd

from helper_functions import clean_number, generate_zsdkap_filename, update_open_mtp_excel

OUTPUT_FILE_PATH = r"P:\Technisch\PLANY PRODUKCJI\PLANIŚCI\PP_TOOLS_TEMP_FILES\13_PPS_OPEN_ORDERS\OUTPUT"

zsdkap_dtypes = {
    'Odbiorca materia≈Ç√≥w': 'string',
    'Materia≈Ç': 'string',
    'Nazwa': 'string',
    'Dokument sprzeda≈ºy': 'string',
    'Pozycja': 'string',
    'Kontroler MRP': 'string',
    'Ilo≈õƒá zlecenia': 'string',
    # 'WA-Datum': 'datetime64[ns]',
}

zsdkap_new_columns_names = {
    'Odbiorca materia≈Ç√≥w': 'receiver',
    'Materia≈Ç': 'mat_number',
    'Nazwa': 'mat_description',
    'Dokument sprzeda≈ºy': 'customer_order_number',
    'Pozycja': 'customer_order_position',
    'Kontroler MRP': 'mrp_controller',
    'Ilo≈õƒá zlecenia': 'orders_quantity',
    'WADAT': 'dispatch_date'
}

department_file_names_map = {
    'mont': 'ROTO_StrategicPlanning_Worksheet_0301_MONT_2026.xlsm',
    'wmo': 'ROTO_StrategicPlanning_Worksheet_2101_WMO_2026.xlsm',
    'wmr': 'ROTO_StrategicPlanning_Worksheet_2101_WMR_2026.xlsm'
}

def create_paths(zsdkap_report_name):
    global ZSDKAP_FILE_PATH
    # ZSDKAP_FILE_PATH = fr'C:\Temp\Kamil\Prywatne\07_Programowanie\99_Moje_projekty\28_PPS_KPI\excel_files/job/{zsdkap_report_name}.csv'
    ZSDKAP_FILE_PATH = fr'\\rfmesrv5\connect\DST_SAP_Transfer\P11\PPS_LUB\02_MID_TERM_PLANNING_ALIGNMENT/{zsdkap_report_name}.csv'


def get_zsdkap_df_grouped_by_date(mrp_controller, mat_name, df, date_limit=None):
    tmp = df.copy()
    if date_limit is not None:
        tmp = tmp[tmp['dispatch_date'] <= date_limit]

    tmp = tmp[(tmp['mrp_controller'].isin(mrp_controller)) & (tmp['mat_description'].str.startswith(mat_name))]
    tmp = tmp[['dispatch_date', 'orders_quantity']]
    return tmp.groupby('dispatch_date', as_index=False).sum()


def collect_open_orders(zsdkap_report_name="zsdkap", mrp_controller='L1K', mat_name='R4'):
    # Ensure mrp_controller is always a tuple
    if not isinstance(mrp_controller, (list, tuple, set, pd.Series)):
        mrp_controller = mrp_controller,

    # Ensure mat_name is always a tuple
    if not isinstance(mat_name, (list, tuple, set, pd.Series)):
        mat_name = mat_name,

    create_paths(zsdkap_report_name)

    zsdkap_df = pd.read_csv(ZSDKAP_FILE_PATH, dtype=zsdkap_dtypes, sep=';', encoding='MacRoman')
    zsdkap_df = zsdkap_df.rename(columns=zsdkap_new_columns_names)
    zsdkap_df['dispatch_date'] = pd.to_datetime(zsdkap_df['dispatch_date'], dayfirst=True, errors='coerce')

    # Przetwarzanie konkretnej kolumny
    zsdkap_df['orders_quantity'] = zsdkap_df['orders_quantity'].apply(clean_number)

    zsdkap_df_grouped = get_zsdkap_df_grouped_by_date(mrp_controller, mat_name, zsdkap_df)
    zsdkap_df_grouped['dispatch_date'] = zsdkap_df_grouped['dispatch_date'].dt.date

    return zsdkap_df_grouped


def open_orders_loop(lines, mrp_controllers, product_names, zsdkap, file_name):
    try:
        for line, mrp, prd_name in zip(lines, mrp_controllers, product_names):
            open_orders_df = collect_open_orders(zsdkap_report_name=zsdkap, mrp_controller=mrp, mat_name=prd_name)

            # Append data to Excel
            open_orders_df.to_excel(f"{OUTPUT_FILE_PATH}/{line}.xlsx", index=False)
            update_open_mtp_excel(open_orders_df, file_name,f"{line}_hs", 36)

    except Exception as e:
        print("Błąd: ", e)
        error_details = traceback.format_exc()
        print("Szczegóły błędu:\n", error_details)
        input("Press Enter...")


def wmo_open_orders(mtp_file_name):
    zsdkap = generate_zsdkap_filename()

    lines = ["P100", "M200", "M300", "M320", "M500", "M600", "MDA", "ASA"]
    mrp_controllers = ['L1K', ('L1H', 'L41', 'L3H', 'L82'), ('L3H', 'L82'), 'L2H', 'LD1', 'LZ1', 'LMD', 'LAS']
    product_names = [('R4', 'R7', 'R3', 'R5', 'EFL_R4', 'EFL_R7'), ('R4', 'R7', 'R3', 'R5', 'EFL_R4', 'EFL_R7', 'EFL 4', 'EFL 7'), ('R6', 'R8', 'EFL_R6', 'EFL_R8', 'EFL 6', 'EFL 8'), ('Q4', 'EFL_Q'), 'R2', ('ZI', 'KO', 'Li'), ('MDA'), ('ASA', 'ASI')]  # Product names starts with...

    open_orders_loop(lines, mrp_controllers, product_names, zsdkap, mtp_file_name)


def wmr_open_orders(mtp_file_name):
    zsdkap = generate_zsdkap_filename()

    lines = ["ZRV", "ZJA", "ZFA", "ZRI", "ZAR"]
    mrp_controllers = [('L2E', 'L2V', 'LI1', 'LI3'), ('L2J', 'LI5', 'LI8'), ('L2F', 'LI6'), 'L2I', ('L2B', 'L2R', 'LI2', 'LI4', 'LI7')]
    product_names = [('ZRE_M', 'ZRE M', 'ZRV_M', 'ZRV M'), ('ZJA', 'ZRE_E', 'ZRE E', 'ZRV_E', 'ZRV E'), 'ZFA', 'ZRI', ('ZAR', 'Auss', 'BHG', 'ZRS')]  # Product names starts with...

    open_orders_loop(lines, mrp_controllers, product_names, zsdkap, mtp_file_name)


def mont_open_orders(mtp_file_name):
    '''
    BMH KPIs
    '''
    zsdkap = generate_zsdkap_filename()

    lines = ["WDF68K", "WDFQK", "ZRO", "QR1", "EDR"]
    mrp_controllers = [
        ('M81', 'M82'),
        ('MQ1', 'MQ2'),
        ('MR1', 'MR2', 'MR3'),
        "MR4",
        ('MEB', 'MED', 'MEE', 'MEI', 'MEH', 'MEJ', 'MEM', 'MEN', 'MEX')
    ]

    product_names = [
        ('R6', 'R8', 'I8', 'EFL', 'ABR'),
        ('Q4', 'QRA', 'Qt4', 'EFL', 'ABR'),
        ('ZRO', 'ZMA'),
        "ZRO",
        ('ED', 'EF', 'EA')
    ]
    # Product names starts with...
    open_orders_loop(lines, mrp_controllers, product_names, zsdkap, mtp_file_name)


if __name__ == "__main__":
    department = sys.argv[1]
    # department = 'wmr'
    try:
        file_name = department_file_names_map[department]

        if department == 'wmo':
            wmo_open_orders(file_name)
        elif department == 'wmr':
            wmr_open_orders(file_name)
        elif department == 'mont':
            mont_open_orders(file_name)
    except Exception as e:
        print(e)
        input("Press Enter...")

    # mont_open_orders(file_name)
    # wmo_open_orders(file_name)
    # wmr_open_orders(file_name)