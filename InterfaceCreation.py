from tkinter import *
from tkinter import filedialog
import customtkinter
import os
import mysql.connector
import pandas as pd
from pandas.io import sql as sql
from pyxlsb import open_workbook as open_xlsb
from pyepc import SGTIN
from pyepc.exceptions import DecodingError
from sqlalchemy import create_engine
import pymysql

pymysql.install_as_MySQLdb()
import Store

global store
store = Store.Store(None, None, None, None, None, None, None, None, None, None, None, None, None, None, None,
                    None, None, None, None)
global store_list
store_list = []
global store_num
global date
global new_store_button
global main_frame

global cycle_count_paths
global qb_master_path
global item_file_path

global app
global conn
global cursor


def import_cycle_count():
    global store
    print("Importing Cycle Count...")
    pop_up_title = "Select Cycle Count Data (.txt)"
    filenames = filedialog.askopenfilenames(initialdir="/", title=pop_up_title,
                                            filetypes=(("txt files", "*.txt"), ("all files", "*.*")))
    global cycle_count_paths
    cycle_count_paths = []
    for filename in filenames:
        cycle_count_paths.append(filename)
    print(cycle_count_paths)
    store.set_cycle(cycle_count_paths)


def decodePreparation():
    epc_list = []
    for filename in cycle_count_paths:
        f = open(filename, "r")
        lines = f.readlines()
        for x in lines:
            epc_list.append(x.split('\n')[0])
        f.close()

    epc_list_no_dupe = [*set(epc_list)]
    epc_list_df = pd.DataFrame(epc_list_no_dupe, columns=['EPCs'])

    print("Preparing to Decode...")
    return epc_list_df


def decodeCycleCount(epc_list_df):
    global store
    epc_list = []
    columns = epc_list_df.columns.tolist()

    for _, i in epc_list_df.iterrows():
        for col in columns:
            epc_list.append(i[col])

    temp_epc_list = []
    for epc in epc_list:
        temp_epc_list.append(str(epc))

    epc_list = temp_epc_list

    res = list(map(''.join, epc_list))
    epc_list = [*set(res)]

    upc_list, error_epcs, error_upcs = [], [], []

    print("Decoding...")
    for epc in epc_list:
        try:
            upc_list.append(SGTIN.decode(epc).gtin)
        except DecodingError as de:
            error_epcs.append(epc)
            error_upcs.append(de)
        except TypeError as te:
            error_epcs.append(epc)
            error_upcs.append(te)

    for epc in error_epcs:
        if epc in epc_list:
            epc_list.remove(epc)

    for upc in range(len(upc_list)):
        upc_list[upc] = upc_list[upc].lstrip('0')

    unique_upc_list = []
    for i in range(len(upc_list)):
        if upc_list[i] not in unique_upc_list:
            unique_upc_list.append(upc_list[i])


    unique_epcs_df = pd.DataFrame(epc_list, columns=['EPCs'])
    store.set_UE(unique_epcs_df)

    duplicate_upcs_df = pd.DataFrame(upc_list, columns=['UPCs'])
    store.set_DU(duplicate_upcs_df)

    unique_upcs_df = pd.DataFrame(unique_upc_list, columns=['UPCs'])
    store.set_UU(unique_upcs_df)

    error_epcs_df = pd.DataFrame(error_epcs, columns=['EPCs'])
    store.set_error_EPCs(error_epcs_df)

    error_messages_df = pd.DataFrame(error_upcs, columns=['Error Messages'])
    store.set_error_messages(error_messages_df)

    return epc_list, upc_list


def import_item_file():
    print("Item File...")
    pop_up_title = "Select Item File (GM) (.csv)"
    filename = filedialog.askopenfilename(initialdir="/", title=pop_up_title,
                                          filetypes=(("csv files", "*.csv"), ("all files", "*.*")))
    global item_file_path
    global store
    item_file_path = filename
    print(item_file_path)
    store.set_item_file(item_file_path)


def import_qb_master_items():
    print("QB Master Items...")
    pop_up_title = "Select QB Master Items (.xlsb)"
    filename = filedialog.askopenfilename(initialdir="/", title=pop_up_title,
                                          filetypes=(("xlsb files", "*.xlsb"), ("all files", "*.*")))
    global qb_master_path
    qb_master_path = filename
    print(qb_master_path)


def validate_store_input():
    global store_num
    store_num = store_entry.get()
    global store
    try:
        int(store_num)
        print("Store Number: {}".format(store_num))
        store.set_store_num(store_entry.get())
        return True
    except:
        print(":: ERROR :: Store Num is not an int!")
        return False


def validate_date_input():
    global date
    global store
    date = date_entry.get()
    try:
        if date == "":
            return False
        date_list = date.split(".")
        if len(date_list[0]) == 4 and isinstance(int(date_list[0]), int):
            if len(date_list[1]) == 2 and isinstance(int(date_list[1]), int):
                if len(date_list[2]) == 2 and isinstance(int(date_list[2]), int):
                    print("Date: {}".format(date))
                    store.set_date_input(date_entry.get())
                    return True
    except:
        print(":: ERROR :: Date input is not valid!")
        return False


def validate_cycle_count_paths():
    global cycle_count_paths
    global store
    try:
        if cycle_count_paths == "":
            return False
        else:
            return True
    except:
        print(":: ERROR :: Cycle Count Data file paths have not been specified!")
        return False


def validate_item_file_path():
    global item_file_path
    global store
    try:
        if item_file_path == "":
            return False
        else:
            return True
    except:
        print(":: ERROR :: Item File path has not been specified!")
        return False


def validate_qb_path():
    global qb_master_path
    global store
    try:
        if qb_master_path == "":
            return False
        else:
            store.set_qb_path(qb_master_path)
            return True
    except:
        print(":: ERROR :: QB Master Items file path has not been specified!")
        return False


def connect_to_mysql():
    try:
        global conn
        conn = mysql.connector.connect(user='root', password='password', host='127.0.0.1', database='instoretracking',
                                       allow_local_infile=True)
        global cursor
        cursor = conn.cursor()
        stmt00 = "SET GLOBAL local_infile=1;"
        cursor.execute(stmt00)
        print("Connected to MySQL...")
    except:
        print(":: ERROR :: Something went wrong! Unable to connect to MySQL!")


def import_cycle_count_sql(epc_list, upc_list):
    global store
    print("Creating UPC Drop Table...")
    stmt = "DROP TABLE if exists UPCDrop;"
    cursor.execute(stmt)
    stmt1 = "CREATE TABLE if not exists UPCDrop(EPCs text, UPCs bigint);"
    cursor.execute(stmt1)
    cursor.executemany("""
                INSERT INTO UPCDrop(EPCs, UPCs)
                VALUES (%s, %s)
            """, list(zip(epc_list, upc_list)))

    conn.commit()
    print('Data entered successfully.')


def import_item_file_sql():
    global store
    stmt = "DROP TABLE IF EXISTS ItemFile"
    cursor.execute(stmt)
    statement_headers = "CREATE TABLE ItemFile(store_number int, REPL_GROUP_NBR int, gtin bigint, ei_onhand_qty int, " \
                        "SNAPSHOT_DATE text, UPC_NBR bigint, UPC_DESC text, ITEM1_DESC text, dept_nbr int, " \
                        "DEPT_DESC text, MDSE_SEGMENT_DESC text, MDSE_SUBGROUP_DESC text, ACCTG_DEPT_DESC text, " \
                        "DEPT_CATG_GRP_DESC text, DEPT_CATEGORY_DESC text, DEPT_SUBCATG_DESC text, VENDOR_NBR int, " \
                        "VENDOR_NAME text, BRAND_OWNER_NAME text, BRAND_FAMILY_NAME text)"
    cursor.execute(statement_headers)

    # --------------Loads both Item Files into single ItemFile table------------------------------------------------
    item_file_path_corrected = item_file_path.replace(" ", "\\ ")

    stmt = "LOAD DATA LOCAL INFILE \'{}\' " \
           "INTO TABLE ItemFile " \
           "CHARACTER SET latin1 " \
           "FIELDS TERMINATED BY \',\' " \
           "OPTIONALLY ENCLOSED BY \'\"\' " \
           "LINES TERMINATED BY \'\\r\\n\' " \
           "IGNORE 1 ROWS;".format(item_file_path_corrected)

    print(" -- Starting Item File import...")
    cursor.execute(stmt)
    print(" -- Item File import complete.")
    conn.commit()


def import_qb_sql():
    global store
    try:
        stmt = "DROP TABLE IF EXISTS QBMasterItems"
        cursor.execute(stmt)
        df = []

        print(" -- Converting QB Master Items File to .csv")

        with open_xlsb(qb_master_path) as wb:
            with wb.get_sheet("Master_Items") as sheet:
                for row in sheet.rows():
                    df.append([item.v for item in row])

        df = pd.DataFrame(df[1:], columns=df[0])

        qb_csv_path = os.path.splitext(qb_master_path)[0]
        qb_csv_path += ".csv"
        count = 1
        while os.path.exists(qb_csv_path):
            qb_csv_path = os.path.splitext(qb_master_path)[0] + " (" + str(count) + ")" + ".csv"
            count += 1

        df.to_csv(qb_csv_path, index=False)  # to generate a .csv file

        with open(qb_csv_path, "rb") as file:
            lines = file.readlines()

        with open(qb_csv_path, "wb") as file:
            for line in lines:
                file.write(line.replace(b"\r\n", b"\n"))

        print(" -- QB Master Items file conversion to .csv complete.")

        statement_headers = "CREATE TABLE QBMasterItems(Year text, Record_ID_NBR text, Item_Validation_Status text, " \
                            "Item_Arrival_Status text, Vendor_Number text, Vendor_Name text, Dept_NBR text, SBU text, UPC text, " \
                            "Item_Description text, Arrival_Month text, Max_Shipped_On_Date text, Offshore text)"
        cursor.execute(statement_headers)

        qb_path_corrected = qb_csv_path.replace(" ", "\\ ")

        stmt1 = "LOAD DATA LOCAL INFILE \'{}\' " \
                "INTO TABLE QBMasterItems " \
                "CHARACTER SET latin1 " \
                "FIELDS TERMINATED BY \',\'" \
                "OPTIONALLY ENCLOSED BY \'\"\' " \
                "LINES TERMINATED BY \'\\n\' " \
                "IGNORE 1 ROWS;".format(qb_path_corrected)

        print(" -- Starting QB Master Items import...")
        cursor.execute(stmt1)
        cursor.execute("UPDATE qbmasteritems SET UPC = REPLACE(UPC, \".0\", \"\");")  # Removes '.0' from end of UPC.
        stmt = "DROP TABLE IF EXISTS QB_IVS;"
        cursor.execute(stmt)
        stmt = "CREATE TABLE QB_IVS AS SELECT * FROM qbmasteritems " \
               "WHERE Item_Validation_Status IN ('Pass', 'Not Submitted');"
        cursor.execute(stmt)
        print(" -- QB Master Items import complete.")
        conn.commit()
        os.remove(qb_csv_path)

    except Exception as e:
        print(":: ERROR :: Could not import QB Master Items file!")
        print(e)


def create_matching_sql():
    global store
    stmt = "DROP TABLE IF EXISTS Matching;"
    cursor.execute(stmt)
    print(" -- Creating Matching table...")
    stmt = """
        Create Table Matching AS SELECT DISTINCT gtin, 
        MAX(DEPT_CATG_GRP_DESC) AS DEPT_CATG_GRP_DESC,
        MAX(DEPT_CATEGORY_DESC) AS DEPT_CATEGORY_DESC,
        MAX(VENDOR_NBR) AS VENDOR_NBR, 
        MAX(VENDOR_NAME) AS VENDOR_NAME,
        MAX(BRAND_FAMILY_NAME) AS BRAND_FAMILY_NAME,
        MAX(dept_nbr) as dept_nbr, 
        MAX(REPL_GROUP_NBR) AS REPL_GROUP_NBR 
        FROM ItemFile 
        WHERE gtin IN (SELECT UPCs from UPCDrop) 
        AND dept_nbr IN ('7','9','14','17','20','22','71','72','74','87') 
        GROUP BY gtin;
    """
    cursor.execute(stmt)
    conn.commit()
    cursor.execute("ALTER TABLE Matching ADD COLUMN UPC_No_Check bigint AFTER REPL_GROUP_NBR;")
    cursor.execute("UPDATE Matching SET UPC_No_Check = LEFT(gtin, length(gtin)-1);")
    conn.commit()

    df = sql.read_sql('SELECT * FROM Matching', conn)
    store.set_matching(df)


def create_qb_matching_sql():
    global store
    stmt = "DROP TABLE IF EXISTS QB_Matching;"
    cursor.execute(stmt)

    stmt = """
        CREATE TABLE QB_Matching AS SELECT 
        m.gtin, m.DEPT_CATG_GRP_DESC, m.DEPT_CATEGORY_DESC, m.VENDOR_NBR, m.VENDOR_NAME, m.BRAND_FAMILY_NAME, m.dept_nbr, 
        m.REPL_GROUP_NBR, qb.Item_Validation_Status
        FROM Matching m
        INNER JOIN QB_IVS qb
        ON qb.UPC = m.UPC_No_Check;
        """
    cursor.execute(stmt)
    conn.commit()
    print(" -- QB Matching Table creation complete.")
    df = sql.read_sql('SELECT * FROM QB_Matching', conn)
    store.set_qb_matching(df)


def create_total_items_sql():
    global store
    cursor.execute("DROP TABLE IF EXISTS TotalItems;")
    stmt = """
            CREATE TABLE TotalItems AS
	        SELECT DISTINCT UPCDrop.EPCs,
	        itemfile.gtin,
	        itemfile.DEPT_CATG_GRP_DESC,
	        itemfile.DEPT_CATEGORY_DESC, 
	        itemfile.VENDOR_NBR,
	        itemfile.VENDOR_NAME,
	        itemfile.BRAND_FAMILY_NAME,
	        itemfile.dept_nbr
	        FROM itemfile
	        INNER JOIN UPCDrop ON UPCDrop.UPCs = itemfile.gtin
	        WHERE UPCDrop.UPCs = itemfile.gtin and dept_nbr IN ('7','9','14','17','20','22','71','72','74','87');
          """
    cursor.execute(stmt)
    cursor.execute("ALTER TABLE TotalItems ADD COLUMN UPC_No_Check bigint AFTER dept_nbr;")
    cursor.execute("UPDATE TotalItems SET UPC_No_Check = LEFT(gtin, length(gtin)-1);")
    conn.commit()

    cursor.execute("DROP TABLE IF EXISTS qb_totalitems;")
    stmt = """
            CREATE TABLE qb_totalitems AS SELECT 
            t.EPCs, 
            t.gtin, 
            t.DEPT_CATG_GRP_DESC,
            t.DEPT_CATEGORY_DESC, 
            t.VENDOR_NBR,
            t.VENDOR_NAME,
            t.BRAND_FAMILY_NAME,
            t.dept_nbr,
            qb.Item_Validation_Status
            FROM totalitems t
            INNER JOIN QB_IVS qb
            ON qb.UPC = t.UPC_No_Check;
        """
    cursor.execute(stmt)
    conn.commit()

    print(" -- Total Items Table creation complete.")
    df = sql.read_sql('SELECT * FROM qb_totalitems', conn)
    store.set_total_items(df)


def create_oh_data_sql():
    print(" -- Creating OH Data Table...")
    cursor.execute("DROP TABLE IF EXISTS OHData;")
    stmt = """
            CREATE TABLE OHData AS 
            SELECT DISTINCT gtin, ei_onhand_qty, dept_nbr, vendor_name, REPL_GROUP_NBR
            FROM itemfile
            WHERE dept_nbr IN ('7', '9', '14', '17', '20', '22', '71', '72', '74', '87');
    """
    cursor.execute(stmt)

    cursor.execute("ALTER TABLE OHData ADD COLUMN UPC_No_Check bigint AFTER REPL_GROUP_NBR;")
    cursor.execute("UPDATE OHData SET UPC_No_Check = LEFT(gtin, length(gtin)-1);")
    conn.commit()

    cursor.execute("DROP TABLE IF EXISTS QB_OHData;")
    stmt = """
            CREATE TABLE QB_OHData AS SELECT 
            o.gtin, o.ei_onhand_qty, o.dept_nbr, o.VENDOR_NAME, o.REPL_GROUP_NBR, qb.Item_Validation_Status
            FROM OHData o
            INNER JOIN QB_IVS qb
            ON qb.UPC = o.UPC_No_Check;
           """
    cursor.execute(stmt)
    conn.commit()


def create_oh_data_dept_sums_sql():
    global store
    cursor.execute("DROP TABLE IF EXISTS OHData_Dept_Sums;")

    stmt = """
            CREATE TABLE OHData_Dept_Sums AS 
            SELECT qb_ohdata.dept_nbr,
            SUM(qb_ohdata.ei_onhand_qty) AS Combined_ei_onhand_qty,
            SUM(CASE WHEN qb_ohdata.vendor_name LIKE "%IMPORT-%" THEN qb_ohdata.ei_onhand_qty ELSE 0 END) AS Import_ei_onhand_qty,
            SUM(CASE WHEN qb_ohdata.vendor_name NOT LIKE "%IMPORT-%" THEN qb_ohdata.ei_onhand_qty ELSE 0 END) AS Domestic_ei_onhand_qty
            FROM qb_ohdata 
            GROUP BY qb_ohdata.dept_nbr
            ORDER BY qb_ohdata.dept_nbr;
    """
    cursor.execute(stmt)
    stmt = """
            INSERT INTO OHData_Dept_Sums SELECT * FROM OHData_Dept_Sums
            UNION SELECT 0 dept_nbr, SUM(Combined_ei_onhand_qty), SUM(Import_ei_onhand_qty), SUM(Domestic_ei_onhand_qty) FROM OHData_Dept_Sums;
    """
    cursor.execute(stmt)
    cursor.execute("SET SQL_SAFE_UPDATES = 0;")
    cursor.execute("DELETE FROM OHData_Dept_Sums LIMIT 10;")
    cursor.execute("SET SQL_SAFE_UPDATES = 1;")
    conn.commit()
    print(" -- OH Data Table Created.")

    df = sql.read_sql('SELECT * FROM OHData_Dept_Sums', conn)
    store.set_expected(df)


def create_repl_breakdown_sql():
    global store
    print(" -- Creating REPL_GROUP_NBR_BREAKDOWN Table...")
    cursor.execute("DROP TABLE IF EXISTS REPL_GROUP_NBR_BREAKDOWN;")
    stmt = """
                    CREATE TABLE REPL_GROUP_NBR_BREAKDOWN (REPL_GROUP_NBR BIGINT NOT NULL);
            """
    cursor.execute(stmt)
    stmt = """
            INSERT INTO repl_group_nbr_breakdown (repl_group_nbr)
            SELECT DISTINCT itemfile.REPL_GROUP_NBR FROM itemfile
            WHERE dept_nbr IN ('7','9','14','17','20','22','71','72','74','87');
    """
    cursor.execute(stmt)
    cursor.fetchall()
    df = sql.read_sql('SELECT REPL_GROUP_NBR FROM REPL_GROUP_NBR_BREAKDOWN', conn)
    store.set_repl_nbr(df)


def new_store_prompt():
    global new_store_button
    new_store_button = customtkinter.CTkButton(master=main_frame, text="New Store", command=reset_interface)
    new_store_button.pack(anchor="ne", padx=10, pady=10)
    print("New Store? Or press 'Quit' to quit.")


def reset_interface():
    print("Resetting interface...")
    global new_store_button
    global store_list
    global cycle_count_paths
    global item_file_path
    global qb_master_path
    global store

    print("Store List1: ", store_list)
    store_list.append(store)
    print("Store List2: ", store_list)
    store = Store.Store(None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None,
                        None, None, None)
    new_store_button.destroy()
    store_entry.delete(0, END)
    date_entry.delete(0, END)
    cycle_count_paths = ""
    item_file_path = ""
    qb_master_path = ""
    store_num = ""
    date = ""


def export_weekly_report():
    matching_file_name = "WeeklyReport{1}.xlsx".format(store_num, date)
    path = os.path.join(os.path.expanduser("~"), "Downloads", matching_file_name)
    str(path)

    global writer
    global store
    print("Exporting Matching...")
    df = sql.read_sql('SELECT * FROM Matching', conn)
    store.set_matching(df)
    matching_sheet_name = "Matching {}".format(store_num)
    str(matching_sheet_name)

    print("Exporting Total Items...")
    df = sql.read_sql('SELECT * FROM qb_totalitems', conn)
    store.set_total_items(df)
    total_items_sheet_name = "Total Items {}".format(store_num)
    str(total_items_sheet_name)

    print("Exporting QB Matching...")
    df = sql.read_sql('SELECT * FROM QB_Matching', conn)
    store.set_qb_matching(df)
    qb_matching_sheet_name = "QB Matching {}".format(store_num)
    str(qb_matching_sheet_name)

    print("Exporting OH Data...")
    df = sql.read_sql('SELECT * FROM OHData_Dept_Sums', conn)
    store.set_expected(df)
    onhand_data_sheet_name = "Expected Items {}".format(store_num)
    str(onhand_data_sheet_name)

    print("Exporting REPL_GROUP_NBR_Breakdown")
    df = sql.read_sql('SELECT * FROM REPL_GROUP_NBR_Breakdown', conn)
    store.set_repl_nbr(df)
    repl_group_nbr_sheet_name = "REPL Breakdown {}".format(store_num)
    str(repl_group_nbr_sheet_name)

    conn.close()
    if (conn):
        conn.close()
        print("\nThe SQLite connection is closed.")


def generate_combined_reports(store_list):
    user = 'root'
    pw = ''
    host = 'localhost'
    port = 3306
    db = 'reportsystem'

    db_data = 'mysql+mysqldb://' + 'root' + ':' + 'password' + '@' + '127.0.0.1' + ':3306/' \
              + 'InStoreTracking' + '?charset=latin1'

    engine = create_engine(db_data)

    connection = mysql.connector.connect(user='root', password='password', host='127.0.0.1', database='InStoreTracking',
                                         allow_local_infile=True)

    cursor = connection.cursor(buffered=True)

    '''
        ------------Combined Matching Generation------------
    '''
    cursor.execute("DROP TABLE IF EXISTS CombinedMatching_Dupes;")
    stmt = "CREATE TABLE CombinedMatching_Dupes LIKE Matching;"
    cursor.execute(stmt)

    for store in store_list:
        store.get_matching().to_sql('CombinedMatching_Dupes', con=engine, if_exists='append', index=False)

    cursor.execute("DROP TABLE IF EXISTS CombinedMatching;")

    stmt = """
            CREATE TABLE CombinedMatching AS SELECT DISTINCT gtin,
            MAX(DEPT_CATG_GRP_DESC) AS DEPT_CATG_GRP_DESC,
            MAX(DEPT_CATEGORY_DESC) AS DEPT_CATEGORY_DESC,
            MAX(VENDOR_NBR) AS VENDOR_NBR,
            MAX(VENDOR_NAME) AS VENDOR_NAME,
            MAX(BRAND_FAMILY_NAME) AS BRAND_FAMILY_NAME,
            MAX(dept_nbr) AS dept_nbr,
            MAX(REPL_GROUP_NBR) AS REPL_GROUP_NBR,
            MAX(UPC_No_Check) AS UPC_No_Check
            FROM CombinedMatching_Dupes GROUP BY gtin;
    """
    cursor.execute(stmt)

    cursor.execute("DROP TABLE IF EXISTS CombinedMatching_Dupes;")

    '''
        ------------Combined QB Matching Generation------------
    '''
    cursor.execute("DROP TABLE IF EXISTS CombinedQBMatching_Dupes;")
    cursor.execute("CREATE TABLE CombinedQBMatching_Dupes LIKE QB_Matching;")

    for store in store_list:
        store.get_qb_matching().to_sql('CombinedQBMatching_Dupes', con=engine, if_exists='append', index=False)

    cursor.execute("DROP TABLE IF EXISTS CombinedQBMatching;")

    stmt = """
                CREATE TABLE CombinedQBMatching AS SELECT DISTINCT gtin,
                MAX(DEPT_CATG_GRP_DESC) AS DEPT_CATG_GRP_DESC,
                MAX(DEPT_CATEGORY_DESC) AS DEPT_CATEGORY_DESC,
                MAX(VENDOR_NBR) AS VENDOR_NBR,
                MAX(VENDOR_NAME) AS VENDOR_NAME,
                MAX(BRAND_FAMILY_NAME) AS BRAND_FAMILY_NAME,
                MAX(dept_nbr) AS dept_nbr,
                MAX(REPL_GROUP_NBR) AS REPL_GROUP_NBR,
                MAX(Item_Validation_Status) AS Item_Validation_Status
                
                FROM CombinedQBMatching_Dupes GROUP BY gtin;
        """
    cursor.execute(stmt)

    cursor.execute("DROP TABLE IF EXISTS CombinedQBMatching_Dupes;")

    '''
        ------------Combined REPL Generation------------
    '''
    stmt2 = "DROP TABLE IF EXISTS CombinedREPL_Dupes;"
    cursor.execute(stmt2)
    # stmt3 = "CREATE TABLE CombinedREPL_Dupes (id INT NOT NULL primary key auto_increment, " \
    #         "REPL_GROUP_NBR BIGINT NOT NULL, REPL_COUNT int);"
    # stmt3 = "CREATE TABLE CombinedREPL_Dupes (id INT NOT NULL primary key auto_increment, " \
    #         "REPL_GROUP_NBR BIGINT NOT NULL);"
    # cursor.execute(stmt3)
    stmt3 = "CREATE TABLE CombinedREPL_Dupes (REPL_GROUP_NBR BIGINT NOT NULL);"

    cursor.execute("DROP TABLE IF EXISTS CombinedREPL;")

    for store in store_list:
        store.get_repl_nbr().to_sql('CombinedREPL_Dupes', con=engine, if_exists='append', index=False)

    # cursor.execute("UPDATE CombinedREPL_Dupes SET REPL_Count=null;")
    stmt4 = "CREATE TABLE CombinedREPL AS SELECT DISTINCT repl_group_nbr FROM CombinedREPL_Dupes"
    cursor.execute(stmt4)

    # stmt = """
    #             SELECT REPL_GROUP_NBR, REPL_COUNT FROM CombinedREPL a
    #             UNION (SELECT 'Total', COUNT(REPL_GROUP_NBR) FROM CombinedREPL)
    #             ORDER BY REPL_COUNT DESC;
    #     """
    # cursor.execute(stmt)


    cursor.execute("DROP TABLE IF EXISTS CombinedREPL_Dupes")

    print("Gathering Combined Matching...")
    df14 = sql.read_sql('SELECT * FROM CombinedMatching', connection)
    store.set_combined(df14)
    print("Gathering Combined QB Matching...")
    df15 = sql.read_sql('SELECT * FROM CombinedQBMatching', connection)
    store.set_qb_combined(df15)
    print("Gathering Combined REPL...")
    df16 = sql.read_sql('SELECT * FROM CombinedREPL', connection)
    store.set_combined_repl(df16)

    engine.dispose()
    connection.close()


def entry_validation():
    return_value = True
    for func in [validate_store_input, validate_date_input, validate_cycle_count_paths, validate_item_file_path,
                 validate_qb_path]:
        if not func():
            print(f"{func.__name__} input is invalid or unspecified!")
            return_value = False
    return return_value


def submit_info():
    if entry_validation():
        print("Successfully submitted. Starting In Store Tracking Report Generation...")
        connect_to_mysql()
        epc_list_df = decodePreparation()
        epc_list, upc_list = decodeCycleCount(epc_list_df)
        import_cycle_count_sql(epc_list, upc_list)
        import_item_file_sql()
        import_qb_sql()
        create_matching_sql()
        create_qb_matching_sql()
        create_total_items_sql()
        create_oh_data_sql()
        create_oh_data_dept_sums_sql()
        create_repl_breakdown_sql()
        new_store_prompt()

    else:
        print("\n------------------------------------------------------------------------"
              "\n:: ERROR :: Invalid inputs! Please enter valid inputs before submitting!"
              "\n------------------------------------------------------------------------")


def quit_app():
    if (store.get_store_num() and store.get_date_input()
            and not store.get_matching().empty and not store.get_qb_matching().empty
            and not store.get_repl_nbr().empty and store.get_cycle() and store.get_item_file()
            and not store.get_total_items().empty and store not in store_list):
        store_list.append(store)
    if store_list:
        print("Store List: ", store_list)
        export_weekly_report()
        generate_combined_reports(store_list)

    print("Quit...")
    app.quit()


class InterfaceCreation:

    def __init__(self, root, w, h):
        self.root = root
        self.width = w
        self.height = h
        self.store_list = []
        self.store_num = None
        self.date_input = None
        self.store = Store.Store(None, None, None, None, None, None, None, None, None, None, None, None, None, None,
                                 None, None, None, None, None)
        self.folder_created = False
        global store
        store = Store.Store(None, None, None, None, None, None, None, None, None, None, None, None, None, None, None,
                            None, None, None, None)

    def updateInterface(self):
        self.store_list = store_list
        self.store_num = store_num
        self.date_input = date
        self.store = store

    customtkinter.set_appearance_mode("Dark")
    customtkinter.set_default_color_theme("dark-blue")
    global app
    app = customtkinter.CTk()
    app.title("In Store Tracking Report System")
    app.geometry("800x600")

    '''
    Frame Creation
    '''
    global main_frame
    main_frame = customtkinter.CTkFrame(master=app, fg_color="transparent")
    top_frame = customtkinter.CTkFrame(master=main_frame, fg_color="transparent")
    top_frame.configure(height=75)
    left_frame = customtkinter.CTkFrame(master=main_frame, fg_color="transparent")
    right_frame = customtkinter.CTkFrame(master=main_frame, fg_color="transparent")
    middle_frame = customtkinter.CTkFrame(master=main_frame, fg_color="transparent")
    center_frame = customtkinter.CTkFrame(master=middle_frame, fg_color="transparent")
    bottom_center_frame = customtkinter.CTkFrame(master=center_frame, fg_color="transparent")
    bottom_frame = customtkinter.CTkFrame(master=main_frame, fg_color="transparent", height=100)

    main_frame.pack(fill="both", expand=True)
    top_frame.pack(side=TOP, fill="x")
    bottom_frame.pack(side=BOTTOM, fill="x")
    middle_frame.pack(side=BOTTOM, fill="x")
    center_frame.pack(fill="y")
    bottom_center_frame.pack(side=BOTTOM, fill="x")
    left_frame.pack(side=LEFT, fill="both", expand=True)
    right_frame.pack(side=RIGHT, fill="both", expand=True)

    '''
    Store and Date Entry Creation
    '''
    global store_entry
    store_entry = customtkinter.CTkEntry(master=left_frame, placeholder_text="Store #")

    global date_entry
    date_entry = customtkinter.CTkEntry(master=right_frame, placeholder_text="Date (YYYY.MM.DD)")

    store_entry.pack(padx=30, pady=50)
    date_entry.pack(padx=30, pady=50)

    '''
    Button Creation
    '''
    cycle_count_button = customtkinter.CTkButton(master=center_frame, text="Cycle Counts", command=import_cycle_count)
    item_file_button = customtkinter.CTkButton(master=center_frame, text="Item File", command=import_item_file)
    qb_master_button = customtkinter.CTkButton(master=center_frame, text="QB Master Items",
                                               command=import_qb_master_items)
    submit_button = customtkinter.CTkButton(master=bottom_center_frame, text="Submit", command=submit_info)
    quit_button = customtkinter.CTkButton(master=bottom_center_frame, text="Quit", command=quit_app)

    cycle_count_button.pack(pady=5)
    item_file_button.pack(pady=15)
    qb_master_button.pack(pady=5)
    submit_button.pack(side=RIGHT, padx=10, pady=50)
    quit_button.pack(side=LEFT, padx=10)

    app.mainloop()
