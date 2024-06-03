import argparse
import os
import shutil

import pyodbc


def duplicateDatabaseFile(fp):
    fp = shutil.copyfile(fp, os.path.join(os.path.dirname(fp), "extracted.mdb"))
    return fp


def openDatabase(fp):
    conn = pyodbc.connect(
        f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={fp};",
        autocommit=True,  # Required to bypass MaxLocksPerFile
    )
    # for table in conn.cursor().tables():
    #     table_name = table.table_name
    #     # print(table_name)
    return conn


def drop_xinsha_result_indexes(conn):
    cursor = conn.cursor()
    cursor.execute(
        "SELECT ResultIndex FROM Results WHERE Description LIKE '%x%' OR Description LIKE '%X%'"
    )
    result_indexes = [row.ResultIndex for row in cursor.fetchall()]
    result_indexes_tuples = [(result_index,) for result_index in result_indexes]
    print(f"All Xinsha result indexes (total {len(result_indexes)}): {result_indexes}")
    table_list = [
        "Batch",
        "Ccstds",
        "Ccurves",
        # "LimitRemarks",
        "LogMean",
        "LogRepl",
        "LogSampinfo",
        "Mean",
        "Methods",
        "Repl",
        "Results",
        "Sampinfo",
        "Universal",
    ]
    batch_size = 100
    for table in table_list:
        print(f"Deleting Xinsha result indexes from {table}...")
        for i in range(0, len(result_indexes_tuples), batch_size):
            batch = result_indexes_tuples[i : i + batch_size]
            cursor.executemany(f"DELETE FROM {table} WHERE ResultIndex = ?", batch)
    return True


def main():

    # Parse Script Arguments
    parser = argparse.ArgumentParser(
        description="Filter Xinsha tests by job reference in description and extract all other ICP test results. Script will look for 'X' in description field in results table."
    )
    parser.add_argument(
        "-f",
        "--filepath",
        help="File path of ICP results Microsoft Access Database",
        required=True,
    )
    args = parser.parse_args()

    # Check if file path exists and test opening database connection
    fp = args.filepath
    if not os.path.exists(fp):
        print("ICP Results database file path does not exist! Please check file path.")
        return
    fp = duplicateDatabaseFile(fp)
    try:
        conn = openDatabase(fp)
    except Exception as e:
        print(e)
        return

    # Drop result index for xinsha records
    status = drop_xinsha_result_indexes(conn)
    if status:
        print(
            f"Access Database filtered and extracted successfully. Extracted database file path is: {fp}"
        )
    return


if __name__ == "__main__":
    main()
