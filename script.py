import argparse
import logging
import os
import shutil

import pandas as pd
import pyodbc

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler("debug.log"), logging.StreamHandler()],
)


def duplicateDatabaseFile(fp):
    fp = shutil.copyfile(fp, os.path.join(os.path.dirname(fp), "extracted.mdb"))
    return fp


def openDatabase(fp):
    conn = pyodbc.connect(
        f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={fp};",
        autocommit=True,  # Required to bypass MaxLocksPerFile
    )
    return conn


def getBaseViewTable(conn):
    df = pd.read_sql("SELECT * FROM Results", con=conn)
    df.drop(columns=["LockedBy", "Version", "Signature"], inplace=True)
    return df


def dropXinshaResultIndexes(conn, additional_result_indexes=[]):
    cursor = conn.cursor()
    cursor.execute(
        "SELECT ResultIndex FROM Results WHERE Description LIKE '%x-%' OR Description LIKE '%X-%'"
    )
    result_indexes = [row.ResultIndex for row in cursor.fetchall()]
    logging.debug(
        f"All Xinsha result indexes (total {len(result_indexes)}): {result_indexes}"
    )
    # This will further filter out additional result indexes defined in optional argument.
    if additional_result_indexes:
        result_indexes.extend(additional_result_indexes)
        result_indexes = list(dict.fromkeys(result_indexes))  # Remove duplicates
        logging.debug(
            f"All result indexes to be removed (total {len(result_indexes)}): {result_indexes}"
        )
    result_indexes_tuples = [(result_index,) for result_index in result_indexes]
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
        logging.debug(f"Deleting Xinsha result indexes from {table}...")
        for i in range(0, len(result_indexes_tuples), batch_size):
            batch = result_indexes_tuples[i : i + batch_size]
            cursor.executemany(f"DELETE FROM {table} WHERE ResultIndex = ?", batch)
    return True


def main():
    # Parse Script Arguments
    parser = argparse.ArgumentParser(
        description="""
        Filter Xinsha tests by job reference in description and extract all other ICP test results.
        Script will look for 'X' in description field in results table.
        """
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
        logging.debug(
            "ICP Results database file path does not exist! Please check file path."
        )
        return
    fp = duplicateDatabaseFile(fp)
    try:
        conn = openDatabase(fp)
    except Exception as e:
        logging.debug(e)
        return

    # Drop result index for xinsha records
    status = dropXinshaResultIndexes(conn)
    if status:
        logging.debug(
            f"Access Database filtered and extracted successfully. Extracted database file path is: {fp}"
        )
    return


if __name__ == "__main__":
    main()
