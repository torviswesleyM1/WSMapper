# WrightSoft Part Mapping Integration Tool
# Author: Torvis Wesley
# Date: January 16, 2025
# Version: 1.0.0
# Description:
# This script updates WrightSoft's local database with a default part mapping scheme
# to support the ECI Bolt Enterprise integreation with Mechanical One' HVAC Estimating workflows.
# It automates mapping WrightSoft's generic parts to custom M1 parts enabling quick BOM creation.


import os
import sqlite3
import csv
import os

# Global Constants
DATABASE_PATH = "C:\ProgramData\Wrightsoft HVAC\Data\RPRUWSF.mdb"
LOG_FILE = "wrightsoft_part_mapprt_log"

# Logging setup
def log_message(message):
       with open(LOG_FILE, "a") as log_file:
            log_file.write(f"{message}\n")

# Database connection setup
def connect_to_database():
    try:
        connection = sqlite3.connect(DATABASE_PATH)
        log_message("Connected to the WrightSoft database")
        return connection
    except sqlite3.Error as e:
        log_message(f"Erro connecting to database: {e}")
        raise

# function to execute SQL query with parameters
def execute_query(connection, query, params):
    try:
        cursor = connection.cursor()
        cursor.execute(query, params)
        connection.commit()
        return cursor
    except sqlite3.Error as e:
            log_message(f"SQL Execution Error: {e}")
            raise
# Function to map parts
def map_parts(csv_file):
    connection = connect_to_database()

    try:
        with open(csv_file, mode="r") as file:
            csv_reader = csv.DictReader(file)

            for row in csv_reader:
                query = """
                INSERT INTO PartMapping (GenItem, PartAssemStr, RecOrigin)
                VALUES (?, ?, ?)
                ON CONFLICT(GenItem) DO UPDATE SET PartAssemStr = excluded.partAssemStr,
                                                                    RecOrgin = excluded.RecOrigin;
                """
                log_message(f"Mapped Genitem: {row['GenItem']} -> {row['PartAssemStr']}")
    except FileNotFoundError:
        log_message(f"CSV file not found: {csv_file}")
        raise
    finally:
        connection.close()
        log_message("Database connection closed.")

# main function
def main():
    csv_file = "Y:\\Estimating\\_Estimating Resource Material\\EstimatingDatadatabase\\_WS_Database Import\\_WS_Setup\\PartPhase.csv"

    if not os.path.exists(csv_file):
        log_message("CSV file does not exist. Exiting...")
        return

    log_message("Starting WrightSoft Part Mapping Integration Tool")

    try:
        map_parts(csv_file)
        log_message("Part mapping completed successfully.")
    except Exception as e:
        log_message(f"Error occurred: {e}")

if __name__ == "__main__":
    main()