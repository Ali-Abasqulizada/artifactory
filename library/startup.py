import os
import sys
import logging
import sqlite3


def check_db_exists_or_fail(path):
    """
    Checks if the database at the path exists. Exits with a non-successful status if the db does not exist at the
    provided path. This causes the backend to exit. Call this function only during startup.
    :param path:    the path to the database file.
    :return:        no return, exits if the path is not found
    """
    if not os.path.exists(path):
        error_message = f"Database file '{path}' does not exist"
        logging.getLogger().error(f"error_message: '{error_message}' | path: '{path}'")
        sys.exit(1)


def create_folder(path):
    """
    Recursively creates a folder specified by path. This method is supposed to be used
    only on startup. It always succeeds.
    :param path:  the path to folder to be created
    :return:      does not return value and is never failing.
    """
    os.makedirs(path, exist_ok=True)


def connect_db(database="artifactory.db"):
    try:
        conn = sqlite3.connect(database)
        conn.execute("pragma journal_mode=WAL")
        conn.row_factory = sqlite3.Row
        return conn, ""
    except sqlite3.Error as err:
        error_message = "Failed to connect to the database"
        logging.getLogger().error(
            f"error_message: '{error_message}' | database: '{database}' | error: {err}"
        )
        return None, err
