# daily_data_udf.py (updated)
import sqlite3
import configparser
import logging
import logging.handlers
import functools
import time
import xlwings as xw
import os

# --- Global Configuration and Initialization ---
CONFIG = {}
LOGGER = None
VALID_FIELDS = None  # Will be populated dynamically based on DB table

def load_config():
    """
    Reads configuration from config.ini.
    If the file isn't found in working dir, uses defaults.
    Also populates VALID_FIELDS by inspecting the DB table columns.
    """
    global CONFIG, VALID_FIELDS
    config = configparser.ConfigParser(interpolation=None)
    config_read_files = config.read('config.ini')
    
    if not config_read_files:
        print("Warning: config.ini not found. Using defaults.")
        config['DATABASE'] = {'DB_PATH': 'mcap.db', 'TABLE_NAME': 'mcap', 'DATE_FORMAT': '%Y-%m-%d'}
        config['LOGGING'] = {'LOG_FILE': 'query_log.txt', 'MAX_BYTES': '1048576', 'BACKUP_COUNT': '5'}
    
    CONFIG = config

    # Try to populate VALID_FIELDS by reading the DB table schema
    try:
        db_path = CONFIG['DATABASE']['DB_PATH']
        table = CONFIG['DATABASE']['TABLE_NAME']
        if os.path.exists(db_path):
            conn = sqlite3.connect(db_path)
            cur = conn.cursor()
            cur.execute(f"PRAGMA table_info({table})")
            cols = cur.fetchall()  # (cid, name, type, notnull, dflt_value, pk)
            VALID_FIELDS = {row[1] for row in cols}
            conn.close()
            if not VALID_FIELDS:
                # fallback to a safe default set if table not present or PRAGMA returned empty
                VALID_FIELDS = {"accord_code", "company_name", "sector", "mcap_category", "date", "mcap"}
        else:
            # If DB not found, set conservative default fields and let runtime errors surface in logs
            VALID_FIELDS = {"accord_code", "company_name", "sector", "mcap_category", "date", "mcap"}
    except Exception as e:
        print(f"Warning: could not populate VALID_FIELDS from DB: {e}")
        VALID_FIELDS = {"accord_code", "company_name", "sector", "mcap_category", "date", "mcap"}

def setup_logging():
    """
    Sets up the rotating file log handler as required (max 1 MB).
    """
    global LOGGER
    if LOGGER is not None:
        return
        
    try:
        log_config = CONFIG['LOGGING']
        log_file = log_config['LOG_FILE']
        max_bytes = int(log_config['MAX_BYTES'])
        backup_count = int(log_config['BACKUP_COUNT'])

        LOGGER = logging.getLogger('QueryLogger')
        LOGGER.setLevel(logging.INFO)

        formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')

        handler = logging.handlers.RotatingFileHandler(log_file, maxBytes=max_bytes, backupCount=backup_count)
        handler.setFormatter(formatter)
        
        # Ensure only one handler is attached
        if not LOGGER.handlers:
            LOGGER.addHandler(handler)
    except Exception as e:
        print(f"Error setting up logging: {e}")
        LOGGER = logging.getLogger('FallbackLogger')

def _connect_db():
    """Establishes connection to the SQLite database using path from config."""
    db_path = CONFIG['DATABASE']['DB_PATH']
    return sqlite3.connect(db_path)

@functools.lru_cache(maxsize=128)
def _execute_query(sql, params=()):
    """
    Executes a parameterized query. This function is cached.
    Params must be a tuple (hashable) for the LRU cache.
    Returns results as a tuple of tuples (immutable), suitable for caching.
    """
    # Ensure params are a tuple (hashable)
    if not isinstance(params, tuple):
        params = tuple(params)
    conn = None
    try:
        conn = _connect_db()
        cursor = conn.cursor()
        cursor.execute(sql, params)
        rows = cursor.fetchall()
        # Convert to tuple of tuples for immutability / safe caching
        return tuple(tuple(r) for r in rows)
    except sqlite3.Error as e:
        raise RuntimeError(f"Database error: {e}")
    finally:
        if conn:
            conn.close()

def log_udf_call(func_name, params, start_time, status, error_msg=""):
    """Records the function call to query_log.txt with execution time and returns status/value for UDF to return."""
    elapsed_time_ms = (time.perf_counter() - start_time) * 1000
    param_str = ', '.join(map(str, params))
    log_message = f"{func_name} | P: ({param_str}) | Time: {elapsed_time_ms:.2f}ms | Status: {status}"
    if error_msg:
        log_message += f" | Error: {error_msg}"
        if LOGGER:
            LOGGER.error(log_message)
        # Return Excel-friendly error tag
        return "#QUERY_ERROR"
    else:
        if LOGGER:
            LOGGER.info(log_message)
        return True

# --- Excel UDFs ---

@xw.func
def get_daily_data(accord_code, field, date):
    start_time = time.perf_counter()
    func_name = 'get_daily_data'
    table = CONFIG['DATABASE']['TABLE_NAME']
    params = (accord_code, field, date)

    # Validate inputs
    if not isinstance(accord_code, (int, float)) or not isinstance(field, str) or not isinstance(date, str):
        return log_udf_call(func_name, params, start_time, 'FAILURE', 'Invalid input type.')

    # Validate field against whitelist
    if field not in VALID_FIELDS:
        return log_udf_call(func_name, params, start_time, 'FAILURE', f'Invalid field: {field}')

    sql = f"SELECT {field} FROM {table} WHERE accord_code = ? AND date = ?"
    query_params = (int(accord_code), date)

    try:
        results = _execute_query(sql, query_params)
        if not results:
            return log_udf_call(func_name, params, start_time, 'FAILURE', 'Data not found.')
        value = results[0][0]
        log_udf_call(func_name, params, start_time, 'SUCCESS')
        return value
    except RuntimeError as e:
        return log_udf_call(func_name, params, start_time, 'FAILURE', str(e))
    except Exception as e:
        return log_udf_call(func_name, params, start_time, 'FAILURE', f"Unexpected error: {e}")

@xw.func
def get_series(accord_code, field, start_date, end_date):
    start_time = time.perf_counter()
    func_name = 'get_series'
    table = CONFIG['DATABASE']['TABLE_NAME']
    params = (accord_code, field, start_date, end_date)

    if not isinstance(accord_code, (int, float)) or not isinstance(field, str):
        return [["#INPUT_ERROR"]]

    if field not in VALID_FIELDS:
        return [["#INVALID_FIELD"]]

    sql = f"""
        SELECT date, {field}
        FROM {table}
        WHERE accord_code = ? AND date BETWEEN ? AND ?
        ORDER BY date
    """
    query_params = (int(accord_code), start_date, end_date)

    try:
        results = _execute_query(sql, query_params)
        if not results:
            log_udf_call(func_name, params, start_time, 'FAILURE', 'Data not found.')
            return [["#N/A_DATA", ""]]
        header = [['Date', field]]
        # convert tuples back to lists for Excel spill
        output = header + [list(r) for r in results]
        log_udf_call(func_name, params, start_time, 'SUCCESS')
        return output
    except RuntimeError as e:
        return [[log_udf_call(func_name, params, start_time, 'FAILURE', str(e))]]
    except Exception as e:
        return [[log_udf_call(func_name, params, start_time, 'FAILURE', f"Unexpected error: {e}")]]

@xw.func
def get_daily_matrix(date, field):
    start_time = time.perf_counter()
    func_name = 'get_daily_matrix'
    table = CONFIG['DATABASE']['TABLE_NAME']
    params = (date, field)

    if not isinstance(date, str) or not isinstance(field, str):
        return [["#INPUT_ERROR"]]

    if field not in VALID_FIELDS:
        return [["#INVALID_FIELD"]]

    sql = f"""
        SELECT accord_code, company_name, sector, mcap_category, {field}
        FROM {table}
        WHERE date = ?
    """
    query_params = (date,)

    try:
        results = _execute_query(sql, query_params)
        if not results:
            log_udf_call(func_name, params, start_time, 'FAILURE', 'Data not found.')
            return [["#N/A_DATA", "", "", "", ""]]
        header = [['accord_code', 'company_name', 'sector', 'mcap_category', field]]
        output = header + [list(r) for r in results]
        log_udf_call(func_name, params, start_time, 'SUCCESS')
        return output
    except RuntimeError as e:
        return [[log_udf_call(func_name, params, start_time, 'FAILURE', str(e))]]
    except Exception as e:
        return [[log_udf_call(func_name, params, start_time, 'FAILURE', f"Unexpected error: {e}")]]


@xw.func
def get_all_mcap(accord_code, field):
    start_time = time.perf_counter()
    func_name = 'get_all_mcap'
    table = CONFIG['DATABASE']['TABLE_NAME']
    params = (accord_code, field)

    if not isinstance(accord_code, (int, float)) or not isinstance(field, str):
        return [["#INPUT_ERROR"]]

    if field not in VALID_FIELDS:
        return [["#INVALID_FIELD"]]

    sql = f"""
        SELECT date, {field}
        FROM {table}
        WHERE accord_code = ?
        ORDER BY date
    """
    query_params = (int(accord_code),)

    try:
        results = _execute_query(sql, query_params)
        if not results:
            log_udf_call(func_name, params, start_time, 'FAILURE', 'Data not found.')
            return [["#N/A_DATA", ""]]
        header = [['Date', field]]
        output = header + [list(r) for r in results]
        log_udf_call(func_name, params, start_time, 'SUCCESS')
        return output
    except RuntimeError as e:
        return [[log_udf_call(func_name, params, start_time, 'FAILURE', str(e))]]
    except Exception as e:
        return [[log_udf_call(func_name, params, start_time, 'FAILURE', f"Unexpected error: {e}")]]


# --- Initialization on script load ---
if __name__ == '__main__':
    load_config()
    setup_logging()
    print("UDF module loaded. Ready for use in Excel with xlwings.")
else:
    load_config()
    setup_logging()
