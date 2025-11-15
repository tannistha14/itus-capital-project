# ITUS Capital Data Analytics Internship Project

## Project Title: Excel-Integrated Financial Data Retrieval using SQLite

This project delivers a Python-based Excel User Defined Function (UDF) system that connects Excel directly to a local SQLite database for retrieving specific financial data (like mcap, sector) for any company on any given date. The system is designed for high performance, utilizing caching and database indexing.

---

###  Key Features

* **Excel UDF Integration**: Exposes four key data retrieval functions to Excel formulas.
* **Database Integration**: Uses `sqlite3` to connect to a local `.sqlite` file and executes parameterized queries for safety.
* **Performance Optimization**: Implements **LRU Caching** and utilizes a database index on `(accord_code, date)` to ensure a response time per query of less than 0.05s.
* **Error Handling & Logging**: UDFs return meaningful errors for invalid inputs. All function calls, parameters, execution times, and success/failure statuses are recorded in `query_log.txt`.

---

### Setup Instructions

To run this project, you will need **Python** (version 3.7+) and the following libraries:

1.  **Install Required Libraries**:
    ```bash
    pip install xlwings configparser
    ```

2.  **File Placement**: Place all project deliverables in the same directory:
    * `daily_data_udf.py`
    * `config.ini`
    * `mcap.db` (The provided SQLite file)
    * `schema.sql`
    * `example.xlsx`

3.  **Database Indexing**: Ensure the index is created in your database. Run the contents of `schema.sql` against your `mcap.db` file to create the required composite index on `(accord_code, date)`.

4.  **Install xlwings Add-in**: If this is your first time using `xlwings` for UDFs, install the Excel add-in:
    ```bash
    xlwings addin install
    ```

---

###  Usage in Excel (`example.xlsx`)

To use the functions, you must first start the UDF server from your project directory:

```bash
python daily_data_udf.py