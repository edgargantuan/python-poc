import pyodbc

db_driver = '{ODBC Driver 17 for SQL Server}'
db_server = "****************************"
db_name = "***********"
db_uid = "********"
db_pwd = "******"

connect_string = f"DRIVER={db_driver};SERVER={db_server};DATABASE={db_name};UID={db_uid};PWD={db_pwd}"

conn = pyodbc.connect("DRIVER={ODBC Driver 17 for SQL Server};SERVER=dpie-eplanning-dev-srvr.database.windows.net;DATABASE=dpie-eplanning-dev-db;UID=dpieeplanningdevsrvradmin;PWD=rules@123")