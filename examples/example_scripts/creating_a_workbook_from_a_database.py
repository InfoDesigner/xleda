from xleda import wb

# <your database goes here>
sqlite_db = "https://github.com/InfoDesigner/xleda/raw/refs/heads/main/examples/data/chinook.db"


# Creates "Chinook.xlsm" in the current directory with 11 dataframes
wb(data=sqlite_db,
   file_name="Chinook")