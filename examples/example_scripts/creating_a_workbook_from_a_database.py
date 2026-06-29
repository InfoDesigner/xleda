from xleda import wb
from pathlib import Path

sqlite_db = Path(__file__).parent.parent / 'data' / r"chinook.db"

# Creates "Chinook.xlsm" in the current directory with 11 dataframes
wb(data=sqlite_db,
   file_name="Chinook")