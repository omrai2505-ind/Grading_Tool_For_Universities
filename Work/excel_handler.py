import pandas as pd
from utils import GradingError

def read_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        if df.empty:
            raise GradingError("The uploaded Excel file is empty.")
        return df
    except Exception as e:
        raise GradingError(f"Failed to read Excel file: {e}")
