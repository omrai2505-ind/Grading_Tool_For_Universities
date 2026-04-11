import pandas as pd
from utils import GradingError

def run_grading(df, mappings, weights, fail_col, fail_threshold, enable_hard_fail, boundaries):
    df = df.copy()
    
    df["Total_Score"] = 0.0
    for comp, col in mappings.items():
        if col not in df.columns:
            raise GradingError(f"Mapped column {col} not found in dataset.")
        w = float(weights[comp]) / 100.0
        df["Total_Score"] += pd.to_numeric(df[col], errors='coerce').fillna(0) * (w * 100)
    
    mean_val = df["Total_Score"].mean()
    std_val = df["Total_Score"].std()
    
    if pd.isna(std_val) or std_val == 0:
        std_val = 1e-9 
        
    df["Z_Score"] = (df["Total_Score"] - mean_val) / std_val
    
    def assign_grade(row):
        if enable_hard_fail and fail_col and fail_col in df.columns:
            try:
                val = float(row[fail_col])
                if val < fail_threshold:
                    return "F"
            except ValueError:
                pass
                
        z = row["Z_Score"]
        sorted_grades = ["A", "A-", "B", "B-", "C", "C-", "D"]
        for g in sorted_grades:
            if z >= boundaries.get(g, -999):
                return g
        return "F"
        
    df["Grade"] = df.apply(assign_grade, axis=1)
    
    return df
