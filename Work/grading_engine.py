import pandas as pd
from utils import GradingError

def run_grading(df, mappings, weights, max_marks, fail_col, fail_threshold, total_fail_threshold, enable_hard_fail, boundaries, grace_limits=None):
    df = df.copy()
    
    df["Total_Score"] = 0.0
    for comp, col in mappings.items():
        if col not in df.columns:
            raise GradingError(f"Mapped column {col} not found in dataset.")
        w = float(weights[comp]) / 100.0
        m = float(max_marks[comp])
        if m == 0: m = 1  # prevent division by zero
        # Proportional score: (raw_score / max) * (weight * 100)
        scaled_val = (pd.to_numeric(df[col], errors='coerce').fillna(0) / m) * (w * 100)
        df[f"Final {col}"] = scaled_val
        df["Total_Score"] += scaled_val
    
    mean_val = df["Total_Score"].mean()
    std_val = df["Total_Score"].std()
    
    if pd.isna(std_val) or std_val == 0:
        std_val = 1e-9 
        
    df["Z_Score"] = (df["Total_Score"] - mean_val) / std_val
    
    def assign_grade(row):
        if enable_hard_fail:
            if fail_col and fail_col in df.columns:
                try:
                    val = float(row[fail_col])
                    if val < fail_threshold:
                        return "F"
                except ValueError:
                    pass
            if row["Total_Score"] < total_fail_threshold:
                return "F"
                
        z = row["Z_Score"]
        sorted_grades = ["A", "A-", "B", "B-", "C", "C-", "D"]
        for g in sorted_grades:
            if z >= boundaries.get(g, -999):
                return g
        return "F"
        
    df["Grade"] = df.apply(assign_grade, axis=1)
    
    grade_order = ["A", "A-", "B", "B-", "C", "C-", "D", "F"]
    target_marks = {}
    for g in grade_order:
        if g == "F":
            target_marks["F"] = 0
        else:
            z_thresh = boundaries.get(g, -999)
            target_marks[g] = mean_val + (z_thresh * std_val)

    def get_gap(row):
        g = row["Grade"]
        if g == "A":
            return "Max", 0.0
        try:
            next_g_idx = grade_order.index(g) - 1
            if next_g_idx < 0:
                return "Max", 0.0
            next_g = grade_order[next_g_idx]
            
            target_m = target_marks[next_g]
            
            if enable_hard_fail and total_fail_threshold > target_m:
                target_m = total_fail_threshold
                
            gap = max(0.0, target_m - row["Total_Score"])
            return next_g, gap
        except Exception:
            return "N/A", 0.0

    res = df.apply(get_gap, axis=1, result_type="expand")
    df["Next_Grade"] = res[0]
    df["marks_for_next_grade"] = res[1]
    
    if "grace_limits" in locals() and grace_limits is not None and isinstance(grace_limits, dict) and any(v > 0 for v in grace_limits.values()):
        def apply_grace(row):
            gap = row["marks_for_next_grade"]
            next_g = row["Next_Grade"]
            if gap > 0 and next_g in grace_limits:
                limit = grace_limits.get(next_g, 0.0)
                if limit > 0 and gap <= limit:
                    return next_g
            return row["Grade"]
        df["Upgraded_Grade"] = df.apply(apply_grace, axis=1)

    return df

