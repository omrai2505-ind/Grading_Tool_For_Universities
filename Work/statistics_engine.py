def calculate_stats(df, score_col="Total_Score"):
    if df is None or score_col not in df.columns or len(df) == 0:
        return {"mean": 0, "std": 0, "max": 0, "min": 0, "count": 0}
        
    stats = {
        "mean": df[score_col].mean(),
        "std": df[score_col].std(),
        "max": df[score_col].max(),
        "min": df[score_col].min(),
        "count": len(df)
    }
    return stats
