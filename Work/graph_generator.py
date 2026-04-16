import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
import pandas as pd

def generate_hist(df):
    if df is None:
        return None
    cols = [c for c in df.columns if c.startswith("Final ")] + ["Total_Score"]
    cols = [c for c in cols if c in df.columns]
    if not cols:
        return None
        
    melted = df.melt(value_vars=cols, var_name="Evaluation", value_name="Score")
    
    plt.figure(figsize=(10, 5))
    sns.violinplot(x="Evaluation", y="Score", data=melted, inner="quartile", palette="muted")
    sns.swarmplot(x="Evaluation", y="Score", data=melted, color="white", alpha=0.5, size=4)
    
    plt.title("Detailed Component Score Distributions (Violin + Swarm)", color="white", fontsize=14)
    plt.xlabel("")
    plt.ylabel("Score", color="white")
    plt.xticks(rotation=15, color="gray")
    plt.yticks(color="gray")
    plt.gca().set_facecolor("none")
    plt.gcf().set_facecolor("none")
    plt.tight_layout()
    
    buf = BytesIO()
    plt.savefig(buf, format="png", facecolor="none", edgecolor="none")
    plt.close()
    buf.seek(0)
    return buf.read()

def generate_bar(df):
    if df is None or "Grade" not in df.columns or "Total_Score" not in df.columns:
        return None
        
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4))
    
    grade_order = ["A", "A-", "B", "B-", "C", "C-", "D", "F"]
    if "Upgraded_Grade" in df.columns:
        df_melt = pd.DataFrame({
            "Grade": list(df["Grade"]) + list(df["Upgraded_Grade"]),
            "Type": ["Original"] * len(df) + ["Upgraded"] * len(df)
        })
        sns.countplot(data=df_melt, x="Grade", hue="Type", order=grade_order, palette="viridis", ax=ax1)
    else:
        sns.countplot(data=df, x="Grade", order=grade_order, palette="viridis", ax=ax1)

    ax1.set_title("Grade Distribution", color="white")
    ax1.set_xlabel("Grade", color="white")
    ax1.set_ylabel("Count", color="white")
    ax1.tick_params(colors="gray")
    ax1.set_facecolor("none")
    
    sns.histplot(data=df, x="Total_Score", bins=10, color="#f39c12", kde=False, ax=ax2)
    ax2.set_title("Marks Range Distribution", color="white")
    ax2.set_xlabel("Total Score", color="white")
    ax2.set_ylabel("Students", color="white")
    ax2.tick_params(colors="gray")
    ax2.set_facecolor("none")
    
    fig.patch.set_facecolor("none")
    plt.tight_layout()
    
    buf = BytesIO()
    plt.savefig(buf, format="png", facecolor="none", edgecolor="none")
    plt.close()
    buf.seek(0)
    return buf.read()
