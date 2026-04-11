import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO

def generate_hist(df):
    if df is None or "Total_Score" not in df.columns:
        return None
    plt.figure(figsize=(6, 4))
    sns.histplot(df["Total_Score"], kde=True, bins=15, color="#00b4d8")
    plt.title("Distribution of Total Scores", color="white")
    plt.xlabel("Score", color="white")
    plt.ylabel("Frequency", color="white")
    plt.tick_params(colors="gray")
    plt.gca().set_facecolor("transparent")
    plt.gcf().set_facecolor("transparent")
    plt.tight_layout()
    
    buf = BytesIO()
    plt.savefig(buf, format="png", transparent=True)
    plt.close()
    buf.seek(0)
    return buf.read()

def generate_bar(df):
    if df is None or "Grade" not in df.columns:
        return None
    plt.figure(figsize=(6, 4))
    grade_order = ["A", "A-", "B", "B-", "C", "C-", "D", "F"]
    sns.countplot(x="Grade", data=df, order=grade_order, palette="viridis")
    plt.title("Grade Distribution", color="white")
    plt.xlabel("Grade", color="white")
    plt.ylabel("Count", color="white")
    plt.tick_params(colors="gray")
    plt.gca().set_facecolor("transparent")
    plt.gcf().set_facecolor("transparent")
    plt.tight_layout()
    
    buf = BytesIO()
    plt.savefig(buf, format="png", transparent=True)
    plt.close()
    buf.seek(0)
    return buf.read()
