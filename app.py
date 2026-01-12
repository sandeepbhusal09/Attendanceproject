import os
import sys
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from flask import Flask, render_template, request, send_file
import nepali_datetime
import webbrowser
from threading import Timer

# Support for local libs folder
libs_path = os.path.join(os.path.dirname(__file__), 'libs')
if os.path.exists(libs_path):
    sys.path.insert(0, libs_path)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, template_folder=BASE_DIR)



# Config
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

def open_browser():
    webbrowser.open("http://127.0.0.1:5000")

def bs_to_ad(date_val):
    try:
        if pd.isna(date_val): return None
        d_str = str(date_val).replace("/", "-").replace(".", "-")
        y, m, d = map(int, d_str.split("-"))
        return nepali_datetime.date(y, m, d).to_datetime_date()
    except: return None

def find_header_row(file_path):
    """Peek into Excel to find the header row dynamically."""
    try:
        peek = pd.read_excel(file_path, header=None, nrows=30)
        for index, row in peek.iterrows():
            vals = [str(v).strip().lower() for v in row.values]
            if 'id' in vals and 'date' in vals: return index
    except: pass
    return 0

@app.route("/", methods=["GET", "POST"])
def index():
    results = None
    stats = {}

    if request.method == "POST":
        files = request.files.getlist("files")
        all_dfs = []
        if not files or files[0].filename == '': return render_template("index.html")

        for f in files:
            if f.filename.endswith((".xlsx", ".xls")):
                path = os.path.join(UPLOAD_DIR, f.filename)
                f.save(path)
                skip = find_header_row(path)
                df = pd.read_excel(path, skiprows=skip)
                df.columns = [str(c).strip() for c in df.columns]
                all_dfs.append(df)

        if all_dfs:
            combined = pd.concat(all_dfs, ignore_index=True)
            combined['date_ad'] = combined['Date'].apply(bs_to_ad)
            combined['dt'] = pd.to_datetime(combined['date_ad'].astype(str) + " " + combined['Time'].astype(str), errors="coerce")
            combined['type'] = combined['Attendance Check Point'].apply(lambda x: 'enter' if 'outside' in str(x).lower() else 'exit')
            combined = combined.dropna(subset=['dt']).sort_values(['ID', 'dt'])
            
            summary = []
            for emp_id, g in combined.groupby("ID"):
                name = g['First Name'].iloc[0] if 'First Name' in g else "Unknown"
                working_days = g['date_ad'].nunique()
                total_mins = 0
                for _, day_data in g.groupby(combined['dt'].dt.date):
                    day_data = day_data.sort_values("dt")
                    for i in range(len(day_data)-1):
                        curr, nxt = day_data.iloc[i], day_data.iloc[i+1]
                        if curr['type'] == "exit" and nxt['type'] == "enter":
                            diff = (nxt['dt'] - curr['dt']).total_seconds() / 60
                            if 0 < diff < 600: total_mins += diff

                summary.append({"ID": emp_id, "Name": name, "Days": working_days, "Mins": round(total_mins, 2), "Hrs": round(total_mins/60, 2)})

            res_df = pd.DataFrame(summary)
            avg_h = res_df["Hrs"].mean()
            stats = {"total": len(res_df), "avg": round(avg_h, 2), "max_name": res_df.loc[res_df['Hrs'].idxmax()]['Name'], "max_val": res_df['Hrs'].max()}

            res_df.sort_values("ID").to_csv(os.path.join(OUTPUT_DIR, "summary.csv"), index=False)

            # --- PROFESSIONAL VERTICAL GRAPH ---
            # Sort Highest to Lowest
            res_plot = res_df.sort_values("Hrs", ascending=False)
            plt.style.use('default')
            
            # Dynamic width: Graph expands as you add more employees
            dynamic_width = max(12, len(res_plot) * 0.5) 
            fig, ax = plt.subplots(figsize=(dynamic_width, 8), dpi=150)
            
            # Bar Chart
            bars = ax.bar(res_plot["ID"].astype(str), res_plot["Hrs"], color='#4f46e5', edgecolor='white', alpha=0.9, width=0.6)
            
            # Unit Average line
            ax.axhline(avg_h, color='#f43f5e', linestyle='--', linewidth=1.5, label=f'Avg: {stats["avg"]}h')

            # Data labels on top
            for bar in bars:
                h = bar.get_height()
                if h > 0:
                    ax.annotate(f'{h}', xy=(bar.get_x() + bar.get_width()/2, h),
                                xytext=(0, 5), textcoords="offset points", ha='center', va='bottom', 
                                fontsize=8, fontweight='bold', color='#1e293b')

            ax.set_title("Employee Outside Time Breakdown", fontsize=18, fontweight='800', pad=30)
            ax.set_ylabel("Total Hours Outside", fontweight='bold', color='#64748b')
            ax.set_xlabel("Employee ID", fontweight='bold', color='#64748b')
            
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.yaxis.grid(True, linestyle='--', alpha=0.3)
            
            plt.xticks(rotation=90, fontsize=9)
            ax.legend(loc='upper right', frameon=False)
            plt.tight_layout()
            
            plt.savefig(os.path.join(OUTPUT_DIR, "graph.png"), bbox_inches='tight')
            plt.close()

            results = res_df.sort_values("ID").to_dict(orient="records")

    return render_template("index.html", results=results, stats=stats)

@app.route("/download/<f>")
def download(f):
    t = "summary.csv" if f == "csv" else "graph.png"
    return send_file(os.path.join(OUTPUT_DIR, t), as_attachment=True)

if __name__ == "__main__":
    Timer(1.5, open_browser).start()
    app.run(port=5000, debug=False)