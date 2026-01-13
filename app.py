import os
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from flask import Flask, render_template, request, send_file
import nepali_datetime
import webbrowser
from threading import Timer
import matplotlib.patches as mpatches

# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, template_folder=BASE_DIR)

# Config
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Employee info
emp_info = {
    160155: {"Name": "tara", "Designation": "Administrative Officer"},
    137209: {"Name": "Janardan", "Designation": "Assistant Manager"},
    175475: {"Name": "PRABHAT", "Designation": "Assistant Manager"},
    118422: {"Name": "Prabina", "Designation": "Assistant Manager"},
    11820: {"Name": "pramod", "Designation": "Assistant Manager"},
    194948: {"Name": "Praveen", "Designation": "Assistant Manager"},
    170067: {"Name": "Rajesh", "Designation": "Assistant Manager"},
    174215: {"Name": "Ram", "Designation": "Assistant Manager"},
    133312: {"Name": "Sarita", "Designation": "Assistant Manager"},
    170176: {"Name": "Saroj", "Designation": "Assistant Manager"},
    152218: {"Name": "Saugat", "Designation": "Assistant Manager"},
    156981: {"Name": "Sunil", "Designation": "Assistant Manager"},
    197099: {"Name": "Rubal", "Designation": "Deputy Manager"},
    152187: {"Name": "Subarna", "Designation": "Deputy Manager"},
    187058: {"Name": "SUMAN", "Designation": "Deputy Manager"},
    193531: {"Name": "Chandan", "Designation": "Director"},
    130165: {"Name": "bijay", "Designation": "Driver"},
    2057: {"Name": "Pema", "Designation": "Driver"},
    789: {"Name": "Rajib", "Designation": "Driver"},
    192576: {"Name": "Anil", "Designation": "Engineer"},
    128538: {"Name": "Bimal", "Designation": "Engineer"},
    170534: {"Name": "Dhanjit", "Designation": "Engineer"},
    176805: {"Name": "Hiralal", "Designation": "Engineer"},
    183539: {"Name": "Neha", "Designation": "Engineer"},
    170276: {"Name": "Purna", "Designation": "Engineer"},
    142832: {"Name": "RAJAN", "Designation": "Engineer"},
    128599: {"Name": "RISHAB", "Designation": "Engineer"},
    197718: {"Name": "Saru", "Designation": "Engineer"},
    177988: {"Name": "Sujeet", "Designation": "Engineer"},
    116608: {"Name": "Suman", "Designation": "Engineer"},
    128345: {"Name": "Suman", "Designation": "Engineer"},
    197768: {"Name": "Sumit", "Designation": "Engineer"},
    149782: {"Name": "Yagya", "Designation": "Foreman Driver"},
    182274: {"Name": "Mohan", "Designation": "Manager"},
    160381: {"Name": "Januka", "Designation": "Office Helper"},
    188484: {"Name": "Mina", "Designation": "Office Helper"},
    121865: {"Name": "UMAKANTA", "Designation": "Senior Operator"},
    161687: {"Name": "Bhawani", "Designation": "Others"},
    170702: {"Name": "Durga", "Designation": "Others"},
    192631: {"Name": "Gopal", "Designation": "Others"},
    456: {"Name": "HARKA", "Designation": "Others"},
    123: {"Name": "Hira", "Designation": "Others"},
    2058: {"Name": "Mamata", "Designation": "Others"},
    2052: {"Name": "Panumaya", "Designation": "Others"},
    115645: {"Name": "Ramesh", "Designation": "Others"},
    136987: {"Name": "Ramji", "Designation": "Others"},
    246: {"Name": "Sandesh", "Designation": "Others"},
    170200: {"Name": "Sandesh", "Designation": "Others"},
    2043: {"Name": "Shom", "Designation": "Others"},
    4397: {"Name": "Surya", "Designation": "Others"}
}

designation_order = [
    "Director", "Manager", "Deputy Manager", "Assistant Manager",
    "Administrative Officer", "Engineer", "Senior Operator",
    "Foreman Driver", "Driver", "Office Helper", "Others"
]

colors_map = {
    "Director": "#FF6F61",
    "Deputy Manager": "#6B5B95",
    "Manager": "#88B04B",
    "Assistant Manager": "#F7CAC9",
    "Administrative Officer": "#92A8D1",
    "Engineer": "#955251",
    "Senior Operator": "#B565A7",
    "Foreman Driver": "#009B77",
    "Driver": "#DD4124",
    "Office Helper": "#45B8AC",
    "Others": "#EFC050"
}

# -----------------------------
def bs_to_ad(date_val):
    try:
        if pd.isna(date_val): return None
        d_str = str(date_val).replace("/", "-").replace(".", "-")
        y, m, d = map(int, d_str.split("-"))
        return nepali_datetime.date(y, m, d).to_datetime_date()
    except:
        return None

def find_header_row(file_path):
    try:
        peek = pd.read_excel(file_path, header=None, nrows=30)
        for index, row in peek.iterrows():
            vals = [str(v).strip().lower() for v in row.values]
            if 'id' in vals and 'date' in vals:
                return index
    except:
        pass
    return 0

# -----------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    results = None
    stats = {}
    report_month = ""

    if request.method == "POST":
        report_month = request.form.get("report_month", "")
        files = request.files.getlist("files")
        all_dfs = []
        if not files or files[0].filename == '':
            return render_template("index.html", month=report_month)

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
            combined['dt'] = pd.to_datetime(
                combined['date_ad'].astype(str) + " " + combined['Time'].astype(str),
                errors="coerce"
            )
            combined['type'] = combined['Attendance Check Point'].apply(
                lambda x: 'enter' if 'outside' in str(x).lower() else 'exit'
            )
            combined = combined.dropna(subset=['dt']).sort_values(['ID', 'dt'])

            # -----------------------------
            # Summary calculation
            summary = []
            for emp_id, g in combined.groupby("ID"):
                info = emp_info.get(emp_id)
                if not info:
                   info = {"Name": f"ID-{emp_id}", "Designation": "Others"}
                name = info["Name"]
                designation = info["Designation"]
                working_days = g['date_ad'].nunique()
                total_mins = 0

                for _, day_data in g.groupby(g['dt'].dt.date):
                    day_data = day_data.sort_values("dt")
                    for i in range(len(day_data)-1):
                        curr, nxt = day_data.iloc[i], day_data.iloc[i+1]
                        if curr['type'] == "exit" and nxt['type'] == "enter":
                            diff = (nxt['dt'] - curr['dt']).total_seconds() / 60
                            # --- NEW LOGIC: ignore intervals >= 10 hours (600 mins)
                            if 0 < diff < 600:
                                total_mins += diff

                summary.append({
                    "ID": emp_id,
                    "Name": name,
                    "Designation": designation,
                    "Days": working_days,
                    "Mins": round(total_mins, 2),
                    "Hrs": round(total_mins/60, 2)
                })

            res_df = pd.DataFrame(summary)
            res_df['Designation_order'] = res_df['Designation'].apply(
                lambda x: designation_order.index(x) if x in designation_order else len(designation_order)
            )

            # Sort by designation then hours
            res_plot = res_df.sort_values(['Designation_order', 'Hrs'], ascending=[True, False]).reset_index(drop=True)

            # -----------------------------
            # Stats
            avg_h = res_plot["Hrs"].mean()
            stats = {
                "total": len(res_plot),
                "avg": round(avg_h, 2),
                "max_name": res_plot.loc[res_plot['Hrs'].idxmax()]['Name'],
                "max_val": res_plot['Hrs'].max()
            }

            # -----------------------------
            # Save Summary CSV
            summary_csv_path = os.path.join(OUTPUT_DIR, "summary.csv")
            title_row = pd.DataFrame([["Attendance Report – " + report_month], [""]])
            with open(summary_csv_path, "w", encoding="utf-8", newline="") as f:
                title_row.to_csv(f, index=False, header=False)
                res_plot.drop(columns=["Designation_order"]).to_csv(f, index=False)

            # -----------------------------
            # Generate Daily Report CSV
            daily_csv_path = os.path.join(OUTPUT_DIR, "daily_report.csv")
            combined['date_bs'] = combined['date_ad'].apply(lambda x: nepali_datetime.date.from_datetime_date(x).strftime("%Y-%m-%d"))
            all_dates = sorted(combined['date_bs'].unique())
            daily_data = []

            for emp_id, g in combined.groupby("ID"):
                info = emp_info.get(emp_id)
                if not info:
                   info = {"Name": f"ID-{emp_id}", "Designation": "Others"}
                row = {"ID": emp_id, "Name": info["Name"], "Designation": info["Designation"]}
                total_minutes = 0
                for d in all_dates:
                    row[d] = 0
                for day, day_data in g.groupby(g['date_bs']):
                    day_data = day_data.sort_values("dt")
                    day_mins = 0
                    for i in range(len(day_data)-1):
                        curr, nxt = day_data.iloc[i], day_data.iloc[i+1]
                        if curr['type'] == "exit" and nxt['type'] == "enter":
                            diff = (nxt['dt'] - curr['dt']).total_seconds() / 60
                            # --- NEW LOGIC: ignore intervals >= 10 hours (600 mins)
                            if 0 < diff < 600:
                                day_mins += diff
                    row[day] = round(day_mins, 2)
                    total_minutes += day_mins
                row["Total Minutes"] = round(total_minutes, 2)
                daily_data.append(row)

            daily_df = pd.DataFrame(daily_data)

            # --- SORT DAILY REPORT by DESIGNATION_ORDER and then by ID
            daily_df['Designation_order'] = daily_df['Designation'].apply(
                lambda x: designation_order.index(x) if x in designation_order else len(designation_order)
            )
            daily_df = daily_df.sort_values(['Designation_order', 'ID']).reset_index(drop=True)
            daily_df = daily_df[["ID", "Name", "Designation"] + all_dates + ["Total Minutes"]]
            daily_df.to_csv(daily_csv_path, index=False)

            # -----------------------------
            # Plot graph
            plt.style.use('default')
            dynamic_width = max(12, len(res_plot) * 0.5)
            fig, ax = plt.subplots(figsize=(dynamic_width, 8), dpi=150)
            colors = [colors_map.get(d, "#4f46e5") for d in res_plot['Designation']]
            bars = ax.bar(range(len(res_plot)), res_plot["Hrs"], color=colors, edgecolor='white', alpha=0.9, width=0.6)

            ax.set_xticks(range(len(res_plot)))
            ax.set_xticklabels(res_plot['ID'].astype(str), rotation=90, fontsize=9)

            for i, bar in enumerate(bars):
                h = bar.get_height()
                if h > 0:
                    ax.annotate(f'{round(h,1)}', xy=(bar.get_x() + bar.get_width()/2, h),
                                xytext=(0, 5), textcoords="offset points", ha='center', va='bottom',
                                fontsize=8, fontweight='bold', color='#1e293b')

            ax.set_title(f"Employee Outside Time Breakdown – {report_month}", fontsize=18, fontweight='800', pad=40)
            ax.set_ylabel("Total Hours Outside", fontweight='bold', color='#64748b')
            ax.set_xlabel("Employee ID", fontweight='bold', color='#64748b')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.yaxis.grid(True, linestyle='--', alpha=0.3)

            patches = [mpatches.Patch(color=colors_map[d], label=d) for d in colors_map if d in res_plot['Designation'].unique()]
            ax.legend(handles=patches, loc='upper center', bbox_to_anchor=(0.5, 1.25), ncol=len(patches), frameon=False)

            plt.tight_layout(rect=[0, 0, 1, 0.9])
            plt.savefig(os.path.join(OUTPUT_DIR, "graph.png"), bbox_inches='tight')
            plt.close()

            results = res_plot.drop(columns=["Designation_order"]).to_dict(orient="records")

    return render_template(
        "index.html",
        results=results,
        stats=stats,
        month=report_month,
        colors_map=colors_map
    )


# -----------------------------
@app.route("/download/<f>")
def download(f):
    if f == "csv":
        t = "summary.csv"
    elif f == "graph":
        t = "graph.png"
    elif f == "daily":
        t = "daily_report.csv"
    else:
        return "File not found"
    return send_file(os.path.join(OUTPUT_DIR, t), as_attachment=True)

# -----------------------------
# -----------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    # Just run Flask normally, no automatic browser open
    app.run(host="0.0.0.0", port=port, debug=True)

