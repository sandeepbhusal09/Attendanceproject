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

# ---------------------------------------------------------
# CONFIGURATION & MASTER DATA
# ---------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, template_folder=BASE_DIR)

UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

emp_info = {
    160155: {"Name": "tara", "Designation": "Administrative Officer"},
    3812: {"Name": "MADHAV", "Designation": "Administrative Officer"},
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

# Color map updated to match the image precisely
colors_map = {
    "Director": "#FF7C70",
    "Manager": "#94B454",
    "Deputy Manager": "#6E639E",
    "Assistant Manager": "#F8D1D1",
    "Administrative Officer": "#94B0D8",
    "Engineer": "#9E6363",
    "Senior Operator": "#BC69AA",
    "Foreman Driver": "#00A37A",
    "Driver": "#E1523D",
    "Office Helper": "#63C1B5",
    "Others": "#F2C764"
}

# ---------------------------------------------------------
# HELPERS
# ---------------------------------------------------------
def bs_to_ad(date_val):
    try:
        if pd.isna(date_val): return None
        d_str = str(date_val).replace("/", "-").replace(".", "-")
        y, m, d = map(int, d_str.split("-"))
        return nepali_datetime.date(y, m, d).to_datetime_date()
    except: return None

def find_header_row(file_path):
    try:
        peek = pd.read_excel(file_path, header=None, nrows=30)
        for index, row in peek.iterrows():
            vals = [str(v).strip().lower() for v in row.values]
            if 'id' in vals and 'date' in vals: return index
    except: pass
    return 0

# ---------------------------------------------------------
# ROUTES
# ---------------------------------------------------------

@app.route("/", methods=["GET", "POST"])
def index():
    results, stats, report_month = None, {}, ""
    if request.method == "POST":
        report_month = request.form.get("report_month", "")
        files = request.files.getlist("files")
        all_dfs = []
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
                info = emp_info.get(emp_id, {"Name": f"ID-{emp_id}", "Designation": "Others"})
                total_mins = 0
                for _, day_data in g.groupby(g['dt'].dt.date):
                    day_data = day_data.sort_values("dt")
                    for i in range(len(day_data)-1):
                        curr, nxt = day_data.iloc[i], day_data.iloc[i+1]
                        if curr['type'] == "exit" and nxt['type'] == "enter":
                            diff = (nxt['dt'] - curr['dt']).total_seconds() / 60
                            if 0 < diff < 600: total_mins += diff
                summary.append({
                    "ID": emp_id, "Name": info["Name"], "Designation": info["Designation"],
                    "Days": g['date_ad'].nunique(), "Mins": round(total_mins, 2), "Hrs": round(total_mins/60, 2)
                })

            res_df = pd.DataFrame(summary)
            res_df['Designation_order'] = res_df['Designation'].apply(lambda x: designation_order.index(x) if x in designation_order else len(designation_order))
            # Sort exactly like the chart: Designation Hierarchy, then Highest Hours
            res_plot = res_df.sort_values(['Designation_order', 'Hrs'], ascending=[True, False]).reset_index(drop=True)
            results = res_plot.drop(columns=["Designation_order"]).to_dict(orient="records")
            
            stats = {
                "total": len(res_plot),
                "avg": round(res_plot["Hrs"].mean(), 2),
                "max_name": res_plot.loc[res_plot['Hrs'].idxmax()]['Name']
            }
            res_plot.to_csv(os.path.join(OUTPUT_DIR, "summary.csv"), index=False)

            # Daily CSV generation
            combined['date_bs'] = combined['date_ad'].apply(lambda x: nepali_datetime.date.from_datetime_date(x).strftime("%Y-%m-%d"))
            all_dates = sorted(combined['date_bs'].unique())
            daily_rows = []
            for emp_id, g in combined.groupby("ID"):
                info = emp_info.get(emp_id, {"Name": f"ID-{emp_id}", "Designation": "Others"})
                row = {"ID": emp_id, "Name": info["Name"], "Designation": info["Designation"]}
                total_m = 0
                for d in all_dates: row[d] = 0
                for day, day_data in g.groupby(g['date_bs']):
                    dmins = 0
                    day_data = day_data.sort_values("dt")
                    for i in range(len(day_data)-1):
                        curr, nxt = day_data.iloc[i], day_data.iloc[i+1]
                        if curr['type'] == "exit" and nxt['type'] == "enter":
                            diff = (nxt['dt'] - curr['dt']).total_seconds() / 60
                            if 0 < diff < 600: dmins += diff
                    row[day] = round(dmins, 2)
                    total_m += dmins
                row["Total Minutes"] = round(total_m, 2)
                daily_rows.append(row)
            daily_df = pd.DataFrame(daily_rows)
            daily_df['Designation_order'] = daily_df['Designation'].apply(lambda x: designation_order.index(x) if x in designation_order else len(designation_order))
            daily_df = daily_df.sort_values(['Designation_order', 'ID']).drop(columns=["Designation_order"])
            daily_df.to_csv(os.path.join(OUTPUT_DIR, "daily_report.csv"), index=False)

            # -----------------------------------------------------
            # GRAPH GENERATION (Exactly matching the image)
            # -----------------------------------------------------
            plt.style.use('default')
            # Extra wide figure to accommodate many bars
            dynamic_width = max(24, len(res_plot) * 0.6)
            fig, ax = plt.subplots(figsize=(dynamic_width, 8), dpi=150)
            
            colors = [colors_map.get(d, "#EFC050") for d in res_plot['Designation']]
            bars = ax.bar(range(len(res_plot)), res_plot["Hrs"], color=colors, width=0.6, zorder=3)

            # X-axis labels (vertical like image)
            ax.set_xticks(range(len(res_plot)))
            ax.set_xticklabels(res_plot['ID'].astype(str), rotation=90, fontsize=9, color="#1e293b")

            # Data labels on top of bars
            for i, bar in enumerate(bars):
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2., height + 1,
                        f'{height}', ha='center', va='bottom', 
                        fontsize=8, fontweight='800', color='#0f172a')

            # Labels and styling
            ax.set_title(f"Employee Outside Time Breakdown – {report_month}", fontsize=20, fontweight='800', pad=60)
            ax.set_ylabel("Total Hours Outside", fontsize=10, fontweight='bold', color='#475569')
            ax.set_xlabel("Employee ID", fontsize=10, fontweight='bold', color='#475569')

            # Gridlines (dashed horizontal only)
            ax.yaxis.grid(True, linestyle='--', color="#e2e8f0", alpha=0.7, zorder=0)
            ax.set_axisbelow(True)

            # Remove top and right spines
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('#cbd5e1')
            ax.spines['bottom'].set_color('#cbd5e1')

            # Legend on TOP
            legend_patches = [mpatches.Patch(color=colors_map[d], label=d) for d in designation_order if d in res_plot['Designation'].unique()]
            ax.legend(handles=legend_patches, loc='upper center', bbox_to_anchor=(0.5, 1.15), 
                      ncol=len(legend_patches), frameon=False, fontsize=10)

            plt.tight_layout()
            plt.savefig(os.path.join(OUTPUT_DIR, "graph.png"), bbox_inches='tight')
            plt.close()

    return render_template("index.html", results=results, stats=stats, month=report_month, colors_map=colors_map)

@app.route("/employee/<int:emp_id>")
def employee_details(emp_id):
    month = request.args.get('month', '')
    report_path = os.path.join(OUTPUT_DIR, "daily_report.csv")
    if not os.path.exists(report_path): return redirect(url_for('index'))
    df = pd.read_csv(report_path)
    emp_subset = df[df['ID'] == emp_id]
    if emp_subset.empty: return "Not Found", 404
    emp_row = emp_subset.iloc[0]
    cols = df.columns.tolist()
    date_cols = cols[cols.index("Designation")+1 : cols.index("Total Minutes")]
    daily_list = [{"date": d, "mins": emp_row[d], "hrs": round(emp_row[d]/60, 2)} for d in date_cols if emp_row[d] > 0]
    summary = {"ID": emp_id, "Name": emp_row['Name'], "Designation": emp_row['Designation'], "TotalMins": emp_row['Total Minutes'], "TotalHrs": round(emp_row['Total Minutes']/60, 2)}
    return render_template("index.html", employee_view=True, daily_list=daily_list, summary=summary, month=month)

@app.route("/download/<f>")
def download(f):
    m = {"csv": "summary.csv", "graph": "graph.png", "daily": "daily_report.csv"}
    return send_file(os.path.join(OUTPUT_DIR, m[f]), as_attachment=True)
    

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    # Just run Flask normally, no automatic browser open
    app.run(host="0.0.0.0", port=port, debug=True)

