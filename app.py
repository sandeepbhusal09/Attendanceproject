

import os
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from flask import Flask, render_template, request, send_file, redirect, url_for
import nepali_datetime
import webbrowser
from threading import Timer
import matplotlib.patches as mpatches
import datetime
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
# ---------------------------------------------------------
# CONFIGURATION & MASTER DATA
# ---------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, template_folder=BASE_DIR)
# --- SECURITY SETUP ---
app.secret_key = 'change_this_secret_key_randomly'  # Required for sessions

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

# User Credentials (Username: admin, Password: password123)
users_db = {
    "admin": generate_password_hash("password123", method='pbkdf2:sha256')
}

class User(UserMixin):
    def __init__(self, id):
        self.id = id

@login_manager.user_loader
def load_user(user_id):
    if user_id in users_db:
        return User(user_id)
    return None

UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Employee Master Data
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

# Explicit list of Shift Employees
SHIFT_IDS = [
    152218, 194948, 133312, 197718, 177988, 116608, 
    192576, 170534, 176805, 128538, 170276, 128345
]

designation_order = [
    "Director", "Manager", "Deputy Manager", "Assistant Manager",
    "Administrative Officer", "Engineer", "Senior Operator",
    "Foreman Driver", "Driver", "Office Helper", "Others"
]

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

def clean_time_val(t):
    try:
        if pd.isna(t): return "00:00:00"
        if hasattr(t, 'strftime'): return t.strftime("%H:%M:%S")
        return str(t).strip()
    except: return "00:00:00"

def find_header_row(file_path):
    try:
        peek = pd.read_excel(file_path, header=None, nrows=30)
        for index, row in peek.iterrows():
            vals = [str(v).strip().lower() for v in row.values]
            if 'id' in vals and 'date' in vals: return index
    except: pass
    return 0

def get_working_window(dt_obj):
    """
    Returns (start_dt, end_dt) based on rules.
    1. Friday: 10 AM - 3 PM
    2. Winter (Kartik 16 - Magh 15): 10 AM - 4 PM
    3. Other: 10 AM - 5 PM
    NOTE: Saturday check is done in the main loop for efficiency.
    """
    bs_date = nepali_datetime.date.from_datetime_date(dt_obj.date())
    
    start_hour, start_min = 10, 0
    end_hour = 17 # Default 5 PM

    # Check Winter: Kartik(7) 16 to Magh(10) 15
    is_winter = False
    m, d = bs_date.month, bs_date.day
    if (m == 7 and d >= 16) or (m == 8) or (m == 9) or (m == 10 and d <= 15):
        is_winter = True
        end_hour = 16 # 4 PM

    # Check Friday (Weekday 4)
    if dt_obj.weekday() == 4:
        end_hour = 15 # 3 PM

    start_time = dt_obj.replace(hour=start_hour, minute=start_min, second=0, microsecond=0)
    end_time = dt_obj.replace(hour=end_hour, minute=0, second=0, microsecond=0)
    
    return start_time, end_time

# ---------------------------------------------------------
# ROUTES
# ---------------------------------------------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('index'))
    error = None
    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")
        if username in users_db and check_password_hash(users_db[username], password):
            login_user(User(username))
            return redirect(url_for('index'))
        else:
            error = "Invalid Username or Password"
    return render_template("login.html", error=error)

@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route("/", methods=["GET", "POST"])
@login_required
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
            
            # Robust ID cleaning to ensure integer matching with SHIFT_IDS
            combined['ID'] = pd.to_numeric(combined['ID'], errors='coerce').fillna(0).astype(int)
            
            combined['date_ad'] = combined['Date'].apply(bs_to_ad)
            combined['time_clean'] = combined['Time'].apply(clean_time_val)
            combined['dt'] = pd.to_datetime(combined['date_ad'].astype(str) + " " + combined['time_clean'], errors="coerce")
            
            combined['type'] = combined['Attendance Check Point'].apply(lambda x: 'enter' if 'outside' in str(x).lower() else 'exit')
            combined = combined.dropna(subset=['dt']).sort_values(['ID', 'dt'])
            combined['date_bs'] = combined['date_ad'].apply(lambda x: nepali_datetime.date.from_datetime_date(x).strftime("%Y-%m-%d"))

            summary = []
            breakdown_list = []
            all_dates = sorted(combined['date_bs'].unique())
            daily_rows = []

            for emp_id, g in combined.groupby("ID"):
                if emp_id == 0: continue
                info = emp_info.get(emp_id, {"Name": f"ID-{emp_id}", "Designation": "Others"})
                
                # Check Shift Status
                is_shift = emp_id in SHIFT_IDS
                
                total_mins = 0
                valid_days_count = 0
                daily_map = {d: 0 for d in all_dates}
                
                for day_bs, day_data in g.groupby(g['date_bs']):
                    day_data = day_data.sort_values("dt")
                    day_mins = 0
                    
                    if len(day_data) > 1:
                        valid_days_count += 1
                    elif len(day_data) == 1 and day_data.iloc[0]['type'] == 'enter':
                        valid_days_count += 1
                    
                    # Calculate Intervals
                    for i in range(len(day_data)-1):
                        curr, nxt = day_data.iloc[i], day_data.iloc[i+1]
                        
                        if curr['type'] == "exit" and nxt['type'] == "enter":
                            exit_time = curr['dt']
                            enter_time = nxt['dt']
                            
                            raw_diff = (enter_time - exit_time).total_seconds() / 60
                            if raw_diff >= 600 or raw_diff < 0: continue

                            # -------------------------------------------------
                            # CALCULATION LOGIC SPLIT
                            # -------------------------------------------------
                            if is_shift:
                                # 1. SHIFT EMPLOYEES
                                # Calculate everything. DO NOT check Saturday.
                                calc_mins = raw_diff
                                if calc_mins > 0:
                                    day_mins += calc_mins
                                    breakdown_list.append({
                                        "ID": emp_id, "Name": info["Name"], "Date": day_bs,
                                        "Exit Time": exit_time.strftime("%H:%M"),
                                        "Enter Time": enter_time.strftime("%H:%M"),
                                        "Calculated Start": exit_time.strftime("%H:%M"),
                                        "Calculated End": enter_time.strftime("%H:%M"),
                                        "Duration": round(calc_mins, 2),
                                        "Type": "Shift"
                                    })
                            
                            else:
                                # 2. NON-SHIFT EMPLOYEES
                                # Rule A: If it is Saturday (Weekday 5), SKIP IT.
                                if exit_time.weekday() == 5:
                                    continue
                                
                                # Rule B: Clamp to Working Hours (10-5, 10-4, 10-3)
                                w_start, w_end = get_working_window(exit_time)
                                
                                eff_start = max(exit_time, w_start)
                                eff_end = min(enter_time, w_end)
                                
                                if eff_end > eff_start:
                                    calc_mins = (eff_end - eff_start).total_seconds() / 60
                                    day_mins += calc_mins
                                    
                                    breakdown_list.append({
                                        "ID": emp_id, "Name": info["Name"], "Date": day_bs,
                                        "Exit Time": exit_time.strftime("%H:%M"),
                                        "Enter Time": enter_time.strftime("%H:%M"),
                                        "Calculated Start": eff_start.strftime("%H:%M"),
                                        "Calculated End": eff_end.strftime("%H:%M"),
                                        "Duration": round(calc_mins, 2),
                                        "Type": "Standard"
                                    })
                            # -------------------------------------------------

                    total_mins += day_mins
                    daily_map[day_bs] = round(day_mins, 2)
                            
                summary.append({
                    "ID": emp_id, 
                    "Name": info["Name"], 
                    "Designation": info["Designation"],
                    "Days": valid_days_count, 
                    "Mins": round(total_mins, 2), 
                    "Hrs": round(total_mins/60, 2)
                })

                d_row = {"ID": emp_id, "Name": info["Name"], "Designation": info["Designation"]}
                d_row.update(daily_map)
                d_row["Total Minutes"] = round(total_mins, 2)
                daily_rows.append(d_row)

            # Outputs
            res_df = pd.DataFrame(summary)
            if not res_df.empty:
                res_df['Designation_order'] = res_df['Designation'].apply(lambda x: designation_order.index(x) if x in designation_order else len(designation_order))
                res_plot = res_df.sort_values(['Designation_order', 'Hrs'], ascending=[True, False]).reset_index(drop=True)
                clean_res = res_plot.drop(columns=["Designation_order"])
                results = clean_res.to_dict(orient="records")
                
                stats = {
                    "total": len(res_plot),
                    "avg": round(res_plot["Hrs"].mean(), 2),
                    "max_name": res_plot.loc[res_plot['Hrs'].idxmax()]['Name']
                }

                with open(os.path.join(OUTPUT_DIR, "summary.csv"), 'w', newline='', encoding='utf-8') as f:
                    f.write(f"Report Month: {report_month}\n")
                    clean_res.to_csv(f, index=False)

                # Graph
                plt.style.use('default')
                dynamic_width = max(24, len(res_plot) * 0.6)
                fig, ax = plt.subplots(figsize=(dynamic_width, 8), dpi=150)
                
                colors = [colors_map.get(d, "#EFC050") for d in res_plot['Designation']]
                bars = ax.bar(range(len(res_plot)), res_plot["Hrs"], color=colors, width=0.6, zorder=3)

                ax.set_xticks(range(len(res_plot)))
                ax.set_xticklabels(res_plot['ID'].astype(str), rotation=90, fontsize=9, color="#1e293b")

                for i, bar in enumerate(bars):
                    height = bar.get_height()
                    ax.text(bar.get_x() + bar.get_width()/2., height + 1,
                            f'{height}', ha='center', va='bottom', 
                            fontsize=8, fontweight='800', color='#0f172a')

                ax.set_title(f"Employee Outside Time Breakdown – {report_month}", fontsize=20, fontweight='800', pad=60)
                ax.set_ylabel("Total Hours Outside", fontsize=10, fontweight='bold', color='#475569')
                ax.set_xlabel("Employee ID", fontsize=10, fontweight='bold', color='#475569')

                ax.yaxis.grid(True, linestyle='--', color="#e2e8f0", alpha=0.7, zorder=0)
                ax.set_axisbelow(True)

                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                ax.spines['left'].set_color('#cbd5e1')
                ax.spines['bottom'].set_color('#cbd5e1')

                legend_patches = [mpatches.Patch(color=colors_map[d], label=d) for d in designation_order if d in res_plot['Designation'].unique()]
                ax.legend(handles=legend_patches, loc='upper center', bbox_to_anchor=(0.5, 1.15), 
                        ncol=len(legend_patches), frameon=False, fontsize=10)

                plt.tight_layout()
                plt.savefig(os.path.join(OUTPUT_DIR, "graph.png"), bbox_inches='tight')
                plt.close()
            
            else:
                results = []
                stats = {"total": 0, "avg": 0, "max_name": "-"}

            # Daily Report
            daily_df = pd.DataFrame(daily_rows)
            if not daily_df.empty:
                cols = ["ID", "Name", "Designation"] + all_dates + ["Total Minutes"]
                daily_df = daily_df[cols]
                daily_df['Designation_order'] = daily_df['Designation'].apply(lambda x: designation_order.index(x) if x in designation_order else len(designation_order))
                daily_df = daily_df.sort_values(['Designation_order', 'ID']).drop(columns=["Designation_order"])
            
            with open(os.path.join(OUTPUT_DIR, "daily_report.csv"), 'w', newline='', encoding='utf-8') as f:
                f.write(f"Report Month: {report_month}\n")
                daily_df.to_csv(f, index=False)

            # Breakdown
            bd_df = pd.DataFrame(breakdown_list)
            if not bd_df.empty:
                bd_df = bd_df.sort_values(['ID', 'Date', 'Exit Time'])
            bd_df.to_csv(os.path.join(OUTPUT_DIR, "detailed_breakdown.csv"), index=False)

    return render_template("index.html", results=results, stats=stats, month=report_month, colors_map=colors_map)

@app.route("/employee/<int:emp_id>")
@login_required
def employee_details(emp_id):
    month = request.args.get('month', '')
    report_path = os.path.join(OUTPUT_DIR, "daily_report.csv")
    breakdown_path = os.path.join(OUTPUT_DIR, "detailed_breakdown.csv")
    
    if not os.path.exists(report_path): return redirect(url_for('index'))
    
    df = pd.read_csv(report_path, skiprows=1)
    emp_subset = df[df['ID'] == emp_id]
    if emp_subset.empty: return "Not Found", 404
    emp_row = emp_subset.iloc[0]
    
    details_map = {}
    if os.path.exists(breakdown_path):
        try:
            bd_df = pd.read_csv(breakdown_path)
            emp_bd = bd_df[bd_df['ID'] == emp_id]
            for _, row in emp_bd.iterrows():
                d = row['Date']
                if d not in details_map: details_map[d] = []
                type_lbl = row.get('Type', 'Standard')
                t_str = f"{row['Calculated Start']} - {row['Calculated End']}"
                note = f"({type_lbl}) Raw: {row['Exit Time']}-{row['Enter Time']}"
                details_map[d].append({"time": t_str, "note": note, "dur": row['Duration']})
        except: pass

    cols = df.columns.tolist()
    daily_list = []
    try:
        start_idx = cols.index("Designation") + 1
        end_idx = cols.index("Total Minutes")
        for d in cols[start_idx:end_idx]:
            val = emp_row[d]
            if val > 0:
                daily_list.append({"date": d, "mins": val, "hrs": round(val/60, 2), "details": details_map.get(d, [])})
    except: pass

    summary = {
        "ID": emp_id, "Name": emp_row['Name'], 
        "Designation": emp_row['Designation'], 
        "TotalMins": emp_row['Total Minutes'], 
        "TotalHrs": round(emp_row['Total Minutes']/60, 2)
    }
    return render_template("index.html", employee_view=True, daily_list=daily_list, summary=summary, month=month)

@app.route("/download/<f>")
@login_required 
def download(f):
    m = {"csv": "summary.csv", "graph": "graph.png", "daily": "daily_report.csv", "breakdown": "detailed_breakdown.csv"}
    if f in m and os.path.exists(os.path.join(OUTPUT_DIR, m[f])):
        return send_file(os.path.join(OUTPUT_DIR, m[f]), as_attachment=True)
    return "File not found", 404
        

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    # Change URL below to /login
    Timer(1, lambda: webbrowser.open(f"http://127.0.0.1:{port}/login")).start() 
    app.run(host="127.0.0.1", port=port, debug=True)

        

if __name__ == "__main__":
    # Get the PORT from Render environment, default to 5000
    port = int(os.environ.get("PORT", 5000))
    
    # IMPORTANT: host must be "0.0.0.0" not "127.0.0.1"
    app.run(host="0.0.0.0", port=port, debug=False)
