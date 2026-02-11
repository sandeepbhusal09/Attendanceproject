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
app.secret_key = 'change_this_secret_key_randomly'

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

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
# HARDCODED ROSTER DATA
# ---------------------------------------------------------
STATIC_ROSTER_RAW = {
    152218: "G,G,G,N,R,R,D,D,N,R,R,D,D,N,R,D,D,D,N,R,R,D,D,N,R,R,D,D,N,R", # Saugat
    194948: "G,G,G,R,N,D,D,D,N,R,G,D,D,R,N,G,R,M,N,R,R,R,D,R,N,G,D,D,N,D", # Praveen
    133312: "M,M,N,R,G,M,M,N,R,M,M,M,N,R,R,R,M,N,R,R,M,M,N,R,R,M,FK,FK,FK,FK", # Sarita
    197718: "M,M,N,G,G,D,M,N,G,G,M,MN,R,R,R,M,R,R,R,R,M,MN,N,R,R,M,MN,N,R,R", # Saru
    177988: "D,N,G,G,M,FK,FK,FK,FK,FK,D,R,N,G,M,MN,R,R,G,M,D,D,R,G,MN,D,N,R,D,M", # Sujeet
    116608: "D,N,R,G,M,MN,N,G,G,M,D,N,D,R,M,D,N,N,D,M,D,N,R,M,M,D,R,N,R,MN", # Suman B
    192576: "D,N,R,G,M,D,N,G,G,M,D,N,D,R,M,D,N,R,M,M,D,N,R,R,M,D,N,R,M,M", # Anil
    170534: "N,R,G,D,D,N,R,G,D,D,N,R,G,D,D,R,N,R,D,D,N,R,R,D,D,N,R,R,D,M", # Dhanjit
    176805: "N,G,D,D,D,N,N,G,D,D,N,R,G,D,D,N,R,R,D,D,N,R,R,D,D,N,R,M,D,D", # Hiralal
    128538: "N,G,D,D,D,N,R,G,D,D,N,R,G,D,D,N,R,R,D,D,R,R,R,D,D,N,FK,FK,FK,FK", # Bimal
    170276: "R,D,M,MN,R,R,G,M,M,N,R,G,M,MN,R,R,M,D,M,N,R,G,M,MN,R,R,M,M,M,N", # Purna
    128345: "R,D,M,M,N,R,G,M,M,N,R,G,M,M,N,R,D,M,G,N,N,R,M,G,G,G,G,G,G,G", # Suman P
}

def get_static_roster():
    """Converts the comma strings into a lookup dict {(id, day_num): code}"""
    roster_map = {}
    for eid, s_str in STATIC_ROSTER_RAW.items():
        codes = [x.strip() for x in s_str.split(',')]
        for i, code in enumerate(codes, start=1):
            roster_map[(eid, i)] = code
    return roster_map

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

def get_general_window(dt_obj):
    """Returns (start_dt, end_dt) for Standard/General Logic."""
    bs_date = nepali_datetime.date.from_datetime_date(dt_obj.date())
    start_hour, start_min = 10, 0
    end_hour = 17 # Default 5 PM

    # Check Winter: Kartik(7) 16 to Magh(10) 15
    m, d = bs_date.month, bs_date.day
    if (m == 7 and d >= 16) or (m == 8) or (m == 9) or (m == 10 and d <= 15):
        end_hour = 16 # 4 PM

    # Check Friday (Weekday 4)
    if dt_obj.weekday() == 4:
        end_hour = 15 # 3 PM

    start_time = dt_obj.replace(hour=start_hour, minute=start_min, second=0, microsecond=0)
    end_time = dt_obj.replace(hour=end_hour, minute=0, second=0, microsecond=0)
    return start_time, end_time

def get_shift_windows(shift_code, date_obj):
    """
    Returns a list of (start, end) tuples for the allowed working hours based on shift code.
    Strict Logic: Any time OUTSIDE these windows is IGNORED.
    """
    windows = []
    base_date = date_obj.replace(minute=0, second=0, microsecond=0)
    code = str(shift_code).strip().upper()

    if code == 'M':
        # 08:00 to 14:00
        windows.append((base_date.replace(hour=8), base_date.replace(hour=14)))
    elif code == 'D':
        # 14:00 to 20:00
        windows.append((base_date.replace(hour=14), base_date.replace(hour=20)))
    elif code == 'N':
        # 20:00 to 08:00 Next Day
        start = base_date.replace(hour=20)
        end = (base_date + datetime.timedelta(days=1)).replace(hour=8)
        windows.append((start, end))
    elif code == 'MN':
        # Morning Part (08:00-14:00)
        windows.append((base_date.replace(hour=8), base_date.replace(hour=14)))
        # Night Part (20:00-08:00 Next Day) - Gap is IGNORED
        start_n = base_date.replace(hour=20)
        end_n = (base_date + datetime.timedelta(days=1)).replace(hour=8)
        windows.append((start_n, end_n))
    elif code == 'G' or code == 'FK': 
        # General duty logic
        s, e = get_general_window(date_obj)
        windows.append((s, e))
        
    return windows

def detect_shift_type(punch_time):
    """Heuristic to guess actual shift based on punch time for violation check."""
    h = punch_time.hour
    if 5 <= h < 11: return 'M'  # Starts morning
    if 11 <= h < 17: return 'D' # Starts day
    if 17 <= h <= 23: return 'N' # Starts night
    return 'U'

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
    roster_data = get_static_roster()

    if request.method == "POST":
        report_month = request.form.get("report_month", "")
        att_files = request.files.getlist("files")

        all_dfs = []
        for f in att_files:
            if f.filename.endswith((".xlsx", ".xls")):
                path = os.path.join(UPLOAD_DIR, f.filename)
                f.save(path)
                skip = find_header_row(path)
                df = pd.read_excel(path, skiprows=skip)
                df.columns = [str(c).strip() for c in df.columns]
                all_dfs.append(df)

        if all_dfs:
            combined = pd.concat(all_dfs, ignore_index=True)
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
                is_shift_employee = emp_id in SHIFT_IDS
                
                total_mins = 0
                valid_days_count = 0
                daily_map = {d: 0 for d in all_dates}
                notes_list = []
                
                # Map to store Actual Shift Detected per day
                day_actual_shift_map = {}

                # --- 1. Calculate Valid Days and Check Violations ---
                unique_days_in_log = g['date_bs'].unique()
                
                for day_str in unique_days_in_log:
                    day_logs = g[g['date_bs'] == day_str]
                    
                    # IGNORE LOGIC: Single Exit (likely Night shift end)
                    if len(day_logs) == 1 and day_logs.iloc[0]['type'] == 'exit':
                        continue

                    # Detect Actual Shift and store it
                    if not day_logs.empty:
                        first_time = day_logs['dt'].min()
                        detected = detect_shift_type(first_time)
                        day_actual_shift_map[day_str] = detected

                    try:
                        bs_obj = nepali_datetime.datetime.strptime(day_str, "%Y-%m-%d")
                        bs_day_num = bs_obj.day
                    except:
                        bs_day_num = 1
                    
                    if is_shift_employee:
                        shift_code = roster_data.get((emp_id, bs_day_num), 'G')
                    else:
                        shift_code = 'G'
                    
                    # MN shift counts as 2 days
                    if str(shift_code).upper() == 'MN':
                        valid_days_count += 2 
                    else:
                        valid_days_count += 1 

                    # Violation Checks (Assignment vs Actual)
                    if is_shift_employee:
                        s_upper = str(shift_code).upper()
                        actual_shift = day_actual_shift_map.get(day_str, 'U')
                        
                        if s_upper == 'R':
                            notes_list.append(f"Present on Rest Day: {day_str}")
                        
                        if s_upper in ['M', 'D', 'N']:
                            if s_upper == 'M' and actual_shift in ['D', 'N']:
                                notes_list.append(f"Shift Violation {day_str}: Assigned M, Worked {actual_shift}")
                            elif s_upper == 'D' and actual_shift in ['M', 'N']:
                                notes_list.append(f"Shift Violation {day_str}: Assigned D, Worked {actual_shift}")
                            elif s_upper == 'N' and actual_shift in ['M', 'D']:
                                notes_list.append(f"Shift Violation {day_str}: Assigned N, Worked {actual_shift}")

                # --- 2. Calculate Outside Time Intervals ---
                g = g.sort_values('dt')
                
                for i in range(len(g)-1):
                    curr = g.iloc[i]
                    nxt = g.iloc[i+1]
                    
                    if curr['type'] == "exit" and nxt['type'] == "enter":
                        exit_time = curr['dt']
                        enter_time = nxt['dt']
                        
                        raw_diff = (enter_time - exit_time).total_seconds() / 60
                        if raw_diff >= 840 or raw_diff < 0: continue 

                        bs_date_obj = nepali_datetime.date.from_datetime_date(exit_time.date())
                        day_bs_str = bs_date_obj.strftime("%Y-%m-%d")
                        bs_day_num = bs_date_obj.day
                        
                        calc_mins = 0
                        note_type = "Standard"
                        windows = []

                        if is_shift_employee:
                            # 1. Get Assigned Code
                            s_code = roster_data.get((emp_id, bs_day_num), 'G')
                            s_code_str = str(s_code).upper()
                            
                            # 2. Determine Working Shift (Shift Done)
                            # Default to Assigned
                            working_shift = s_code_str
                            
                            # Override with Actual if Assigned was M, D, N, R or MN
                            detected_actual = day_actual_shift_map.get(day_bs_str, 'U')
                            
                            if s_code_str in ['M', 'D', 'N', 'MN', 'R'] and detected_actual in ['M', 'D', 'N']:
                                # Special check for MN: MN starts Morning (M). So if detected is M, it matches.
                                if s_code_str == 'MN' and detected_actual == 'M':
                                    working_shift = 'MN'
                                else:
                                    working_shift = detected_actual

                            if working_shift == 'R':
                                calc_mins = 0 
                            else:
                                windows = get_shift_windows(working_shift, exit_time)
                                note_type = f"Shift {working_shift}"
                        else:
                            if exit_time.weekday() == 5: # Saturday
                                continue
                            w_start, w_end = get_general_window(exit_time)
                            windows = [(w_start, w_end)]
                            note_type = "Standard"

                        # Calculate Overlap and Capture Effective Segments
                        effective_segments = []
                        total_segment_mins = 0
                        
                        for w_start, w_end in windows:
                            eff_s = max(exit_time, w_start)
                            eff_e = min(enter_time, w_end)
                            if eff_e > eff_s:
                                mins = (eff_e - eff_s).total_seconds() / 60
                                total_segment_mins += mins
                                effective_segments.append((eff_s, eff_e))

                        if total_segment_mins > 0:
                            total_mins += total_segment_mins
                            if day_bs_str in daily_map:
                                daily_map[day_bs_str] += total_segment_mins
                            
                            c_start = effective_segments[0][0].strftime("%H:%M")
                            c_end = effective_segments[-1][1].strftime("%H:%M")
                            
                            breakdown_list.append({
                                "ID": emp_id, "Name": info["Name"], "Date": day_bs_str,
                                "Exit Time": exit_time.strftime("%H:%M"),
                                "Enter Time": enter_time.strftime("%H:%M"),
                                "Calculated Start": c_start, 
                                "Calculated End": c_end,     
                                "Duration": round(total_segment_mins, 2),
                                "Type": note_type
                            })

                final_notes = "; ".join(notes_list)

                summary.append({
                    "ID": emp_id, 
                    "Name": info["Name"], 
                    "Designation": info["Designation"],
                    "Days": valid_days_count, 
                    "Mins": round(total_mins, 2), 
                    "Hrs": round(total_mins/60, 2),
                    "Notes": final_notes
                })

                d_row = {"ID": emp_id, "Name": info["Name"], "Designation": info["Designation"]}
                d_row.update(daily_map)
                d_row["Total Minutes"] = round(total_mins, 2)
                d_row["Notes"] = final_notes
                daily_rows.append(d_row)

            # Outputs Generation
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

                plt.style.use('default')
                dynamic_width = max(24, len(res_plot) * 0.6)
                fig, ax = plt.subplots(figsize=(dynamic_width, 8), dpi=150)
                colors = [colors_map.get(d, "#EFC050") for d in res_plot['Designation']]
                bars = ax.bar(range(len(res_plot)), res_plot["Hrs"], color=colors, width=0.6, zorder=3)
                ax.set_xticks(range(len(res_plot)))
                ax.set_xticklabels(res_plot['ID'].astype(str), rotation=90, fontsize=9, color="#1e293b")
                for i, bar in enumerate(bars):
                    height = bar.get_height()
                    ax.text(bar.get_x() + bar.get_width()/2., height + 0.1, f'{height}', ha='center', va='bottom', fontsize=8, fontweight='800')
                ax.set_title(f"Employee Outside Time Breakdown – {report_month}", fontsize=20, fontweight='800', pad=60)
                ax.set_ylabel("Total Hours Outside", fontsize=10, fontweight='bold')
                ax.legend(handles=[mpatches.Patch(color=colors_map[d], label=d) for d in designation_order if d in res_plot['Designation'].unique()], loc='upper center', bbox_to_anchor=(0.5, 1.15), ncol=10, frameon=False)
                plt.tight_layout()
                plt.savefig(os.path.join(OUTPUT_DIR, "graph.png"), bbox_inches='tight')
                plt.close()
            else:
                results = []
                stats = {"total": 0, "avg": 0, "max_name": "-"}

            # Daily Report & Breakdown
            daily_df = pd.DataFrame(daily_rows)
            if not daily_df.empty:
                cols = ["ID", "Name", "Designation"] + all_dates + ["Total Minutes", "Notes"]
                for c in cols:
                    if c not in daily_df.columns: daily_df[c] = 0
                daily_df = daily_df[cols]
            
            with open(os.path.join(OUTPUT_DIR, "daily_report.csv"), 'w', newline='', encoding='utf-8') as f:
                f.write(f"Report Month: {report_month}\n")
                daily_df.to_csv(f, index=False)

            bd_df = pd.DataFrame(breakdown_list)
            if not bd_df.empty:
                bd_df = bd_df.sort_values(['ID', 'Date', 'Exit Time'])
            bd_df.to_csv(os.path.join(OUTPUT_DIR, "detailed_breakdown.csv"), index=False)

    return render_template("index.html", results=results, stats=stats, month=report_month, colors_map=colors_map)

# ---------------------------------------------------------
# REMAINING ROUTES
# ---------------------------------------------------------
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
        exclude = ["ID", "Name", "Designation", "Total Minutes", "Notes"]
        date_cols = [c for c in cols if c not in exclude]
        for d in date_cols:
            val = emp_row[d]
            if val > 0 or d in details_map:
                daily_list.append({"date": d, "mins": val, "hrs": round(val/60, 2), "details": details_map.get(d, [])})
    except: pass
    
    extra_notes = emp_row.get('Notes', '')
    if pd.isna(extra_notes): extra_notes = ""

    summary = {
        "ID": emp_id, "Name": emp_row['Name'], 
        "Designation": emp_row['Designation'], 
        "TotalMins": emp_row['Total Minutes'], 
        "TotalHrs": round(emp_row['Total Minutes']/60, 2),
        "Notes": extra_notes
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
    Timer(1, lambda: webbrowser.open(f"http://127.0.0.1:{port}/login")).start() 
    app.run(host="127.0.0.1", port=port, debug=True)
