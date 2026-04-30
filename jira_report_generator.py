"""
Jira Weekly Report Generator
Desktop app to generate Excel reports from Jira issues
VSCode Dark + Pixel Art Theme
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import requests
import requests.auth
import json
import datetime
from datetime import timedelta
import calendar
import os
import re
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


# VSCode Dark Theme Colors
VSCODE_BG = "#1e1e1e"
VSCODE_SURFACE = "#252526"
VSCODE_SURFACE_ALT = "#2d2d2d"
VSCODE_BORDER = "#3c3c3c"
VSCODE_BLUE = "#007acc"
VSCODE_CYAN = "#4ec9b0"
VSCODE_ORANGE = "#ce9178"
VSCODE_GREEN = "#6a9955"
VSCODE_RED = "#f14c4c"
VSCODE_YELLOW = "#dcdcaa"
VSCODE_TEXT = "#d4d4d4"
VSCODE_TEXT_DIM = "#808080"
VSCODE_SELECT = "#264f78"


class JiraReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Jira Report")
        self.root.geometry("720x580")
        self.root.minsize(620, 480)
        self.root.configure(bg=VSCODE_BG)

        # Jira API settings
        self.base_url = "https://ticket.quectel.com"
        self.session = requests.Session()
        self.logged_in = False
        self.username = None
        self.user_email = None

        # Config file
        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".jira_config")
        self.last_save_dir = os.path.expanduser("~")

        self.load_credentials()
        self.setup_ui()

    def load_credentials(self):
        self.saved_username = ""
        self.saved_password = ""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r") as f:
                    data = json.load(f)
                    self.saved_username = data.get("username", "")
                    self.saved_password = data.get("password", "")
                    self.last_save_dir = data.get("last_save_dir", os.path.expanduser("~"))
            except:
                pass

    def save_credentials(self, username, password):
        try:
            with open(self.config_file, "w") as f:
                json.dump({
                    "username": username,
                    "password": password,
                    "last_save_dir": self.last_save_dir
                }, f)
        except:
            pass

    def style_widgets(self):
        """Apply VSCode Dark + Pixel style to ttk widgets"""
        style = ttk.Style()
        style.theme_use('clam')

        # Frame
        style.configure("TFrame", background=VSCODE_SURFACE)

        # Labelframe
        style.configure("TLabelframe", background=VSCODE_SURFACE, foreground=VSCODE_CYAN,
                       bordercolor=VSCODE_BORDER, relief="solid")
        style.configure("TLabelframe.Label", background=VSCODE_SURFACE, foreground=VSCODE_CYAN,
                       font=("Courier New", 10, "bold"))

        # Button - pixel style
        style.configure("Pixel.TButton", background=VSCODE_SURFACE_ALT, foreground=VSCODE_TEXT,
                       borderwidth=2, bordercolor=VSCODE_BORDER, relief="solid",
                       font=("Courier New", 9))
        style.map("Pixel.TButton",
                 background=[("active", VSCODE_BLUE), ("pressed", VSCODE_SELECT)],
                 foreground=[("active", VSCODE_TEXT)])

        # Entry
        style.configure("Pixel.TEntry", fieldbackground=VSCODE_SURFACE_ALT,
                       foreground=VSCODE_TEXT, bordercolor=VSCODE_BORDER,
                       borderwidth=2, relief="solid")

        # Combobox
        style.configure("Pixel.TCombobox", fieldbackground=VSCODE_SURFACE_ALT,
                       foreground=VSCODE_TEXT, background=VSCODE_SURFACE_ALT,
                       bordercolor=VSCODE_BORDER, borderwidth=2, relief="solid")
        style.map("Pixel.TCombobox",
                 fieldbackground=[("readonly", VSCODE_SURFACE_ALT)],
                 selectbackground=[("readonly", VSCODE_SELECT)],
                 selectforeground=[("readonly", VSCODE_TEXT)])

        # Checkbutton
        style.configure("Pixel.TCheckbutton", background=VSCODE_SURFACE,
                       foreground=VSCODE_TEXT, font=("Courier New", 9))
        style.map("Pixel.TCheckbutton",
                 background=[("active", VSCODE_SURFACE)],
                 indicatorcolor=[("selected", VSCODE_BLUE), ("!selected", VSCODE_SURFACE_ALT)])

        # Scrollbar
        style.configure("Vertical.TScrollbar", background=VSCODE_SURFACE_ALT,
                       troughcolor=VSCODE_SURFACE, bordercolor=VSCODE_BORDER,
                       arrowcolor=VSCODE_TEXT)

    def setup_ui(self):
        self.style_widgets()

        # Title with pixel art style
        title_frame = tk.Frame(self.root, bg=VSCODE_BG, pady=10)
        title_frame.pack(fill=tk.X)
        title_label = tk.Label(
            title_frame,
            text="◆ JIRA WEEKLY REPORT ◆",
            font=("Courier New", 18, "bold"),
            fg=VSCODE_CYAN,
            bg=VSCODE_BG
        )
        title_label.pack()

        # Version tag
        version_label = tk.Label(
            title_frame,
            text="[ v1.0 ]",
            font=("Courier New", 8),
            fg=VSCODE_TEXT_DIM,
            bg=VSCODE_BG
        )
        version_label.pack()

        # Main container
        main_frame = tk.Frame(self.root, bg=VSCODE_SURFACE, padx=15, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # === Login Section ===
        login_frame = tk.Frame(main_frame, bg=VSCODE_SURFACE_ALT, padx=10, pady=10,
                               relief="solid", borderwidth=2, highlightbackground=VSCODE_BORDER, highlightthickness=2)
        login_frame.pack(fill=tk.X, pady=(0, 8))

        # Login title
        login_title = tk.Label(login_frame, text="▼ LOGIN",
                              font=("Courier New", 10, "bold"),
                              fg=VSCODE_ORANGE, bg=VSCODE_SURFACE_ALT)
        login_title.grid(row=0, column=0, columnspan=8, sticky=tk.W, pady=(0, 5))

        ttk.Label(login_frame, text="User:", style="Pixel.TLabel" if False else "").grid(
            row=1, column=0, sticky=tk.W, padx=(0, 5))
        login_frame.columnconfigure(1, weight=0)

        self.username_var = tk.StringVar(value="")
        username_entry = ttk.Entry(login_frame, textvariable=self.username_var, width=18,
                                  style="Pixel.TEntry")
        username_entry.grid(row=1, column=1, padx=(0, 8))

        ttk.Label(login_frame, text="Pwd:").grid(row=1, column=2, sticky=tk.W, padx=(0, 5))

        self.password_var = tk.StringVar(value="")
        self.password_entry = ttk.Entry(login_frame, textvariable=self.password_var, show="*",
                                       width=18, style="Pixel.TEntry")
        self.password_entry.grid(row=1, column=3, padx=(0, 5))

        # Show/hide password toggle
        self.show_password_var = tk.BooleanVar(value=False)
        show_pwd_btn = tk.Checkbutton(login_frame, text="◉", variable=self.show_password_var,
                                     command=self.toggle_password_visibility,
                                     bg=VSCODE_SURFACE_ALT, fg=VSCODE_TEXT,
                                     selectcolor=VSCODE_BLUE, font=("Courier New", 10))
        show_pwd_btn.grid(row=1, column=4, padx=2)

        self.remember_var = tk.BooleanVar(value=False)
        remember_btn = ttk.Checkbutton(login_frame, text="Remember", variable=self.remember_var,
                                      style="Pixel.TCheckbutton")
        remember_btn.grid(row=1, column=5, padx=5)

        self.login_btn = ttk.Button(login_frame, text="Login", command=self.login,
                                    width=7, style="Pixel.TButton")
        self.login_btn.grid(row=1, column=6, padx=(3, 3))

        self.logout_btn = ttk.Button(login_frame, text="Logout", command=self.logout,
                                     state=tk.DISABLED, width=7, style="Pixel.TButton")
        self.logout_btn.grid(row=1, column=7)

        self.login_status = tk.Label(login_frame, text="● Not logged in",
                                    font=("Courier New", 8),
                                    fg=VSCODE_RED, bg=VSCODE_SURFACE_ALT)
        self.login_status.grid(row=2, column=0, columnspan=8, sticky=tk.W, pady=(5, 0))

        # === Filter Section ===
        filter_frame = tk.Frame(main_frame, bg=VSCODE_SURFACE_ALT, padx=10, pady=10,
                               relief="solid", borderwidth=2, highlightbackground=VSCODE_BORDER, highlightthickness=2)
        filter_frame.pack(fill=tk.X, pady=(0, 8))

        # Filter title
        filter_title = tk.Label(filter_frame, text="▼ FILTERS",
                               font=("Courier New", 10, "bold"),
                               fg=VSCODE_ORANGE, bg=VSCODE_SURFACE_ALT)
        filter_title.grid(row=0, column=0, columnspan=8, sticky=tk.W, pady=(0, 8))

        # Date row
        date_inner = tk.Frame(filter_frame, bg=VSCODE_SURFACE_ALT)
        date_inner.grid(row=1, column=0, columnspan=8, sticky=tk.W, pady=(0, 5))

        ttk.Label(date_inner, text="Start:").grid(row=0, column=0, sticky=tk.W)
        self.start_date_var = tk.StringVar()
        ttk.Entry(date_inner, textvariable=self.start_date_var, width=12,
                 style="Pixel.TEntry").grid(row=0, column=1, padx=(5, 15))

        ttk.Label(date_inner, text="End:").grid(row=0, column=2, sticky=tk.W)
        self.end_date_var = tk.StringVar()
        ttk.Entry(date_inner, textvariable=self.end_date_var, width=12,
                 style="Pixel.TEntry").grid(row=0, column=3, padx=(5, 15))

        btn_frame = tk.Frame(date_inner, bg=VSCODE_SURFACE_ALT)
        btn_frame.grid(row=0, column=4, padx=2)

        # Pixel style buttons
        for txt, cmd, w in [("This Week", lambda: self.set_quick_date("week"), 9),
                            ("Last Week", lambda: self.set_quick_date("last_week"), 9),
                            ("This Month", lambda: self.set_quick_date("month"), 10)]:
            btn = tk.Button(btn_frame, text=txt, command=cmd,
                           bg=VSCODE_SURFACE_ALT, fg=VSCODE_TEXT,
                           activebackground=VSCODE_BLUE, activeforeground=VSCODE_TEXT,
                           relief="solid", borderwidth=2, highlightbackground=VSCODE_BORDER, highlightthickness=2,
                           font=("Courier New", 8), cursor="hand2", width=w)
            btn.pack(side=tk.LEFT, padx=1)

        # Default dates
        today = datetime.date.today()
        week_ago = today - timedelta(days=7)
        self.start_date_var.set(week_ago.strftime("%Y-%m-%d"))
        self.end_date_var.set(today.strftime("%Y-%m-%d"))

        # Status row
        status_select_frame = tk.Frame(filter_frame, bg=VSCODE_SURFACE_ALT)
        status_select_frame.grid(row=2, column=0, columnspan=8, sticky=tk.W, pady=(0, 5))

        ttk.Label(status_select_frame, text="Status:").grid(row=0, column=0, sticky=tk.W)
        self.status_filter_var = tk.StringVar(value="ALL")
        status_combo = ttk.Combobox(status_select_frame, textvariable=self.status_filter_var,
                                    width=18, state="readonly", style="Pixel.TCombobox")
        status_combo["values"] = ["ALL", "WAIT FAE INFO", "WORKED AROUND", "WORKING",
                                   "CLOSED", "RESOLVED", "WAIT 3RD PARTY"]
        status_combo.grid(row=0, column=1, padx=5, sticky=tk.W)
        status_combo.bind("<<ComboboxSelected>>", lambda e: None)

        # Quick filter info
        filter_info = tk.Label(filter_frame,
                              text="► Includes assigned issues + issues you commented on",
                              font=("Courier New", 8),
                              fg=VSCODE_TEXT_DIM, bg=VSCODE_SURFACE_ALT)
        filter_info.grid(row=3, column=0, columnspan=8, sticky=tk.W, pady=(0, 5))

        # Column order row
        column_order_frame = tk.Frame(filter_frame, bg=VSCODE_SURFACE_ALT)
        column_order_frame.grid(row=4, column=0, columnspan=8, sticky=tk.W, pady=(0, 5))

        ttk.Label(column_order_frame, text="Columns:").grid(row=0, column=0, sticky=tk.W)
        self.column_order_var = tk.StringVar(value="1,2,3,4,5,6,7")
        ttk.Entry(column_order_frame, textvariable=self.column_order_var, width=20,
                 style="Pixel.TEntry").grid(row=0, column=1, padx=5, sticky=tk.W)
        col_help = tk.Label(column_order_frame,
                           text="(1=Cust, 2=Mod, 3=Sum, 4=Jira#, 5=Sts, 6=Key, 7=Prog)",
                           font=("Courier New", 7), fg=VSCODE_TEXT_DIM, bg=VSCODE_SURFACE_ALT)
        col_help.grid(row=0, column=2, sticky=tk.W)

        # Fetch comment toggle
        fetch_comment_frame = tk.Frame(filter_frame, bg=VSCODE_SURFACE_ALT)
        fetch_comment_frame.grid(row=5, column=0, columnspan=8, sticky=tk.W, pady=0)
        self.fetch_comment_var = tk.BooleanVar(value=False)
        fetch_cb = tk.Checkbutton(fetch_comment_frame, text="[ ] Fetch latest comment for Progress",
                                 variable=self.fetch_comment_var,
                                 bg=VSCODE_SURFACE_ALT, fg=VSCODE_TEXT,
                                 selectcolor=VSCODE_CYAN, font=("Courier New", 9),
                                 cursor="hand2")
        fetch_cb.pack(side=tk.LEFT)

        # === Output Section ===
        output_frame = tk.Frame(main_frame, bg=VSCODE_SURFACE_ALT, padx=10, pady=10,
                               relief="solid", borderwidth=2, highlightbackground=VSCODE_BORDER, highlightthickness=2)
        output_frame.pack(fill=tk.X, pady=(0, 8))

        output_title = tk.Label(output_frame, text="▼ OUTPUT",
                               font=("Courier New", 10, "bold"),
                               fg=VSCODE_ORANGE, bg=VSCODE_SURFACE_ALT)
        output_title.grid(row=0, column=0, columnspan=8, sticky=tk.W, pady=(0, 5))

        output_inner = tk.Frame(output_frame, bg=VSCODE_SURFACE_ALT)
        output_inner.grid(row=1, column=0, columnspan=8, sticky=tk.EW)
        output_inner.columnconfigure(1, weight=1)

        ttk.Label(output_inner, text="Save:").grid(row=0, column=0, sticky=tk.W)
        self.filepath_var = tk.StringVar()
        self.filepath_var.set(os.path.join(self.last_save_dir,
                         f"Jira_Weekly_Report_{datetime.date.today().strftime('%Y%m%d')}.xlsx"))
        ttk.Entry(output_inner, textvariable=self.filepath_var, width=45,
                 style="Pixel.TEntry").grid(row=0, column=1, padx=5, sticky=tk.EW)
        ttk.Button(output_inner, text="...", command=self.browse_file, width=4,
                  style="Pixel.TButton").grid(row=0, column=2, padx=(0, 5))

        # === Generate Button ===
        gen_frame = tk.Frame(main_frame, bg=VSCODE_SURFACE)
        gen_frame.pack(pady=5)

        self.generate_btn = tk.Button(
            gen_frame,
            text="▶ GENERATE REPORT",
            command=self.generate_report,
            state=tk.DISABLED,
            bg=VSCODE_BLUE, fg=VSCODE_TEXT,
            activebackground=VSCODE_SELECT, activeforeground=VSCODE_TEXT,
            relief="solid", borderwidth=3, highlightbackground=VSCODE_BORDER, highlightthickness=3,
            font=("Courier New", 12, "bold"), cursor="hand2", padx=20, pady=5
        )
        self.generate_btn.pack()

        # === Status Bar ===
        self.status_bar = tk.Label(
            self.root,
            text="► Ready - Please login to continue",
            bd=2,
            relief="solid",
            anchor=tk.W,
            padx=10,
            font=("Courier New", 9),
            fg=VSCODE_TEXT_DIM,
            bg=VSCODE_BG,
            borderwidth=2,
            highlightbackground=VSCODE_BORDER
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def set_quick_date(self, period):
        today = datetime.date.today()
        if period == "week":
            days_since_monday = today.weekday()
            start = today - timedelta(days=days_since_monday)
            end = start + timedelta(days=6)
        elif period == "last_week":
            days_since_monday = today.weekday()
            this_monday = today - timedelta(days=days_since_monday)
            last_monday = this_monday - timedelta(days=7)
            last_sunday = last_monday + timedelta(days=6)
            start = last_monday
            end = last_sunday
        else:
            start = today.replace(day=1)
            _, last_day = calendar.monthrange(today.year, today.month)
            end = today.replace(day=last_day)
        self.start_date_var.set(start.strftime("%Y-%m-%d"))
        self.end_date_var.set(end.strftime("%Y-%m-%d"))

    def browse_file(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=self.last_save_dir,
            initialfile=os.path.basename(self.filepath_var.get())
        )
        if file_path:
            self.filepath_var.set(file_path)
            self.last_save_dir = os.path.dirname(file_path)

    def toggle_password_visibility(self):
        if self.show_password_var.get():
            self.password_entry.config(show="")
        else:
            self.password_entry.config(show="*")

    def login(self):
        url = self.base_url
        username = self.username_var.get().strip()
        password = self.password_var.get()

        if not username or not password:
            messagebox.showerror("Error", "Please enter username and password")
            return

        self.login_btn.config(state=tk.DISABLED)
        self.update_status("Logging in...")

        import threading
        thread = threading.Thread(target=self._login_thread, args=(url, username, password))
        thread.daemon = True
        thread.start()

    def _login_thread(self, url, username, password):
        try:
            result = self._do_api_login(url, username, password)
            self.root.after(0, lambda: self._handle_login_result(result))
        except Exception as e:
            self.root.after(0, lambda: self._handle_login_error(str(e)))

    def _do_api_login(self, url, username, password):
        try:
            auth = requests.auth.HTTPBasicAuth(username, password)
            response = self.session.get(f"{url}/rest/api/2/myself", auth=auth, timeout=30)

            if response.status_code == 200:
                user_data = response.json()
                self.user_email = user_data.get("email", "")
                self.session.auth = auth
                return {"success": True, "username": username}
            elif response.status_code == 401:
                return self._do_cookie_login(url, username, password)
            else:
                return {"success": False, "error": f"API returned status {response.status_code}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _do_cookie_login(self, url, username, password):
        try:
            login_page = self.session.get(f"{url}/login.jsp", timeout=30)
            atl_token_match = re.search(
                r'name="atl_token"\s*type="hidden"\s*value="([^"]+)"',
                login_page.text
            )
            atl_token = atl_token_match.group(1) if atl_token_match else ""

            form_data = {
                "os_username": username,
                "os_password": password,
                "os_destination": "/",
                "atl_token": atl_token,
                "user_role": "",
                "os_cookie": "true"
            }
            login_response = self.session.post(
                f"{url}/dologin.jsp",
                data=form_data,
                timeout=30,
                allow_redirects=True
            )

            if "invalid" in login_response.text.lower() or "incorrect" in login_response.text.lower():
                return {"success": False, "error": "Invalid username or password"}

            api_check = self.session.get(f"{url}/rest/api/2/myself", timeout=30)
            if api_check.status_code == 200:
                user_data = api_check.json()
                self.user_email = user_data.get("email", "")
                return {"success": True, "username": username}
            else:
                return {"success": False, "error": f"Verification failed (status: {api_check.status_code})"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _handle_login_result(self, result):
        self.login_btn.config(state=tk.NORMAL)
        if result["success"]:
            self.username = result["username"]
            self.on_login_success(self.username)
        else:
            messagebox.showerror("Login Failed", result["error"])
            self.update_status("Login failed")

    def _handle_login_error(self, error):
        self.login_btn.config(state=tk.NORMAL)
        messagebox.showerror("Connection Error", f"Cannot connect to Jira server:\n{error}")
        self.update_status("Connection error")

    def on_login_success(self, username):
        self.logged_in = True
        self.username = username

        if self.remember_var.get():
            self.save_credentials(username, self.password_var.get())

        self.login_status.config(text=f"● Logged in: {username}", fg=VSCODE_GREEN)
        self.login_btn.config(state=tk.DISABLED)
        self.logout_btn.config(state=tk.NORMAL)
        self.generate_btn.config(state=tk.NORMAL)
        self.update_status(f"► Logged in as {username}")
        messagebox.showinfo("Success", "Login successful!")

    def logout(self):
        if self.logged_in:
            try:
                self.session.delete(f"{self.base_url}/rest/auth/1/session")
            except:
                pass
            self.logged_in = False
            self.username = None
            self.user_email = None
            self.login_status.config(text="● Not logged in", fg=VSCODE_RED)
            self.login_btn.config(state=tk.NORMAL)
            self.logout_btn.config(state=tk.DISABLED)
            self.generate_btn.config(state=tk.DISABLED)
            self.update_status("► Logged out")

    def update_status(self, message):
        self.status_bar.config(text=f"► {message}")
        self.root.update_idletasks()

    def generate_report(self):
        if not self.logged_in:
            messagebox.showerror("Error", "Please login first")
            return

        try:
            start_date = datetime.datetime.strptime(self.start_date_var.get(), "%Y-%m-%d").date()
            end_date = datetime.datetime.strptime(self.end_date_var.get(), "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("Error", "Invalid date format. Use YYYY-MM-DD")
            return

        if end_date < start_date:
            messagebox.showerror("Error", "End date must be after start date")
            return

        selected_status = self.status_filter_var.get().strip()
        if selected_status == "ALL":
            status_clause = ""
        else:
            status_clause = f'status = "{selected_status}" '

        engineer_field = "Software Development Engineer 软件开发工程师"

        filepath = self.filepath_var.get().strip()
        if not filepath:
            messagebox.showerror("Error", "Please select a save path")
            return
        if not filepath.endswith(".xlsx"):
            filepath += ".xlsx"

        save_dir = os.path.dirname(filepath)
        if save_dir and not os.path.exists(save_dir):
            os.makedirs(save_dir)

        self.update_status("Fetching issues...")

        try:
            # Normal issues: engineer = currentUser, updated within date range
            jql_normal = f'"{engineer_field}" IN (currentUser()) AND updated >= {start_date} AND updated <= "{end_date} 23:59"'
            if status_clause:
                jql_normal += f' AND {status_clause}'

            # WAIT 3RD PARTY issues: engineer = currentUser, status = WAIT 3RD PARTY (no date limit)
            jql_wait3rd = f'"{engineer_field}" IN (currentUser()) AND status = "WAIT 3RD PARTY"'
            if status_clause:
                jql_wait3rd += f' AND {status_clause}'

            # Assist: normal issues
            jql_assist_normal = f'comment ~ currentUser() AND "{engineer_field}" != currentUser() AND updated >= {start_date} AND updated <= "{end_date} 23:59"'
            if status_clause:
                jql_assist_normal += f' AND {status_clause}'

            # Assist: WAIT 3RD PARTY issues
            jql_assist_wait3rd = f'comment ~ currentUser() AND "{engineer_field}" != currentUser() AND status = "WAIT 3RD PARTY"'
            if status_clause:
                jql_assist_wait3rd += f' AND {status_clause}'

            self.update_status("Searching assigned issues...")
            issues_assigned_normal = self.fetch_issues(jql_normal)
            issues_assigned_wait3rd = self.fetch_issues(jql_wait3rd)
            issues_assigned = issues_assigned_normal + issues_assigned_wait3rd
            self.update_status(f"Found {len(issues_assigned_normal)} normal + {len(issues_assigned_wait3rd)} WAIT_3RD assigned")

            issues_assist_normal = self.fetch_issues(jql_assist_normal)
            issues_assist_wait3rd = self.fetch_issues(jql_assist_wait3rd)
            issues_assist = issues_assist_normal + issues_assist_wait3rd
            self.update_status(f"Found {len(issues_assist_normal)} normal + {len(issues_assist_wait3rd)} WAIT_3RD assist")

            # Filter: for closed issues, only skip if no comment in 3 months
            # others require comment in date range
            # WAIT 3RD PARTY requires no comment (waiting for external)
            closed_statuses = {"CLOSED", "RESOLVED"}
            no_comment_required_statuses = {"WAIT 3RD PARTY"}
            self.update_status(f"Filtering {len(issues_assigned)} assigned issues...")
            issues_assigned_filtered = []
            for issue in issues_assigned:
                status = issue.get("fields", {}).get("status", {}).get("name", "")
                if status in closed_statuses:
                    # Closed issues: skip only if user has NOT commented in 3 months
                    if not self.user_commented_within_months(issue['key'], months=3):
                        self.update_status(f"Skipping {issue['key']} - {status} (no comment in 3 months)")
                        continue
                    issues_assigned_filtered.append(issue)
                elif status in no_comment_required_statuses:
                    # WAIT 3RD PARTY: include regardless of comments
                    issues_assigned_filtered.append(issue)
                elif self.user_commented_in_date_range(issue['key'], start_date, end_date):
                    issues_assigned_filtered.append(issue)
                else:
                    self.update_status(f"Skipping {issue['key']} - no comment in date range")
            issues_assigned = issues_assigned_filtered

            self.update_status("Filtering assist issues...")
            issues_assist_filtered = []
            for issue in issues_assist:
                status = issue.get("fields", {}).get("status", {}).get("name", "")
                if status in closed_statuses:
                    # Closed issues: skip only if user has NOT commented in 3 months
                    if not self.user_commented_within_months(issue['key'], months=3):
                        self.update_status(f"Skipping {issue['key']} - {status} (no comment in 3 months)")
                        continue
                    issues_assist_filtered.append(issue)
                elif status in no_comment_required_statuses:
                    # WAIT 3RD PARTY: include regardless of comments
                    issues_assist_filtered.append(issue)
                elif self.user_commented_in_date_range(issue['key'], start_date, end_date):
                    issues_assist_filtered.append(issue)
                else:
                    self.update_status(f"Skipping {issue['key']} - no comment in date range")
            issues_assist = issues_assist_filtered

            all_issues = {issue['key']: issue for issue in issues_assigned + issues_assist}
            issues = list(all_issues.values())

            # Sort issues by status: CLOSED -> RESOLVED -> WORKING -> others
            status_order = {
                "CLOSED": 0,
                "RESOLVED": 1,
                "WORKING": 2,
                "WORKED AROUND": 3,
                "WAIT FAE INFO": 4,
                "WAIT 3RD PARTY": 5,
            }
            issues.sort(key=lambda x: (
                status_order.get(x.get("fields", {}).get("status", {}).get("name", ""), 99),
                x.get("key", "")
            ))

            self.update_status(f"Found {len(issues)} total issues (assigned + assists)")

            self.update_status("Generating Excel file...")
            self.create_excel(issues, filepath, selected_status, start_date, end_date)
            self.update_status(f"Report saved: {filepath}")
            messagebox.showinfo("Success", f"Report generated successfully!\n\n{len(issues)} issues exported to:\n{filepath}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed:\n{str(e)}")
            self.update_status("Failed")

    def fetch_issues(self, jql, start_at=0, max_results=100):
        all_issues = []
        url = f"{self.base_url}/rest/api/2/search"

        params = {
            "jql": jql,
            "startAt": start_at,
            "maxResults": max_results,
            "fields": "summary,status,priority,created,updated,creator,key,customfield_11029,customfield_12031"
        }

        while True:
            try:
                response = self.session.get(url, params=params, timeout=30)

                if response.status_code >= 400:
                    error_detail = response.text[:500] if response.text else "No details"
                    raise Exception(f"Error {response.status_code}: {error_detail}")

                response.raise_for_status()
                data = response.json()

                issues = data.get("issues", [])
                all_issues.extend(issues)

                total = data.get("total", 0)
                if start_at + len(issues) >= total:
                    break

                start_at += max_results
                self.update_status(f"Fetching {len(all_issues)}/{total}...")

            except requests.exceptions.RequestException as e:
                raise Exception(f"Fetch error: {str(e)}")

        return all_issues

    def user_commented_in_date_range(self, issue_key, start_date, end_date):
        try:
            url = f"{self.base_url}/rest/api/2/issue/{issue_key}/comment"
            response = self.session.get(url, timeout=30)

            if response.status_code != 200:
                return False

            data = response.json()
            comments = data.get("comments", [])

            current_user_short = self.username.split("@")[0] if self.username else ""

            for comment in comments:
                author = comment.get("author", {})
                author_email = author.get("emailAddress", "")
                author_short = author_email.split("@")[0] if author_email else ""

                is_current_user = (current_user_short and (
                    current_user_short == author_short or
                    self.username == author_email
                ))

                if is_current_user:
                    created_str = comment.get("created", "")
                    if created_str:
                        comment_date = datetime.datetime.strptime(created_str[:19], "%Y-%m-%dT%H:%M:%S").date()
                        if start_date <= comment_date <= end_date:
                            return True

            return False
        except Exception:
            return False

    def user_commented_within_months(self, issue_key, months=3):
        """Check if current user commented on this issue within the last N months"""
        try:
            url = f"{self.base_url}/rest/api/2/issue/{issue_key}/comment"
            response = self.session.get(url, timeout=30)

            if response.status_code != 200:
                return False

            data = response.json()
            comments = data.get("comments", [])

            current_user_short = self.username.split("@")[0] if self.username else ""
            since_date = datetime.date.today() - timedelta(days=months * 30)

            for comment in comments:
                author = comment.get("author", {})
                author_email = author.get("emailAddress", "")
                author_short = author_email.split("@")[0] if author_email else ""

                is_current_user = (current_user_short and (
                    current_user_short == author_short or
                    self.username == author_email
                ))

                if is_current_user:
                    created_str = comment.get("created", "")
                    if created_str:
                        comment_date = datetime.datetime.strptime(created_str[:19], "%Y-%m-%dT%H:%M:%S").date()
                        if comment_date >= since_date:
                            return True

            return False
        except Exception:
            return False

    def get_user_latest_comment(self, issue_key, start_date, end_date):
        try:
            url = f"{self.base_url}/rest/api/2/issue/{issue_key}/comment"
            response = self.session.get(url, timeout=30)

            if response.status_code != 200:
                return None

            data = response.json()
            comments = data.get("comments", [])

            current_user_short = self.username.split("@")[0] if self.username else ""

            latest_comment = None
            latest_date = None

            for comment in comments:
                author = comment.get("author", {})
                author_email = author.get("emailAddress", "")
                author_short = author_email.split("@")[0] if author_email else ""

                is_current_user = (current_user_short and (
                    current_user_short == author_short or
                    self.username == author_email
                ))

                if is_current_user:
                    created_str = comment.get("created", "")
                    if created_str:
                        comment_date = datetime.datetime.strptime(created_str[:19], "%Y-%m-%dT%H:%M:%S").date()
                        if start_date <= comment_date <= end_date:
                            body = comment.get("body", "")
                            text = re.sub(r'<[^>]+>', '', body).strip()
                            if text and (latest_date is None or comment_date > latest_date):
                                latest_date = comment_date
                                latest_comment = text

            return latest_comment
        except Exception:
            return None

    def create_excel(self, issues, filepath, statuses, start_date, end_date):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Report"

        font_chinese = Font(name="Microsoft YaHei", size=10)
        font_english = Font(name="JetBrains Mono", size=10)
        font_header = Font(name="Microsoft YaHei", bold=True, color="FF000000", size=10)
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        col_order = [int(x.strip()) - 1 for x in self.column_order_var.get().split(",")]

        header_names = ["客户名称 Customer", "模组型号 Module", "问题描述 Issue Description",
                       "单号 Jira Number", "状态 Status", "是否为重点问题 Is Key Issue",
                       "进展 Progress"]

        def has_chinese(text):
            return any('一' <= c <= '鿿' for c in str(text))

        def set_cell_font(cell, value):
            text = str(value) if value is not None else ""
            cell.value = text if text else value
            cell.font = font_chinese if has_chinese(text) else font_english

        for col, idx in enumerate(col_order, 1):
            header = header_names[idx]
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = font_header
            cell.alignment = header_alignment

        status_col = col_order.index(4) + 1
        key_issue_col = col_order.index(5) + 1
        status_range = f"{get_column_letter(status_col)}2:{get_column_letter(status_col)}{len(issues) + 1}"
        key_issue_range = f"{get_column_letter(key_issue_col)}2:{get_column_letter(key_issue_col)}{len(issues) + 1}"

        ws_options = wb.create_sheet("_Options")
        ws_options.sheet_state = "hidden"
        for i, opt in enumerate(["WAIT FAE INFO", "WORKED AROUND", "WORKING", "CLOSED", "RESOLVED", "WAIT 3RD PARTY"], 1):
            ws_options.cell(row=i, column=1, value=opt)
        for i, opt in enumerate(["是", "否"], 1):
            ws_options.cell(row=i, column=2, value=opt)
        status_options_range = f"_Options!$A$1:$A$6"
        key_issue_options_range = f"_Options!$B$1:$B$2"

        dv_status = DataValidation(type="list", formula1=status_options_range, allow_blank=True)
        dv_status.error = "Please select a valid status"
        dv_status.errorTitle = "Invalid Status"
        ws.add_data_validation(dv_status)
        dv_status.sqref = status_range

        dv_key_issue = DataValidation(type="list", formula1=key_issue_options_range, allow_blank=True)
        dv_key_issue.error = "Please select 是 or 否"
        dv_key_issue.errorTitle = "Invalid Key Issue"
        ws.add_data_validation(dv_key_issue)
        dv_key_issue.sqref = key_issue_range

        for row, issue in enumerate(issues, 2):
            fields = issue.get("fields", {})
            issue_key = issue.get("key", "")

            latest_comment = ""
            if self.fetch_comment_var.get():
                self.update_status(f"Fetching comment for {issue_key}...")
                latest_comment = self.get_user_latest_comment(issue_key, start_date, end_date) or ""

            values = [
                fields.get("customfield_11029", ""),
                fields.get("customfield_12031", ""),
                fields.get("summary", ""),
                issue_key,
                fields.get("status", {}),
                fields.get("priority", {}),
                latest_comment,
            ]

            for col, idx in enumerate(col_order, 1):
                val = values[idx]
                if idx == 4:
                    val = val.get("name", "") if isinstance(val, dict) else val
                elif idx == 5:
                    priority_name = val.get("name", "") if isinstance(val, dict) else val
                    val = "是" if priority_name in ("Highest", "High") else "否"
                elif idx == 1:
                    module_field = val
                    module = ""
                    if isinstance(module_field, dict):
                        child = module_field.get("child", {})
                        if isinstance(child, dict):
                            module = child.get("value", "")
                        elif not module:
                            module = module_field.get("value", "")
                    elif isinstance(module_field, list) and len(module_field) > 0:
                        parts = []
                        for item in module_field:
                            if isinstance(item, dict):
                                child = item.get("child", {})
                                if isinstance(child, dict):
                                    parts.append(child.get("value", ""))
                                else:
                                    parts.append(item.get("value", str(item)))
                            else:
                                parts.append(str(item))
                        module = " - ".join(parts)
                    val = module

                cell = ws.cell(row=row, column=col, value=val)
                if idx == 3:
                    cell.hyperlink = f"{self.base_url}/browse/{issue_key}"
                set_cell_font(cell, val)
                cell.alignment = cell_alignment

        for col in range(1, 8):
            max_length = 0
            column_letter = get_column_letter(col)
            for row in range(1, ws.max_row + 1):
                cell = ws.cell(row=row, column=col)
                try:
                    if cell.value:
                        cell_len = len(str(cell.value))
                        max_length = max(max_length, cell_len)
                except:
                    pass
            adjusted_width = min(max_length + 5, 60)
            ws.column_dimensions[column_letter].width = adjusted_width

        ws.row_dimensions[1].height = 25

        wb.save(filepath)


def main():
    root = tk.Tk()
    app = JiraReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
