"""
Jira Weekly Report Generator
Desktop app to generate Excel reports from Jira issues
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


class JiraReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Jira Report")
        self.root.geometry("700x550")
        self.root.minsize(600, 450)

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

    def setup_ui(self):
        # Title
        title_label = tk.Label(
            self.root,
            text="Jira Report",
            font=("Segoe UI", 16, "bold"),
            fg="#1E5AA8"
        )
        title_label.pack(pady=(15, 5))

        # Main container
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # === Login Section ===
        login_frame = ttk.LabelFrame(main_frame, text="Login", padding="10")
        login_frame.pack(fill=tk.X, pady=(0, 10))

        login_inner = ttk.Frame(login_frame)
        login_inner.pack(fill=tk.X)

        ttk.Label(login_inner, text="User:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.username_var = tk.StringVar(value="")
        ttk.Entry(login_inner, textvariable=self.username_var, width=20).grid(row=0, column=1, padx=(0, 10))

        ttk.Label(login_inner, text="Pwd:").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        self.password_var = tk.StringVar(value="")
        self.password_entry = ttk.Entry(login_inner, textvariable=self.password_var, show="*", width=20)
        self.password_entry.grid(row=0, column=3, padx=(0, 5))

        # Show/hide password toggle
        self.show_password_var = tk.BooleanVar(value=False)
        self.show_pwd_check = ttk.Checkbutton(login_inner, text="👁", variable=self.show_password_var, command=self.toggle_password_visibility, width=2)
        self.show_pwd_check.grid(row=0, column=4, padx=(0, 5))

        self.remember_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(login_inner, text="Remember", variable=self.remember_var).grid(row=0, column=5, padx=(5, 10))

        self.login_btn = ttk.Button(login_inner, text="Login", command=self.login, width=8)
        self.login_btn.grid(row=0, column=6, padx=(0, 5))
        self.logout_btn = ttk.Button(login_inner, text="Logout", command=self.logout, state=tk.DISABLED, width=8)
        self.logout_btn.grid(row=0, column=7)

        self.login_status = tk.Label(login_inner, text="Not logged in", font=("Segoe UI", 8), fg="#6B778C")
        self.login_status.grid(row=1, column=0, columnspan=8, sticky=tk.W, pady=(5, 0))

        # Track status_vars for backward compatibility (may be referenced elsewhere)
        self.status_vars = {}

        # === Filter Section ===
        filter_frame = ttk.LabelFrame(main_frame, text="Filters", padding="10")
        filter_frame.pack(fill=tk.X, pady=(0, 10))

        # Date row
        date_inner = ttk.Frame(filter_frame)
        date_inner.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(date_inner, text="Start:").grid(row=0, column=0, sticky=tk.W)
        self.start_date_var = tk.StringVar()
        ttk.Entry(date_inner, textvariable=self.start_date_var, width=12).grid(row=0, column=1, padx=(5, 10))

        ttk.Label(date_inner, text="End:").grid(row=0, column=2, sticky=tk.W)
        self.end_date_var = tk.StringVar()
        ttk.Entry(date_inner, textvariable=self.end_date_var, width=12).grid(row=0, column=3, padx=(5, 10))

        btn_frame = ttk.Frame(date_inner)
        btn_frame.grid(row=0, column=4, padx=2)
        ttk.Button(btn_frame, text="This Week", command=lambda: self.set_quick_date("week"), width=9).pack(side=tk.LEFT, padx=1)
        ttk.Button(btn_frame, text="Last Week", command=lambda: self.set_quick_date("last_week"), width=9).pack(side=tk.LEFT, padx=1)
        ttk.Button(btn_frame, text="This Month", command=lambda: self.set_quick_date("month"), width=10).pack(side=tk.LEFT, padx=1)

        # Default dates
        today = datetime.date.today()
        week_ago = today - timedelta(days=7)
        self.start_date_var.set(week_ago.strftime("%Y-%m-%d"))
        self.end_date_var.set(today.strftime("%Y-%m-%d"))

        # Status row - now a single status dropdown for precise filtering
        status_select_frame = ttk.Frame(filter_frame)
        status_select_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(status_select_frame, text="Status Filter:").grid(row=0, column=0, sticky=tk.W)
        self.status_filter_var = tk.StringVar(value="ALL")
        status_combo = ttk.Combobox(status_select_frame, textvariable=self.status_filter_var, width=20, state="readonly")
        status_combo["values"] = ["ALL", "WAIT FAE INFO", "WORKED AROUND", "WORKING", "CLOSED", "RESOLVED", "WAIT 3RD PARTY"]
        status_combo.grid(row=0, column=1, padx=5, sticky=tk.W)
        status_combo.bind("<<ComboboxSelected>>", lambda e: None)

        # Quick filter info
        filter_info = tk.Label(filter_frame, text="💡 Includes both assigned issues and issues you commented on", font=("Segoe UI", 8), fg="#6B778C")
        filter_info.pack(fill=tk.X, pady=(0, 5))

        # Column order row
        column_order_frame = ttk.Frame(filter_frame)
        column_order_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(column_order_frame, text="Column Order:").grid(row=0, column=0, sticky=tk.W)
        self.column_order_var = tk.StringVar(value="1,2,3,4,5,6,7")
        ttk.Entry(column_order_frame, textvariable=self.column_order_var, width=25).grid(row=0, column=1, padx=5, sticky=tk.W)
        ttk.Label(column_order_frame, text="(1=Customer, 2=Module, 3=Summary, 4=Jira#, 5=Status, 6=Priority, 7=Progress)").grid(row=0, column=2, sticky=tk.W)

        # === Output Section ===
        output_frame = ttk.LabelFrame(main_frame, text="Output", padding="10")
        output_frame.pack(fill=tk.X, pady=(0, 10))

        output_inner = ttk.Frame(output_frame)
        output_inner.pack(fill=tk.X)

        ttk.Label(output_inner, text="Save:").grid(row=0, column=0, sticky=tk.W)
        self.filepath_var = tk.StringVar()
        self.filepath_var.set(os.path.join(self.last_save_dir, f"Jira_Weekly_Report_{datetime.date.today().strftime('%Y%m%d')}.xlsx"))
        ttk.Entry(output_inner, textvariable=self.filepath_var, width=50).grid(row=0, column=1, padx=5, sticky=tk.EW)
        ttk.Button(output_inner, text="...", command=self.browse_file, width=4).grid(row=0, column=2, padx=(0, 5))

        output_inner.columnconfigure(1, weight=1)

        # === Generate Button ===
        self.generate_btn = ttk.Button(
            main_frame,
            text="Generate Report",
            command=self.generate_report,
            state=tk.DISABLED
        )
        self.generate_btn.pack(pady=(5, 0))

        # === Status Bar ===
        self.status_bar = tk.Label(
            self.root,
            text="Ready - Please login to continue",
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W,
            padx=10,
            font=("Segoe UI", 8)
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def set_quick_date(self, period):
        today = datetime.date.today()
        if period == "week":
            # Monday = 0, Sunday = 6
            days_since_monday = today.weekday()
            start = today - timedelta(days=days_since_monday)
            # End = next Sunday (6 days after Monday)
            end = start + timedelta(days=6)
        elif period == "last_week":
            days_since_monday = today.weekday()
            this_monday = today - timedelta(days=days_since_monday)
            last_monday = this_monday - timedelta(days=7)
            last_sunday = last_monday + timedelta(days=6)
            start = last_monday
            end = last_sunday
        else:
            # This month - use calendar for proper month end handling (leap years, etc.)
            start = today.replace(day=1)
            # Get last day of month using calendar.monthrange
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

        # Disable login button to prevent double-click
        self.login_btn.config(state=tk.DISABLED)
        self.update_status("Logging in...")

        import threading
        thread = threading.Thread(target=self._login_thread, args=(url, username, password))
        thread.daemon = True
        thread.start()

    def _login_thread(self, url, username, password):
        """Login in background thread to keep UI responsive"""
        try:
            result = self._do_api_login(url, username, password)
            self.root.after(0, lambda: self._handle_login_result(result))
        except Exception as e:
            self.root.after(0, lambda: self._handle_login_error(str(e)))

    def _do_api_login(self, url, username, password):
        """Login using Jira REST API with Basic Auth or session"""
        try:
            # Try Basic Auth first (simplest approach)
            auth = requests.auth.HTTPBasicAuth(username, password)
            response = self.session.get(f"{url}/rest/api/2/myself", auth=auth, timeout=30)

            if response.status_code == 200:
                user_data = response.json()
                self.user_email = user_data.get("email", "")
                # Keep auth for subsequent requests
                self.session.auth = auth
                return {"success": True, "username": username}
            elif response.status_code == 401:
                # Basic Auth failed, try cookie-based login
                return self._do_cookie_login(url, username, password)
            else:
                return {"success": False, "error": f"API returned status {response.status_code}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _do_cookie_login(self, url, username, password):
        """Cookie-based login (form POST)"""
        try:
            # Get login page to extract CSRF token
            login_page = self.session.get(f"{url}/login.jsp", timeout=30)
            atl_token_match = re.search(
                r'name="atl_token"\s*type="hidden"\s*value="([^"]+)"',
                login_page.text
            )
            atl_token = atl_token_match.group(1) if atl_token_match else ""

            # Submit login form
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

            # Verify authentication
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

        self.login_status.config(text=f"Logged in as: {username}", fg="#00A906")
        self.login_btn.config(state=tk.DISABLED)
        self.logout_btn.config(state=tk.NORMAL)
        self.generate_btn.config(state=tk.NORMAL)
        self.update_status(f"Logged in as {username}")
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
            self.login_status.config(text="Not logged in", fg="#6B778C")
            self.login_btn.config(state=tk.NORMAL)
            self.logout_btn.config(state=tk.DISABLED)
            self.generate_btn.config(state=tk.DISABLED)
            self.update_status("Logged out")

    def update_status(self, message):
        self.status_bar.config(text=message)
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

        # Hardcoded engineer field - not exposed in UI
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
            # Fetch both: issues assigned to me AND issues I commented on but was not the engineer
            # Query 1: Engineer assigned to me
            jql_assigned = f'"{engineer_field}" IN (currentUser()) AND updated >= {start_date} AND updated <= "{end_date} 23:59"'
            if status_clause:
                jql_assigned += f' AND {status_clause}'

            # Query 2: I commented on but not my assignment
            jql_assist = f'comment ~ currentUser() AND "{engineer_field}" != currentUser() AND updated >= {start_date} AND updated <= "{end_date} 23:59"'
            if status_clause:
                jql_assist += f' AND {status_clause}'

            self.update_status(f"Searching assigned issues...")
            issues_assigned = self.fetch_issues(jql_assigned)
            self.update_status(f"Found {len(issues_assigned)} assigned issues, searching assist issues...")
            issues_assist = self.fetch_issues(jql_assist)

            # Merge and deduplicate by issue key
            all_issues = {issue['key']: issue for issue in issues_assigned + issues_assist}
            issues = list(all_issues.values())

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

    def create_excel(self, issues, filepath, statuses, start_date, end_date):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Report"

        # Fonts
        font_chinese = Font(name="Microsoft YaHei", size=10)
        font_english = Font(name="JetBrains Mono", size=10)
        font_header = Font(name="Microsoft YaHei", bold=True, color="FFFFFFFF", size=10)

        header_fill = PatternFill(start_color="FF1E5AA8", end_color="FF1E5AA8", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        # Parse column order (1-based to 0-based index)
        # 1=Customer, 2=Module, 3=Summary, 4=Jira#, 5=Status, 6=Priority, 7=Progress
        col_order = [int(x.strip()) - 1 for x in self.column_order_var.get().split(",")]

        header_keys = ["customfield_11029", "customfield_12031", "summary", "key", "status", "priority", "progress"]
        header_names = ["Customer 客户名称", "Module 模组型号", "Issue Description 问题描述", "Jira Number 单号", "Status 状态", "Priority 优先级", "Progress 进展"]

        def has_chinese(text):
            return any('一' <= c <= '鿿' for c in str(text))

        def set_cell_font(cell, value):
            """Set cell value with appropriate font based on content"""
            text = str(value) if value is not None else ""
            cell.value = text if text else value
            cell.font = font_chinese if has_chinese(text) else font_english

        # Write headers in custom order
        for col, idx in enumerate(col_order, 1):
            header = header_names[idx]
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = font_header
            cell.fill = header_fill
            cell.alignment = header_alignment

        # Add dropdown validation for Status column (determined by column order)
        status_col = col_order.index(4) + 1
        status_range = f"{get_column_letter(status_col)}2:{get_column_letter(status_col)}{len(issues) + 1}"

        # Create hidden sheet with status options for validation reference
        ws_options = wb.create_sheet("_Options")
        ws_options.sheet_state = "hidden"
        for i, opt in enumerate(["WAIT FAE INFO", "WORKED AROUND", "WORKING", "CLOSED", "RESOLVED", "WAIT 3RD PARTY"], 1):
            ws_options.cell(row=i, column=1, value=opt)
        options_range = f"_Options!$A$1:$A$6"

        dv = DataValidation(type="list", formula1=options_range, allow_blank=True)
        dv.error = "Please select a valid status"
        dv.errorTitle = "Invalid Status"
        ws.add_data_validation(dv)
        dv.sqref = status_range

        for row, issue in enumerate(issues, 2):
            fields = issue.get("fields", {})
            values = [
                fields.get("customfield_11029", ""),
                fields.get("customfield_12031", ""),
                fields.get("summary", ""),
                issue.get("key", ""),
                fields.get("status", {}),
                fields.get("priority", {}),
                "",
            ]

            for col, idx in enumerate(col_order, 1):
                val = values[idx]
                # Handle status and priority dicts
                if idx == 4:  # Status
                    val = val.get("name", "") if isinstance(val, dict) else val
                elif idx == 5:  # Priority
                    val = val.get("name", "") if isinstance(val, dict) else val
                elif idx == 1:  # Module (customfield_12031)
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
                if idx == 3:  # Jira# column - add hyperlink
                    key = issue.get("key", "")
                    cell.hyperlink = f"{self.base_url}/browse/{key}"
                set_cell_font(cell, val)
                cell.alignment = cell_alignment

        # Auto-fit column widths
        for col in range(1, 8):
            max_length = 0
            column_letter = get_column_letter(col)
            for row in range(1, ws.max_row + 1):
                cell = ws.cell(row=row, column=col)
                try:
                    if cell.value:
                        # Handle both string and hyperlink values
                        cell_len = len(str(cell.value))
                        max_length = max(max_length, cell_len)
                except:
                    pass
            adjusted_width = min(max_length + 5, 60)  # Cap at 60
            ws.column_dimensions[column_letter].width = adjusted_width

        ws.row_dimensions[1].height = 25

        wb.save(filepath)


def main():
    root = tk.Tk()
    app = JiraReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()