"""
Jira API Client - Reusable Jira connection and data fetching
"""
import requests
import requests.auth
import datetime
import re
from datetime import timedelta
import os


class JiraClient:
    def __init__(self, base_url: str, username: str, password: str):
        self.base_url = base_url
        self.username = username
        self.password = password
        self.session = requests.Session()
        self.logged_in = False
        self.user_email = None

    def login(self):
        """Try API login, fall back to cookie login"""
        result = self._do_api_login()
        if result.get("success"):
            return True
        return self._do_cookie_login()

    def _do_api_login(self):
        try:
            auth = requests.auth.HTTPBasicAuth(self.username, self.password)
            response = self.session.get(f"{self.base_url}/rest/api/2/myself", auth=auth, timeout=30)
            if response.status_code == 200:
                user_data = response.json()
                self.user_email = user_data.get("email", "")
                self.session.auth = auth
                self.logged_in = True
                return {"success": True, "username": self.username}
            elif response.status_code == 401:
                return {"success": False, "needs_cookie_login": True}
            else:
                return {"success": False, "error": f"API returned status {response.status_code}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _do_cookie_login(self):
        try:
            login_page = self.session.get(f"{self.base_url}/login.jsp", timeout=30)
            atl_token_match = re.search(
                r'name="atl_token"\s*type="hidden"\s*value="([^"]+)"',
                login_page.text
            )
            atl_token = atl_token_match.group(1) if atl_token_match else ""

            form_data = {
                "os_username": self.username,
                "os_password": self.password,
                "os_destination": "/",
                "atl_token": atl_token,
                "user_role": "",
                "os_cookie": "true"
            }
            login_response = self.session.post(
                f"{self.base_url}/dologin.jsp",
                data=form_data,
                timeout=30,
                allow_redirects=True
            )

            if "invalid" in login_response.text.lower() or "incorrect" in login_response.text.lower():
                return {"success": False, "error": "Invalid username or password"}

            api_check = self.session.get(f"{self.base_url}/rest/api/2/myself", timeout=30)
            if api_check.status_code == 200:
                user_data = api_check.json()
                self.user_email = user_data.get("email", "")
                self.logged_in = True
                return {"success": True, "username": self.username}
            else:
                return {"success": False, "error": f"Verification failed (status: {api_check.status_code})"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def logout(self):
        if self.logged_in:
            try:
                self.session.delete(f"{self.base_url}/rest/auth/1/session")
            except:
                pass
            self.logged_in = False
            self.username = None
            self.user_email = None

    def fetch_issues(self, jql: str, start_at: int = 0, max_results: int = 100,
                     fields: str = None):
        """Fetch issues with pagination"""
        if fields is None:
            fields = ("summary,status,priority,issuetype,created,updated,resolutiondate,"
                    "creator,key,assignee,customfield_12001,customfield_11029,customfield_12031,"
                    "customfield_11043,customfield_11044,customfield_10100,customfield_10102,"
                    "customfield_10400,customfield_10401")

        all_issues = []
        url = f"{self.base_url}/rest/api/2/search"

        params = {
            "jql": jql,
            "startAt": start_at,
            "maxResults": max_results,
            "fields": fields
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

            except requests.exceptions.RequestException as e:
                raise Exception(f"Fetch error: {str(e)}")

        return all_issues

    def get_comments(self, issue_key: str):
        """Get all comments for an issue"""
        try:
            url = f"{self.base_url}/rest/api/2/issue/{issue_key}/comment"
            response = self.session.get(url, timeout=30)

            if response.status_code != 200:
                return []

            data = response.json()
            return data.get("comments", [])
        except Exception:
            return []

    def _parse_jira_datetime(self, created_str: str):
        """Parse Jira timestamp like 2026-05-08T17:32:01.123+0800 into datetime"""
        if not created_str:
            return None
        raw = str(created_str).strip()
        if not raw:
            return None

        # Normalize timezone suffix +0800 -> +08:00 for fromisoformat
        if re.search(r'[+-]\d{4}$', raw):
            raw = raw[:-5] + raw[-5:-2] + ":" + raw[-2:]

        try:
            return datetime.datetime.fromisoformat(raw)
        except ValueError:
            pass

        for fmt in ("%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z",
                   "%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S"):
            try:
                return datetime.datetime.strptime(raw, fmt)
            except ValueError:
                continue
        return None

    def _user_identity_values(self, user):
        """Return comparable Jira user identifiers from a REST user object"""
        if isinstance(user, list):
            values = set()
            for item in user:
                values.update(self._user_identity_values(item))
            return values

        if not isinstance(user, dict):
            return set()

        values = set()
        for key in ("name", "key", "emailAddress", "accountId", "displayName"):
            raw = str(user.get(key, "") or "").strip().lower()
            if raw:
                values.add(raw)
                if "@" in raw:
                    values.add(raw.split("@", 1)[0])
        return values

    def _current_user_identity_values(self):
        values = set()
        for raw in (self.username, self.user_email):
            text = str(raw or "").strip().lower()
            if text:
                values.add(text)
                if "@" in text:
                    values.add(text.split("@", 1)[0])
        return values

    def is_current_user(self, user) -> bool:
        return bool(self._current_user_identity_values() & self._user_identity_values(user))

    def _field_to_text(self, value):
        """Convert Jira field values to display text"""
        if value is None:
            return ""
        if isinstance(value, str):
            return value.strip()
        if isinstance(value, dict):
            child = value.get("child")
            child_text = self._field_to_text(child) if child is not None else ""
            val_text = str(value.get("value", "") or "").strip()
            key_text = str(value.get("key", "") or "").strip()
            if child_text and val_text:
                return f"{val_text} - {child_text}"
            return child_text or val_text or key_text
        if isinstance(value, list):
            parts = [self._field_to_text(item) for item in value]
            parts = [p for p in parts if p]
            return " - ".join(parts)
        return str(value).strip()

    def user_commented_in_date_range(self, issue_key: str, start_date, end_date) -> bool:
        """Check if current user commented within date range"""
        try:
            comments = self.get_comments(issue_key)
            for comment in comments:
                author = comment.get("author", {})
                if self.is_current_user(author):
                    created_str = comment.get("created", "")
                    if created_str:
                        comment_dt = self._parse_jira_datetime(created_str)
                        if not comment_dt:
                            continue
                        comment_date = comment_dt.date()
                        if start_date <= comment_date <= end_date:
                            return True
            return False
        except Exception:
            return False

    def get_user_latest_comment(self, issue_key: str, start_date, end_date,
                                timestamp_prefix: bool = False):
        """Get current user's latest comment in date range"""
        try:
            comments = self.get_comments(issue_key)
            latest_comment = None
            latest_dt = None

            for comment in comments:
                author = comment.get("author", {})
                if self.is_current_user(author):
                    created_str = comment.get("created", "")
                    if created_str:
                        comment_dt = self._parse_jira_datetime(created_str)
                        if not comment_dt:
                            continue
                        comment_date = comment_dt.date()
                        if start_date <= comment_date <= end_date:
                            body = comment.get("body", "")
                            text = self._clean_comment_for_display(body)
                            if text and (latest_dt is None or comment_dt > latest_dt):
                                latest_dt = comment_dt
                                if timestamp_prefix:
                                    latest_comment = f"[{comment_dt.strftime('%m-%d %H:%M')}] {text}"
                                else:
                                    latest_comment = text

            return latest_comment
        except Exception:
            return None

    def _clean_comment_for_display(self, body: str) -> str:
        """Light cleanup for comment display"""
        if not body:
            return ""

        text = body
        text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</(?:p|li|div|tr)>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'<[^>]+>', '', text)
        text = (text.replace('&nbsp;', ' ').replace('&amp;', '&')
                    .replace('&lt;', '<').replace('&gt;', '>')
                    .replace('&quot;', '"').replace('&#39;', "'"))
        text = re.sub(r'[\r\n]+', ' ', text)
        text = re.sub(r'\s+', ' ', text).strip()
        return text