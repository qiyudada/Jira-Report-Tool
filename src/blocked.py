"""
Blocked-status marker logic, shared between the GUI and the CLI/skill path.

A Jira issue is "blocked" when its status is one of BLOCKED_STATUSES and there is
no comment in the current report period. These issues get a static marker in the
"进展" column instead of an AI summary, so the reader can see at a glance that
the work is parked waiting on someone else (FAE / official release / workaround).

The marker text is status-specific via BLOCKED_STATUS_MARKERS, with
BLOCKED_STATUS_MARKER_DEFAULT as the fallback for unlisted blocked statuses.
The date in the marker is the date of the last comment in the context window
(or end_date if there are no comments at all), wrapped in <...>.
"""
import datetime


# Statuses that mean the issue is parked waiting on someone else.
BLOCKED_STATUSES = frozenset({
    "WAIT FAE INFO",
    "WORKED AROUND",
    "WAIT OFFICIAL RELEASE",
})

# Marker text per status. Lookup is case-insensitive against the status name.
# Unlisted blocked statuses fall back to BLOCKED_STATUS_MARKER_DEFAULT.
BLOCKED_STATUS_MARKERS = {
    "WAIT OFFICIAL RELEASE": "[问题已解决，等待SPM同步版本]",
}
BLOCKED_STATUS_MARKER_DEFAULT = "[当前阻塞中，等待信息确认]"


def compute_blocked_marker(status, comments, end_date):
    """Return a static "blocked, waiting" marker for blocked-status issues
    with no in-period activity. Returns None if the issue is not blocked or
    has activity in the period (caller should proceed with AI summary).

    Args:
        status: Jira status name (str). Case-insensitive.
        comments: iterable of dicts, each with at least {'date': date, 'in_period': bool}.
        end_date: datetime.date, used as fallback when no comments exist.

    Returns:
        str like "[问题已解决，等待SPM同步版本] <2026-05-28>", or None.
    """
    if not status or status.upper() not in BLOCKED_STATUSES:
        return None
    if any(c.get('in_period', True) for c in comments):
        return None

    last_comment_date = None
    for c in comments:
        d = c.get('date')
        if d is None:
            continue
        if last_comment_date is None or d > last_comment_date:
            last_comment_date = d

    marker_text = BLOCKED_STATUS_MARKERS.get(status.upper(), BLOCKED_STATUS_MARKER_DEFAULT)
    marker_date = last_comment_date or end_date
    if isinstance(marker_date, datetime.datetime):
        marker_date = marker_date.date()
    return f"{marker_text} <{marker_date.isoformat()}>"
