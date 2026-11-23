#!/usr/bin/env python3
"""
DA Dashboard - one script: Control Room -> local SQLite history -> interactive Excel dashboard.

Answers, per user:  who CREATED requests, who was ALLOCATED / ACCEPTED / PROCESSED them, who REJECTED them,
what is still OPEN or UNASSIGNED, and who the REVIEWERS are.  Source: Automation Co-Pilot (AARI) human tasks,
which is where Document Automation validation tasks live.

Usage
  python da_dashboard.py                         sync from Control Room, fetch task details, build the dashboard
  python da_dashboard.py --offline               rebuild from the local database only (no login)
  python da_dashboard.py --mock                  demo data, no Control Room needed (separate database)
  python da_dashboard.py --person alex.reviewer open the dashboard with that user already selected
  python da_dashboard.py --every 15              keep running and refresh the dashboard every 15 minutes
  python da_dashboard.py --no-details            skip the per-task detail calls (faster, less decision data)

Set CR_URL, CR_USERNAME and CR_PASSWORD at the top of this file (leave the placeholders to be prompted instead).
config.json is optional and no longer needed.
Requires: pip install requests pandas xlsxwriter
"""
import argparse
import getpass
import json
import os
import random
import re
import sqlite3
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

# ============================================================== CONNECTION SETTINGS - fill in here
CR_URL = "https://your-company.my.automationanywhere.digital"
CR_USERNAME = "YOUR_USERNAME"
CR_PASSWORD = "YOUR_PASSWORD"
DB_PATH = "da_tracker.db"

# ============================================================== constants
SYSTEM_USERS = {"", "unassigned", "riw.user.unassigned", "n/a", "na", "none", "null", "system", "nan",
                "unknown user", "(system / no user)"}
UNASSIGNED = "Unassigned"
SYSTEM_LABEL = "(system / no user)"
ALL = "(All)"

DEFAULT_CFG = {
    "control_room_url": CR_URL,
    "verify_ssl": True,
    "db_path": DB_PATH,
    "tasks_api": {"path": "/aari/v2/tasks/humans/list", "body": {}, "page_size": 100, "results_path": "list"},
}

TASK_COLUMNS = ["task_id", "status", "task_name", "assigned_to", "requested_by", "reference", "request_title",
                "task_created", "task_updated", "opened_on", "submitted_on"]

# Decision words, checked in this order so a rejection anywhere beats an approval elsewhere.
DECISION_ORDER = [
    ("Rejected", ("REJECT", "DENIED", "DECLINED", "NOT APPROVED")),
    ("Sent for reprocess", ("REPROCESS",)),
    ("Marked invalid", ("INVALID",)),
    ("Canceled", ("CANCEL",)),
    ("Validated / approved", ("VALIDATION_SUCCESS", "APPROVED", "APPROVE", "COMPLETE", "SUCCESS")),
]
NOT_A_DECISION = ("PENDING", "SEND", "FOR APPROVAL", "SUBMIT", "IN PROGRESS", "QUEUE", "ASSIGNED")
REJECT_DECISIONS = ("Rejected", "Marked invalid", "Sent for reprocess")
LEAF_SUFFIXES = (".string", ".number", ".boolean")

# Dashboard palette
NAVY, AMBER, GREEN, RED, GREY, TEAL = "#17365D", "#BF7B1A", "#217346", "#B94A3E", "#6B7280", "#2F6F8F"
PALE, LINE = "#F3F6FA", "#C9D3E0"


# ============================================================== small helpers
def utcnow():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


LOG_FILE = "da_dashboard.log"
LOCK_FILE = "da_dashboard.lock"
LOCK_STALE_SECONDS = 2 * 3600


def log(msg):
    """Print and append to the log file, so scheduled runs leave a trail."""
    line = f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}"
    print(msg, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def acquire_lock():
    """True if this run may proceed. A lock older than LOCK_STALE_SECONDS is treated as left over from a crash."""
    for _ in range(2):
        try:
            fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, f"{os.getpid()} {utcnow()}".encode())
            os.close(fd)
            return True
        except FileExistsError:
            try:
                if time.time() - os.path.getmtime(LOCK_FILE) > LOCK_STALE_SECONDS:
                    os.remove(LOCK_FILE)
                    continue
            except OSError:
                pass
            return False
    return False


def release_lock():
    try:
        os.remove(LOCK_FILE)
    except OSError:
        pass


def publish(tmp_path, final_path, retries=3, wait=10):
    """Swap the freshly built file in. If Excel has the target open, retry; if still locked, keep the temp
    file (it is published by the next run) and report False."""
    for attempt in range(retries):
        try:
            os.replace(tmp_path, final_path)
            return True
        except PermissionError:
            if attempt < retries - 1:
                time.sleep(wait)
    return False


def clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    return None if s == "" or s.lower() in ("nan", "null") else s


def get_path(obj, path):
    for part in (path or "").split("."):
        if not part:
            return None
        obj = obj.get(part) if isinstance(obj, dict) else None
    return obj


def norm_ts(value):
    """Any ISO-ish timestamp (also 2026-01-01T10:00:00+02:00[Europe/Paris]) -> 'YYYY-MM-DDTHH:MM:SSZ' in UTC."""
    if value in (None, ""):
        return None
    try:
        v = re.sub(r"\[[^\]]*\]\s*$", "", str(value))
        return pd.to_datetime(v, utc=True).strftime("%Y-%m-%dT%H:%M:%SZ")
    except Exception:
        return None


def norm_status(v):
    """HUMAN_PENDING -> Pending, HUMAN_SUBMITTED -> Submitted, HUMAN_CANCELED -> Canceled."""
    if not v:
        return None
    t = str(v).strip().upper()
    if t.startswith("HUMAN_"):
        t = t[6:]
    return t.replace("_", " ").title()


def person_name(v):
    """AARI returns people as objects; keep the username."""
    if isinstance(v, dict):
        v = v.get("username") or v.get("fullName") or v.get("name")
    return clean(v)


def is_human(v):
    return clean(v) is not None and str(v).strip().lower() not in SYSTEM_USERS


def flatten(obj, prefix=""):
    out = {}
    if isinstance(obj, dict):
        for k, v in obj.items():
            out.update(flatten(v, f"{prefix}{k}."))
    else:
        out[prefix[:-1]] = obj
    return out


def extract_outputs(outputs):
    """User-entered form values only (labels / structural metadata are dropped)."""
    flat = flatten(outputs) if isinstance(outputs, dict) else {}
    return {k: v for k, v in flat.items() if k.endswith(LEAF_SUFFIXES) and v not in (None, "", "None", [])}


def people_list(raw, umap):
    """approvers / declinedBy JSON -> list of usernames (accepts names, user objects or numeric ids)."""
    if not isinstance(raw, str) or not raw.strip():
        return []
    try:
        data = json.loads(raw)
    except Exception:
        return []
    items = data if isinstance(data, list) else [data]
    names = []
    for it in items:
        n = person_name(it) if isinstance(it, dict) else clean(it)
        if n and n.isdigit():
            n = umap.get(n, None)
        if n and is_human(n):
            names.append(n)
    return names


# ============================================================== Control Room client
class ControlRoom:
    def __init__(self, cfg):
        self.base = cfg["control_room_url"].rstrip("/")
        self.verify = cfg.get("verify_ssl", True)
        self.session = requests.Session()
        self.lock = threading.Lock()
        self.token = None
        self.generation = 0
        self._cred = None

    def login(self, seen_generation=None):
        with self.lock:
            if seen_generation is not None and seen_generation != self.generation:
                return  # another thread already renewed the token
            if self._cred is None:
                user = CR_USERNAME if CR_USERNAME != "YOUR_USERNAME" else \
                    (os.environ.get("DA_USER") or input("Control Room username: "))
                apikey = os.environ.get("DA_APIKEY")
                if apikey:
                    self._cred = {"username": user, "apiKey": apikey}
                else:
                    pw = CR_PASSWORD if CR_PASSWORD != "YOUR_PASSWORD" else \
                        (os.environ.get("DA_PASSWORD") or getpass.getpass("Password: "))
                    self._cred = {"username": user, "password": pw}
            r = self.session.post(self.base + "/v2/authentication", json=self._cred, timeout=60, verify=self.verify)
            if r.status_code in (400, 401, 403):
                self._cred = None
                sys.exit(f"Login failed (HTTP {r.status_code}). Check the username / password / API key.")
            r.raise_for_status()
            self.token = r.json()["token"]
            self.generation += 1

    def call(self, method, path, **kw):
        if self.token is None:
            self.login()
        last = None
        for attempt in range(5):
            gen = self.generation
            try:
                r = self.session.request(method, self.base + path, headers={"X-Authorization": self.token},
                                         timeout=120, verify=self.verify, **kw)
            except requests.RequestException as e:
                last = e
                time.sleep(2 ** attempt)
                continue
            if r.status_code == 401:
                self.login(gen)
                continue
            if r.status_code in (429, 500, 502, 503, 504):
                last = RuntimeError(f"HTTP {r.status_code}")
                time.sleep(2 ** attempt)
                continue
            return r
        raise RuntimeError(f"{method} {path} failed after retries: {last}")


# ============================================================== store
def connect(db_path):
    conn = sqlite3.connect(db_path, timeout=60)
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS tasks(
            task_key TEXT PRIMARY KEY,
            status TEXT, task_name TEXT, contributor TEXT, assigned_to TEXT, requested_by TEXT,
            reference TEXT, request_title TEXT, task_created TEXT, task_updated TEXT, loaded_at TEXT
        );
        CREATE TABLE IF NOT EXISTS task_snapshots(taken_at TEXT, status TEXT, cnt INTEGER);
        CREATE TABLE IF NOT EXISTS users(user_id TEXT PRIMARY KEY, username TEXT);
        CREATE TABLE IF NOT EXISTS task_status_history(
            task_key TEXT, status TEXT, changed_at TEXT, UNIQUE(task_key, status, changed_at));
        CREATE TABLE IF NOT EXISTS requests(
            request_id TEXT PRIMARY KEY, ref TEXT, process TEXT, title TEXT, status TEXT, created_by TEXT,
            user_group TEXT, created TEXT, updated TEXT, loaded_at TEXT);
        """
    )
    have = {r[1] for r in conn.execute("PRAGMA table_info(tasks)")}
    for col in ["task_id", "outcome", "outcome_info", "decision", "updated_by_id", "detail_loaded", "outputs_json",
                "approvers_json", "declined_by_json", "approval_buttons_json", "approval_count",
                "assignment_source", "assignment_rule_type", "step_state", "task_uuid", "activity_unique_id",
                "opened_on", "submitted_on"]:
        if col not in have:
            conn.execute(f"ALTER TABLE tasks ADD COLUMN {col} TEXT")
    conn.execute("CREATE INDEX IF NOT EXISTS ix_tasks_id ON tasks(task_id)")
    conn.commit()
    return conn


def parse_list_row(r):
    ref = clean(r.get("requestRef"))
    if ref is None and r.get("requestId") not in (None, ""):
        ref = str(r["requestId"])
    return {
        "task_id": clean(r.get("id")),
        "status": norm_status(r.get("stepSubState")),
        "task_name": clean(r.get("taskRef")),
        "assigned_to": person_name(r.get("assignee")),
        "requested_by": person_name(r.get("createdBy")),
        "reference": ref,
        "request_title": clean(r.get("requestTitle")),
        "task_created": norm_ts(r.get("createdOn")),
        "task_updated": norm_ts(r.get("updatedOn")),
        "opened_on": norm_ts(r.get("openedOn")),
        "submitted_on": norm_ts(r.get("submittedOn")),
    }


def upsert_tasks(conn, records):
    """Insert new tasks, update changed ones. A status change clears the detail flag so the new outcome is re-read."""
    now = utcnow()
    new = changed = skipped = 0
    for r in records:
        if not r["task_id"] or not r["reference"] or not r["task_created"]:
            skipped += 1
            continue
        row = conn.execute("SELECT task_key, status, assigned_to FROM tasks WHERE task_id=?",
                           (r["task_id"],)).fetchone()
        key = row[0] if row else f'{r["reference"]}|{r["task_name"]}|{r["task_created"]}'
        if row is None:
            row = conn.execute("SELECT task_key, status, assigned_to FROM tasks WHERE task_key=?", (key,)).fetchone()
        if row is None:
            new += 1
        elif row[1] != r["status"] or row[2] != r["assigned_to"]:
            changed += 1
        reset = row is not None and row[1] != r["status"]
        conn.execute(
            """INSERT INTO tasks(task_key,status,task_name,assigned_to,requested_by,reference,request_title,
                                 task_created,task_updated,loaded_at,task_id,opened_on,submitted_on)
               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
               ON CONFLICT(task_key) DO UPDATE SET status=excluded.status, assigned_to=excluded.assigned_to,
                 requested_by=COALESCE(excluded.requested_by, requested_by), task_updated=excluded.task_updated,
                 loaded_at=excluded.loaded_at, task_id=excluded.task_id, request_title=excluded.request_title,
                 opened_on=excluded.opened_on, submitted_on=excluded.submitted_on""",
            (key, r["status"], r["task_name"], r["assigned_to"], r["requested_by"], r["reference"],
             r["request_title"], r["task_created"], r["task_updated"], now, r["task_id"], r["opened_on"],
             r["submitted_on"]))
        if reset:
            conn.execute("UPDATE tasks SET detail_loaded=NULL WHERE task_key=?", (key,))
        if r["status"] and (row is None or row[1] != r["status"]):
            conn.execute("INSERT OR IGNORE INTO task_status_history VALUES(?,?,?)", (key, r["status"], now))
    conn.commit()
    for s, c in conn.execute("SELECT status, COUNT(*) FROM tasks GROUP BY status").fetchall():
        conn.execute("INSERT INTO task_snapshots VALUES(?,?,?)", (now, s, c))
    conn.commit()
    return new, changed, skipped


# ============================================================== collectors
def prune_missing(conn, table, id_col, live_ids, read, expected, label):
    """Mirror deletions: rows that Control Room no longer returns are removed from the database, so the dashboard
    always equals the live system. Only done after a COMPLETE read, and never for more than 20% of the rows
    (a bigger gap points to a partial or filtered response, not to real deletions)."""
    if not live_ids or (expected and read < int(expected)):
        return 0
    have = [r[0] for r in conn.execute(f"SELECT {id_col} FROM {table} WHERE {id_col} IS NOT NULL")]
    stale = [i for i in have if i not in live_ids]
    if not stale:
        return 0
    if len(stale) > 0.2 * max(len(have), 1):
        log(f"WARNING: {len(stale)} {label} in the database are not in Control Room's answer; more than 20% - "
            "not deleting them.")
        return 0
    conn.executemany(f"DELETE FROM {table} WHERE {id_col}=?", [(i,) for i in stale])
    conn.commit()
    log(f"   removed {len(stale)} {label} that no longer exist in Control Room")
    return len(stale)


def sync_task_list(cr, cfg, conn):
    src = {**DEFAULT_CFG["tasks_api"], **(cfg.get("tasks_api") or {})}
    size = int(src.get("page_size", 100))
    rows, seen, offset, expected, grand_total = [], set(), 0, None, None
    while True:
        body = json.loads(json.dumps(src.get("body", {})))
        body["page"] = {"offset": offset, "length": size}
        r = cr.call("POST", src["path"], json=body)
        if r.status_code in (401, 403):
            sys.exit(f"HTTP {r.status_code} from {src['path']}: this account cannot list Co-Pilot tasks.")
        r.raise_for_status()
        data = r.json()
        page = get_path(data, src.get("results_path", "list")) or []
        if isinstance(data, dict) and isinstance(data.get("page"), dict):
            expected = data["page"].get("totalFilter") or data["page"].get("total") or expected
            grand_total = data["page"].get("total") or grand_total
        fresh = [p for p in page if p.get("id") not in seen]
        seen.update(p.get("id") for p in fresh)
        rows += fresh
        print(f"\r   tasks read: {len(rows)}" + (f" / {expected}" if expected else ""), end="", flush=True)
        if len(page) < size or not fresh:
            break
        offset += size
    print()
    if expected and len(rows) < int(expected):
        log(f"WARNING: the API reports {expected} tasks but only {len(rows)} were read.")
    if grand_total and expected and int(grand_total) > int(expected):
        log(f"NOTE: Control Room holds {grand_total} tasks in total but its default filter returns {expected}. "
            "If the screen you compare with shows more, that filter is the reason.")
    recs = [parse_list_row(r) for r in rows]
    for r in rows:   # id -> username, learned for free from the list
        for who in ("assignee", "createdBy"):
            u = r.get(who)
            if isinstance(u, dict) and u.get("id") is not None and u.get("username"):
                conn.execute("INSERT OR IGNORE INTO users VALUES(?,?)", (str(u["id"]), str(u["username"])))
    new, changed, skipped = upsert_tasks(conn, recs)
    opened = sum(1 for r in recs if r["opened_on"])
    log(f"   {len(recs)} tasks: {new} new, {changed} changed" + (f", {skipped} skipped (no id/ref/date)" if skipped else "")
        + f"; opened-on timestamp present on {opened}")
    prune_missing(conn, "tasks", "task_id", {r["task_id"] for r in recs if r["task_id"]}, len(rows), expected, "tasks")
    in_db = conn.execute("SELECT COUNT(*) FROM tasks").fetchone()[0]
    if expected and in_db < int(expected):
        log(f"WARNING: reconcile check failed - API reports {expected} tasks, database holds {in_db}.")
    return {"api_reported": expected, "api_read": len(rows), "in_db": in_db, "skipped": skipped}


REQUESTS_PATH = "/aari/v3/requests/list"   # the list behind Automation Co-Pilot > Requests


def sync_requests(cr, cfg, conn):
    """Every request (Open / Success / Failed), including those that never created a human task.
    A failure here is logged and does not stop the task sync."""
    size = int(cfg["tasks_api"].get("page_size", 100))
    rows, seen, offset, expected = [], set(), 0, None
    try:
        while True:
            r = cr.call("POST", REQUESTS_PATH, json={"page": {"offset": offset, "length": size}})
            if r.status_code != 200:
                raise RuntimeError(f"HTTP {r.status_code} from {REQUESTS_PATH}: {r.text[:200]}")
            data = r.json()
            page = data.get("list") or []
            if isinstance(data.get("page"), dict):
                expected = data["page"].get("totalFilter") or data["page"].get("total") or expected
            fresh = [p for p in page if p.get("id") not in seen]
            seen.update(p.get("id") for p in fresh)
            rows += fresh
            print(f"\r   requests read: {len(rows)}" + (f" / {expected}" if expected else ""), end="", flush=True)
            if len(page) < size or not fresh:
                break
            offset += size
        print()
    except Exception as e:
        print()
        log(f"WARNING: requests could not be read ({e}). Request totals use the last stored copy.")
        return
    now = utcnow()
    for r in rows:
        grp = (r.get("createdByUserGroup") or {}).get("groupName") or r.get("teamname")
        conn.execute(
            "INSERT OR REPLACE INTO requests VALUES(?,?,?,?,?,?,?,?,?,?)",
            (str(r["id"]), clean(r.get("ref")), clean(r.get("process")), clean(r.get("title")),
             clean(r.get("status")), person_name(r.get("createdBy")), clean(grp),
             norm_ts(r.get("createdOn")), norm_ts(r.get("updatedOn")), now))
    conn.commit()
    prune_missing(conn, "requests", "request_id", {str(r["id"]) for r in rows}, len(rows), expected, "requests")
    in_db = conn.execute("SELECT COUNT(*) FROM requests").fetchone()[0]
    log(f"   {len(rows)} requests read; database holds {in_db}")
    if expected and len(rows) < int(expected):
        log(f"WARNING: the API reports {expected} requests but only {len(rows)} were read.")


def sync_users(cr, conn):
    rows, offset = [], 0
    try:
        while True:
            r = cr.call("POST", "/v2/usermanagement/users/list", json={"page": {"offset": offset, "length": 100}})
            if r.status_code != 200:
                raise RuntimeError(f"HTTP {r.status_code}")
            page = r.json().get("list") or []
            rows += page
            if len(page) < 100:
                break
            offset += 100
    except Exception as e:
        print(f"   user list unavailable ({e}); names fall back to those seen in the task list")
        return
    for u in rows:
        if u.get("id") is not None and u.get("username"):
            conn.execute("INSERT OR REPLACE INTO users VALUES(?,?)", (str(u["id"]), str(u["username"])))
    conn.commit()
    print(f"   users: {len(rows)}")


DETAIL_PATHS = {
    "outcome": "status", "outcome_info": "info", "decision": "outputs.Dropdown0.string", "updated_by_id": "updatedBy",
    "outputs": "outputs", "approvers": "approvers", "declined_by": "declinedBy",
    "approval_buttons": "approvalController.approvalButtons", "approval_count": "noofapproval",
    "assignment_source": "assignmentSource", "assignment_rule_type": "assignmentRuleType",
    "step_state": "stepState", "task_uuid": "taskUuid", "activity_unique_id": "activityUniqueId",
}


def store_detail(conn, key, d):
    g = lambda n: get_path(d, DETAIL_PATHS[n])
    blob = lambda v: json.dumps(v) if v else None
    text = lambda v: None if isinstance(v, (dict, list)) else clean(v)
    conn.execute(
        "UPDATE tasks SET outcome=?, outcome_info=?, decision=?, updated_by_id=?, outputs_json=?, approvers_json=?, "
        "declined_by_json=?, approval_buttons_json=?, approval_count=?, assignment_source=?, assignment_rule_type=?, "
        "step_state=?, task_uuid=?, activity_unique_id=?, detail_loaded=? WHERE task_key=?",
        (text(g("outcome")), text(g("outcome_info")), text(g("decision")), text(g("updated_by_id")),
         blob(extract_outputs(g("outputs"))), blob(g("approvers")), blob(g("declined_by")),
         blob(g("approval_buttons")), text(g("approval_count")), text(g("assignment_source")),
         text(g("assignment_rule_type")), text(g("step_state")), text(g("task_uuid")),
         text(g("activity_unique_id")), utcnow(), key))


def sync_details(cr, conn, workers, refresh=False, limit=None):
    """Detail of every Submitted task not yet read (decision, reason, who acted). Safe to Ctrl+C and re-run."""
    q = "SELECT task_key, task_id FROM tasks WHERE status='Submitted' AND task_id IS NOT NULL"
    if not refresh:
        q += " AND detail_loaded IS NULL"
    todo = conn.execute(q + " ORDER BY task_created DESC").fetchall()
    if limit:
        todo = todo[:limit]
    if not todo:
        print("   details: nothing to fetch")
        return
    print(f"   details: fetching {len(todo)} submitted tasks with {workers} workers (Ctrl+C is safe; re-run continues)")

    def one(item):
        key, tid = item
        r = cr.call("GET", f"/aari/v2/tasks/humans/{tid}")
        return key, r.status_code, (r.json() if r.status_code == 200 else None)

    done = failed = 0
    ex = ThreadPoolExecutor(max_workers=workers)
    try:
        futs = [ex.submit(one, it) for it in todo]
        for i, f in enumerate(as_completed(futs), 1):
            try:
                key, code, d = f.result()
            except Exception:
                failed += 1
                continue
            if d is not None:
                store_detail(conn, key, d)
                done += 1
            elif code == 404:   # task no longer exists: mark read so it is not retried forever
                conn.execute("UPDATE tasks SET detail_loaded=? WHERE task_key=?", (utcnow(), key))
            else:
                failed += 1
            if i % 50 == 0:
                conn.commit()
                print(f"\r   details: {i}/{len(todo)}", end="", flush=True)
    except KeyboardInterrupt:
        print("\n   interrupted - keeping what was fetched")
    finally:
        ex.shutdown(wait=False, cancel_futures=True)
        conn.commit()
    print(f"\n   details stored: {done}, failed: {failed}" + (" (re-run to retry)" if failed else ""))


# ============================================================== demo data
class MockControlRoom:
    """Generates Control-Room-shaped responses, so the whole pipeline is exercised without a server."""

    def __init__(self, n=420, seed=11):
        rnd = random.Random(seed)
        self.rnd = rnd
        creators = [("anna.creator", 237), ("ben.creator", 311), ("carla.creator", 267), ("dev.creator", 330)]
        reviewers = [("alex.reviewer", 240), ("bella.reviewer", 501), ("chris.reviewer", 502),
                     ("dana.reviewer", 503), ("emma.reviewer", 504)]
        now = datetime.now(timezone.utc)
        self.users = creators + reviewers
        self.list, self.detail = [], {}
        iso = lambda d: d.strftime("%Y-%m-%dT%H:%M:%S.000Z")
        person = lambda t: {"id": t[1], "username": t[0], "fullName": t[0]}
        for i in range(1, n + 1):
            c = rnd.choice(creators)
            created = now - timedelta(days=rnd.uniform(0, 240), hours=rnd.uniform(0, 12))
            kind = rnd.choices(["open", "unassigned", "ok", "reject", "invalid", "reprocess", "cancel"],
                               [22, 6, 45, 8, 6, 5, 4])[0]
            rv = rnd.choice(reviewers[:4] if rnd.random() < 0.9 else reviewers)
            row = {"id": 5000 + i, "stepState": "IN_PROGRESS", "stepSubState": "HUMAN_PENDING",
                   "createdOn": iso(created), "updatedOn": iso(created), "requestId": 900 + i // 2,
                   "requestRef": f"{100 + i // 2}-{1 + i % 3}", "taskRef": f"{100 + i // 2}-{1 + i % 3}-{i % 4}",
                   "requestTitle": rnd.choice(["Validation - invoice.pdf", "Validation - claim.pdf", "Validation - PO.pdf"]),
                   "assignee": person(rv), "createdBy": person(c), "openedOn": "", "submittedOn": ""}
            if kind == "unassigned":
                row["assignee"] = {"id": -1, "username": "Unassigned"}
            if kind == "open" and rnd.random() < 0.45:
                row["openedOn"] = iso(created + timedelta(hours=rnd.uniform(1, 30)))
            if kind not in ("open", "unassigned"):
                opened = created + timedelta(hours=rnd.uniform(0.5, 30))
                done = opened + timedelta(minutes=rnd.uniform(3, 40))
                row.update({"stepState": "COMPLETED", "stepSubState": "HUMAN_SUBMITTED", "updatedOn": iso(done)})
                if rnd.random() < 0.55:
                    row["openedOn"], row["submittedOn"] = iso(opened), iso(done)
                if kind == "cancel":
                    row["stepSubState"] = "HUMAN_CANCELED"
                status = {"ok": "VALIDATION_SUCCESS", "reject": "Reject", "invalid": "VALIDATION_INVALID",
                          "reprocess": "VALIDATION_REPROCESS", "cancel": "Cancel"}[kind]
                info = "VALIDATION_INVALID - " + rnd.choice(["Wrong document type", "I can't read the document",
                                                              "Other"]) if kind == "invalid" else status
                self.detail[5000 + i] = {"id": 5000 + i, "status": status, "info": info, "stepState": "COMPLETED",
                                         "updatedBy": rv[1] if rnd.random() < 0.5 else 0, "outputs": {},
                                         "approvers": [], "declinedBy": [{"id": rv[1], "username": rv[0]}]
                                         if kind == "reject" and rnd.random() < 0.4 else []}
            self.list.append(row)
        # requests: one per task reference, plus requests that failed before any human task existed
        by_ref = {}
        for t in self.list:
            by_ref.setdefault(t["requestRef"], t)
        self.requests = []
        for k, (ref, t) in enumerate(by_ref.items()):
            done = t["stepSubState"] != "HUMAN_PENDING"
            self.requests.append({"id": 100000 + k, "status": "SUCCESS" if done else "OPENED", "process": "DA_Process",
                                  "title": t["requestTitle"], "createdOn": t["createdOn"] + "[Etc/UTC]",
                                  "updatedOn": t["updatedOn"] + "[Etc/UTC]", "createdBy": t["createdBy"],
                                  "ref": ref, "createdByUserGroup": {"groupName": rnd.choice(["", "DemoTeam"])}})
        for k in range(70):
            c = rnd.choice(creators)
            when = now - timedelta(days=rnd.uniform(0, 240))
            self.requests.append({"id": 200000 + k, "status": "ERROR", "process": "DA_Process", "title": "failed run",
                                  "createdOn": iso(when) + "[Etc/UTC]", "updatedOn": iso(when) + "[Etc/UTC]",
                                  "createdBy": person(c), "ref": f"900-{k}", "createdByUserGroup": {"groupName": ""}})

    def users_payload(self):
        return [{"id": i, "username": u} for u, i in self.users]


class MockClient:
    def __init__(self):
        self.m = MockControlRoom()

    def call(self, method, path, **kw):
        class R:
            def __init__(s, code, data):
                s.status_code, s._d = code, data

            def json(s):
                return s._d

            def raise_for_status(s):
                pass
        if path.endswith("/requests/list"):
            p = kw["json"]["page"]
            lst = self.m.requests[p["offset"]:p["offset"] + p["length"]]
            return R(200, {"page": {"total": len(self.m.requests), "totalFilter": len(self.m.requests)}, "list": lst})
        if path.endswith("/tasks/humans/list"):
            p = kw["json"]["page"]
            lst = self.m.list[p["offset"]:p["offset"] + p["length"]]
            return R(200, {"page": {"total": len(self.m.list), "totalFilter": len(self.m.list)}, "list": lst})
        if "/tasks/humans/" in path:
            d = self.m.detail.get(int(path.rsplit("/", 1)[1]))
            return R(200 if d else 404, d)
        if "usermanagement" in path:
            p = kw["json"]["page"]
            return R(200, {"list": self.m.users_payload()[p["offset"]:p["offset"] + p["length"]]})
        return R(404, None)


# ============================================================== classification
def classify(outcome, outcome_info, decision, outputs_json):
    """(decision category, field that decided it). Trusted fields first; rejection before approval."""
    def match(text):
        for cat, keys in DECISION_ORDER:
            if any(k in text for k in keys):
                return cat
        return None

    def norm(v):
        t = str(v if v is not None else "").strip().upper()
        return "" if t in ("NONE", "NULL", "NAN", "TRUE", "FALSE") else t

    for name, val in (("outcome", outcome), ("outcome info", outcome_info), ("review form", decision)):
        t = norm(val)
        if t and not any(x in t for x in NOT_A_DECISION):
            cat = match(t)
            if cat:
                return cat, name
    if isinstance(outputs_json, str) and outputs_json.strip() not in ("", "{}"):
        try:
            data = json.loads(outputs_json)
        except Exception:
            data = {}
        for k, v in (data.items() if isinstance(data, dict) else []):
            t = norm(v)
            if not t or len(t) > 40 or t.count(" ") > 3 or any(x in t for x in NOT_A_DECISION):
                continue
            cat = match(t)
            if cat:
                return cat, f"form field '{k.rsplit('.', 1)[0]}'"
    return None, None


def group_of(status, decision):
    s = (status or "").lower()
    if "cancel" in s or "expire" in s or decision == "Canceled":
        return "Canceled"
    if s == "submitted":
        return "Reject" if decision in REJECT_DECISIONS else "Success"
    return "Open"


# ============================================================== data preparation
def prepare(conn, tz=None):
    raw = pd.read_sql("SELECT * FROM tasks", conn)
    if raw.empty:
        sys.exit("The database has no tasks yet. Run without --offline to sync from Control Room (or use --mock).")
    umap = dict(conn.execute("SELECT user_id, username FROM users").fetchall())

    def dt(col):
        s = pd.to_datetime(raw[col], utc=True, errors="coerce")
        if tz:
            s = s.dt.tz_convert(tz)
        return s.dt.tz_localize(None)

    df = pd.DataFrame({"task_id": raw["task_id"], "reference": raw["reference"], "title": raw["request_title"],
                       "task_ref": raw["task_name"], "raw_status": raw["status"]})
    df["created_by"] = [n if is_human(n) else SYSTEM_LABEL for n in raw["requested_by"]]
    df["person"] = [n if is_human(n) else UNASSIGNED for n in raw["assigned_to"]]
    df["created"], df["updated"] = dt("task_created"), dt("task_updated")
    df["opened"], df["submitted"] = dt("opened_on"), dt("submitted_on")

    dec = [classify(*r) for r in raw[["outcome", "outcome_info", "decision", "outputs_json"]].itertuples(index=False)]
    decision = [d[0] for d in dec]
    declined = [people_list(x, umap) for x in raw["declined_by_json"]]
    approvers = [people_list(x, umap) for x in raw["approvers_json"]]
    decision = ["Rejected" if declined[i] and decision[i] is None else decision[i] for i in range(len(df))]
    df["decision"] = decision
    df["decision_field"] = [d[1] for d in dec]
    df["group"] = [group_of(s, d) for s, d in zip(raw["status"], decision)]

    recorded = [umap.get(str(x).strip()) if clean(x) else None for x in raw["updated_by_id"]]
    acted = []
    for i, g in enumerate(df["group"]):
        if g == "Open":
            acted.append(None)
        elif declined[i]:
            acted.append(declined[i][0])
        elif is_human(recorded[i]):
            acted.append(recorded[i])
        elif is_human(raw["assigned_to"].iat[i]):
            acted.append(raw["assigned_to"].iat[i])
        else:
            acted.append(None)
    df["acted_by"] = [a if a else "" for a in acted]
    df["acted_known"] = [is_human(r) for r in recorded]
    df["reviewers"] = [", ".join(a) for a in approvers]

    def reason(info, dec_):
        t = clean(info) or ""
        return t.split(" - ", 1)[1] if " - " in t else (t if dec_ in REJECT_DECISIONS and t.upper() not in
                                                         ("VALIDATION_REPROCESS", "REJECT") else "")
    df["reason"] = [reason(i, d) for i, d in zip(raw["outcome_info"], decision)]

    # timing quality: Open -> Started / Not started; closed -> Valid / Invalid / Missing
    timing, dur = [], []
    for g, o, s in zip(df["group"], df["opened"], df["submitted"]):
        if g == "Open":
            timing.append("Started" if pd.notna(o) else "Not started")
            dur.append(None)
        elif pd.isna(o) or pd.isna(s):
            timing.append("Missing")
            dur.append(None)
        elif s < o:
            timing.append("Invalid")
            dur.append(None)
        else:
            timing.append("Valid")
            dur.append((s - o).total_seconds() / 86400)
    df["timing"], df["duration"] = timing, dur
    first = df.groupby("reference")["created"].transform("min")
    ids = df.assign(_f=(df["created"] == first)).groupby("reference")["_f"].transform(lambda s: s.cumsum() == 1)
    df["first"] = (ids & (df["created"] == first)).astype(int)
    for c in ("title", "task_ref", "decision", "decision_field", "reason", "reviewers", "reference", "raw_status"):
        df[c] = df[c].astype(object).where(df[c].notna(), "").astype(str)
    df = df.sort_values(["created", "task_id"]).reset_index(drop=True)
    meta = {"details": int(raw["detail_loaded"].notna().sum()), "submitted": int((raw["status"] == "Submitted").sum()),
            "snapshots": pd.read_sql("SELECT * FROM task_snapshots", conn),
            "unknown_status": sorted({s for s in raw["status"].dropna().unique()} - {"Pending", "Submitted", "Canceled"})}
    return df, meta


REQUEST_STATUS = {"SUCCESS": "Success", "OPENED": "Open", "ERROR": "Failed"}


def prepare_requests(conn, tz=None):
    """Requests (the Co-Pilot > Requests screen) as a DataFrame with a friendly status."""
    raw = pd.read_sql("SELECT * FROM requests", conn)
    cols = ["request_id", "ref", "process", "title", "status", "created_by", "user_group", "created", "updated"]
    if raw.empty:
        return pd.DataFrame(columns=cols)
    rq = pd.DataFrame({"request_id": raw["request_id"], "ref": raw["ref"].fillna(""), "process": raw["process"].fillna(""),
                       "title": raw["title"].fillna(""),
                       "status": [REQUEST_STATUS.get(str(s).upper(), str(s).title()) for s in raw["status"]],
                       "created_by": [n if is_human(n) else SYSTEM_LABEL for n in raw["created_by"]],
                       "user_group": raw["user_group"].fillna("")})
    for c in ("created", "updated"):
        s = pd.to_datetime(raw[c], utc=True, errors="coerce")
        rq[c] = (s.dt.tz_convert(tz) if tz else s).dt.tz_localize(None)
    return rq.sort_values("created", ascending=False).reset_index(drop=True)


# ============================================================== formula + cached value engine
SEL = object()    # the allocated person chosen in the dashboard dropdown
SELC = object()   # the "created by" user chosen in the dashboard dropdown


class Calc:
    """Builds a COUNTIFS formula and, at the same time, the value it has with the initial filters
    (so the workbook shows correct numbers even in viewers that do not recalculate)."""

    def __init__(self, df, person=None, dfrom=None, dto=None, rq=None, creator=None):
        self.df, self.person, self.dfrom, self.dto = df, (person if person not in (None, ALL) else None), dfrom, dto
        self.rq = rq if rq is not None else pd.DataFrame(columns=["status", "created_by", "created"])
        self.creator = creator if creator not in (None, ALL) else None

    def rcount(self, status=None, creator=None, dated=True):
        """COUNTIFS over the Requests sheet. creator: SELC (dropdown), (cell, name) or None."""
        d, m, p = self.rq, pd.Series(True, index=self.rq.index), []
        if creator is SELC:
            p.append("r_creator,f_creator")
            if self.creator:
                m &= d["created_by"] == self.creator
        elif creator is not None:
            p.append(f"r_creator,{creator[0]}")
            m &= d["created_by"] == creator[1]
        if status:
            p.append(f'r_status,"{status}"')
            m &= d["status"] == status
        if dated:
            p.append('r_created,">="&f_from,r_created,"<"&f_to')
            if self.dfrom is not None:
                m &= d["created"] >= self.dfrom
            if self.dto is not None:
                m &= d["created"] < self.dto + timedelta(days=1)
        return "=COUNTIFS(" + ",".join(p or ['r_id,"<>"']) + ")", int(m.sum())

    def _mask(self, person, group, decision, timing, creator, acted, first, month, dated):
        d, m = self.df, pd.Series(True, index=self.df.index)
        if person is SEL:
            if self.person:
                m &= d["person"] == self.person
        elif person is not None:
            m &= d["person"] == person[1]
        if group:
            m &= d["group"] == group
        if decision:
            m &= d["decision"] == decision
        if timing:
            m &= d["timing"] == timing
        if creator is not None:
            m &= d["created_by"] == creator[1]
        if acted is not None:
            m &= d["acted_by"] == acted[1]
        if first:
            m &= d["first"] == 1
        if month is not None:
            m &= (d["created"] >= month[1]) & (d["created"] < month[1] + pd.offsets.MonthBegin(1))
        if dated:
            if self.dfrom is not None:
                m &= d["created"] >= self.dfrom
            if self.dto is not None:
                m &= d["created"] < self.dto + timedelta(days=1)
        return m

    def parts(self, person=None, group=None, decision=None, timing=None, creator=None, acted=None, first=False,
              month=None, dated=True):
        p = []
        if person is SEL:
            p.append("d_person,f_person")
        elif person is not None:
            p.append(f"d_person,{person[0]}")
        if group:
            p.append(f'd_group,"{group}"')
        if decision:
            p.append(f'd_decision,"{decision}"')
        if timing:
            p.append(f'd_timing,"{timing}"')
        if creator is not None:
            p.append(f"d_creator,{creator[0]}")
        if acted is not None:
            p.append(f"d_acted,{acted[0]}")
        if first:
            p.append("d_first,1")
        if month is not None:
            p.append(f'd_created,">="&{month[0]},d_created,"<"&EDATE({month[0]},1)')
        if dated:
            p.append('d_created,">="&f_from,d_created,"<"&f_to')
        return p

    def count(self, **kw):
        m = self._mask(*(kw.get(k) for k in ("person", "group", "decision", "timing", "creator", "acted")),
                       kw.get("first", False), kw.get("month"), kw.get("dated", True))
        return "=COUNTIFS(" + ",".join(self.parts(**kw) or ['d_id,"<>"']) + ")", int(m.sum())

    def avg_duration(self, **kw):
        m = self._mask(kw.get("person"), None, None, "Valid", None, None, False, None, True)
        v = self.df.loc[m, "duration"].dropna()
        f = "=IFERROR(AVERAGEIFS(d_dur," + ",".join(self.parts(timing="Valid", **kw)) + "),0)"
        return f, float(v.mean()) if len(v) else 0.0


# ============================================================== workbook
def build_workbook(df, rq, meta, out, init):
    n = len(df)
    last = n + 1
    nr = len(rq)
    rlast = max(nr, 1) + 1
    # Either dropdown selects a user: the one left blank falls back to the other, so picking a name in
    # "Created by" OR "Allocated person" shows that user's whole activity (requests and tasks).
    ip, ic = init["person"], init.get("creator")
    ip = None if ip in (None, ALL) else ip
    ic = None if ic in (None, ALL) else ic
    calc = Calc(df, ip or ic, None, None, rq, ic or ip)
    wb = xlsxwriter.Workbook(out, {"default_date_format": "dd-mmm-yyyy hh:mm", "nan_inf_to_errors": True})
    F = lambda **k: wb.add_format({"font_name": "Calibri", "font_size": 10, "valign": "vcenter", **k})

    f_title = F(bold=True, font_size=20, font_color="#FFFFFF", bg_color=NAVY, indent=1)
    f_hint = F(italic=True, font_color="#333333")
    f_sec = F(bold=True, font_size=12, font_color=NAVY)
    f_th = F(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="center", text_wrap=True, border=1, border_color=NAVY)
    f_thl = F(bold=True, font_color="#FFFFFF", bg_color=NAVY, text_wrap=True, border=1, border_color=NAVY)
    f_txt = F(bottom=1, bottom_color="#E3E8EF")
    f_num = F(bottom=1, bottom_color="#E3E8EF", num_format="#,##0;-#,##0;0", align="right")
    f_tot = F(bold=True, top=1, bottom=1, bg_color=PALE, align="right", num_format="#,##0")
    f_totl = F(bold=True, top=1, bottom=1, bg_color=PALE)
    f_ital = F(italic=True, font_color=GREY, bottom=1, bottom_color="#E3E8EF")
    f_inum = F(italic=True, font_color=GREY, bottom=1, bottom_color="#E3E8EF", num_format="#,##0", align="right")
    f_lbl = F(bottom=1, bottom_color="#E3E8EF")
    f_val = F(bold=True, align="right", bottom=1, bottom_color="#E3E8EF", num_format="#,##0")
    f_dur = F(bold=True, align="right", bottom=1, bottom_color="#E3E8EF", num_format="[h]:mm:ss")
    f_in = F(bg_color="#FFF7D6", border=1, border_color="#D4B84A", bold=True)
    f_ind = F(bg_color="#FFF7D6", border=1, border_color="#D4B84A", bold=True, num_format="dd-mmm-yyyy", align="left")
    f_flab = F(bold=True, font_color=NAVY)
    f_note = F(italic=True, font_color=GREY, text_wrap=True, valign="top", font_size=9)
    tile_head = lambda c: F(bold=True, font_color="#FFFFFF", bg_color=c, indent=1, font_size=10)
    tile_val = F(bold=True, font_size=26, align="right", bg_color="#F3F6FA", border=1, border_color="#9FB0C6",
                 num_format="#,##0", indent=1)
    f_dh = F(bold=True, font_color="#FFFFFF", bg_color=NAVY, text_wrap=True, align="center", valign="vcenter")
    f_dd = F(num_format="dd-mmm-yyyy hh:mm")
    f_dn = F(num_format="[h]:mm:ss")
    f_dg = F()

    ws = wb.add_worksheet("Dashboard")
    wu = wb.add_worksheet("User Activity")
    wp = wb.add_worksheet("Pending Queue")
    wr = wb.add_worksheet("Rejections")
    wq = wb.add_worksheet("Requests")
    wd = wb.add_worksheet("Data")
    wl = wb.add_worksheet("Lists")
    wl.hide()

    # ---------------------------------------------------------- Data sheet
    cols = [("Task ID", "task_id", "id", 10, f_dg), ("Request Ref", "reference", "ref", 12, f_dg),
            ("Request Title", "title", "title", 34, f_dg), ("Task Ref", "task_ref", "tref", 14, f_dg),
            ("Created By", "created_by", "creator", 22, f_dg), ("Allocated Person", "person", "person", 22, f_dg),
            ("Acted By", "acted_by", "acted", 22, f_dg), ("Status Group", "group", "group", 13, f_dg),
            ("Decision Detail", "decision", "decision", 20, f_dg), ("Reason", "reason", "reason", 30, f_dg),
            ("Reviewer(s)", "reviewers", "reviewers", 20, f_dg),
            ("Task Created", "created", "created", 17, f_dd), ("Task Updated", "updated", "updated", 17, f_dd),
            ("Opened On", "opened", "opened", 17, f_dd), ("Submitted On", "submitted", "submitted", 17, f_dd),
            ("Duration", "duration", "dur", 11, f_dn), ("Timing Check", "timing", "timing", 13, f_dg),
            ("First Task Of Request", "first", "first", 11, f_dg), ("Raw Status", "raw_status", "rawstatus", 12, f_dg),
            ("Decided By Field", "decision_field", "dfield", 22, f_dg)]
    colref = {}
    for j, (head, col, name, width, fm) in enumerate(cols):
        wd.write(0, j, head, f_dh)
        wd.set_column(j, j, width, fm)
        L = xl_col_to_name(j)
        wb.define_name(f"d_{name}", f"=Data!${L}$2:${L}${last}")
        colref[name] = j
    wd.set_row(0, 32)
    for i, rec in enumerate(df[[c[1] for c in cols]].itertuples(index=False), start=1):
        for j, v in enumerate(rec):
            if v is None or (not isinstance(v, str) and pd.isna(v)):
                continue
            if isinstance(v, pd.Timestamp):
                wd.write_datetime(i, j, v.to_pydatetime())
            elif isinstance(v, (int, float)):
                wd.write_number(i, j, float(v))
            else:
                wd.write_string(i, j, str(v))
    wd.freeze_panes(1, 0)
    wd.autofilter(0, 0, max(n, 1), len(cols) - 1)

    # ---------------------------------------------------------- Requests sheet (the Co-Pilot > Requests list)
    rcols = [("Request ID", "request_id", "id", 12), ("Reference", "ref", "ref", 12), ("Process", "process", "process", 28),
             ("Request Title", "title", "title", 34), ("Status", "status", "status", 11),
             ("Created By", "created_by", "creator", 22), ("User Group", "user_group", "group", 18),
             ("Created", "created", "created", 17), ("Updated", "updated", "updated", 17)]
    for j, (head, col, name, width) in enumerate(rcols):
        wq.write(0, j, head, f_dh)
        wq.set_column(j, j, width, f_dd if col in ("created", "updated") else f_dg)
        L = xl_col_to_name(j)
        wb.define_name(f"r_{name}", f"=Requests!${L}$2:${L}${rlast}")
    wq.set_row(0, 28)
    for i, rec in enumerate(rq[[c[1] for c in rcols]].itertuples(index=False), start=1):
        for j, v in enumerate(rec):
            if v is None or (not isinstance(v, str) and pd.isna(v)):
                continue
            if isinstance(v, pd.Timestamp):
                wq.write_datetime(i, j, v.to_pydatetime())
            else:
                wq.write_string(i, j, str(v))
    wq.freeze_panes(1, 0)
    wq.autofilter(0, 0, max(nr, 1), len(rcols) - 1)

    # ---------------------------------------------------------- Lists sheet (filters, dropdown, chart feeders)
    by_vol = df["person"].value_counts()
    persons = list(by_vol.index)
    everyone = persons + [c for c in (rq["created_by"].value_counts().index if nr else []) if c not in persons]
    wl.write(0, 0, "Users")           # row 2 is left empty on purpose: it is the blank choice in the dropdown
    wl.write(2, 0, ALL)
    for i, p in enumerate(everyone, start=3):
        wl.write(i, 0, p)
    wl.write(0, 1, "Effective filters")
    P, C = "Dashboard!$K$10", "Dashboard!$K$11"
    wl.write_formula(1, 1, f'=IF(OR({P}="",{P}="(All)"),IF(OR({C}="",{C}="(All)"),"*",{C}),{P})', None,
                     (ip or ic) or "*")
    wl.write_number(2, 1, 0)          # no date filter: everything is always included
    wl.write_number(3, 1, 9999999)
    wl.write_formula(4, 1, f'=IF(OR({C}="",{C}="(All)"),IF(OR({P}="",{P}="(All)"),"*",{P}),{C})', None,
                     (ic or ip) or "*")
    wb.define_name("f_person", "=Lists!$B$2")
    wb.define_name("f_from", "=Lists!$B$3")
    wb.define_name("f_to", "=Lists!$B$4")
    wb.define_name("f_creator", "=Lists!$B$5")

    # ---------------------------------------------------------- Dashboard
    ws.hide_gridlines(2)
    ws.set_column(0, 11, 19)
    ws.set_zoom(85)
    ws.set_row(0, 6)
    ws.set_row(1, 36)
    ws.merge_range(1, 0, 1, 11, "DOCUMENT AUTOMATION | REQUEST OVERVIEW", f_title)
    ws.merge_range(2, 0, 2, 11,
                   "Pick a user in FILTERS (yellow cells) to filter the tiles, timing, charts and the full activity "
                   "lists below. Choose (All) to show everyone.", f_hint)

    def put(r, c, fv, fmt, c2=None):
        if c2 is not None and c2 > c:
            ws.merge_range(r, c, r, c2, "", fmt)
        ws.write_formula(r, c, fv[0], fmt, fv[1])

    # first four tiles count REQUESTS (follow "Created by"); last two count human TASKS (follow "Allocated person")
    tiles = [("TOTAL REQUESTS", NAVY, "req", {}), ("OPEN REQUESTS", AMBER, "req", dict(status="Open")),
             ("SUCCESS REQUESTS", GREEN, "req", dict(status="Success")),
             ("FAILED REQUESTS", RED, "req", dict(status="Failed")),
             ("TASKS REJECTED", "#8E3B33", "task", dict(group="Reject")),
             ("TASKS UNASSIGNED", TEAL, "task", dict(group="Open"))]
    tile_cells = {}
    for k, (label, colr, scope, kw) in enumerate(tiles):
        c = k * 2
        ws.merge_range(3, c, 3, c + 1, label, tile_head(colr))
        ws.merge_range(4, c, 5, c + 1, "", tile_val)
        if scope == "req":
            fv = calc.rcount(creator=SELC, **kw)
        elif label.startswith("TASKS UNASSIGNED"):
            fv = (f'=IF(f_person="*",COUNTIFS(d_person,"{UNASSIGNED}",d_group,"Open",d_created,">="&f_from,'
                  f'd_created,"<"&f_to),COUNTIFS(d_person,f_person,d_person,"{UNASSIGNED}",d_group,"Open",'
                  f'd_created,">="&f_from,d_created,"<"&f_to))',
                  int(((df["group"] == "Open") & (df["person"] == UNASSIGNED) &
                       ((df["person"] == calc.person) if calc.person else True) &
                       ((df["created"] >= calc.dfrom) if calc.dfrom is not None else True) &
                       ((df["created"] < calc.dto + timedelta(days=1)) if calc.dto is not None else True)).sum()))
        else:
            fv = calc.count(person=SEL, **kw)
        ws.write_formula(4, c, fv[0], tile_val, fv[1])
        tile_cells[label] = xl_rowcol_to_cell(4, c, row_abs=True, col_abs=True)
    ws.set_row(4, 22)
    ws.set_row(5, 22)

    # --- allocation table (left)
    ws.write(7, 0, "REQUEST ALLOCATION | BY ALLOCATED PERSON", f_sec)
    hdr = [("Allocated person", None), ("OPEN", AMBER), ("SUCCESS", GREEN), ("REJECT", RED), ("CANCELED", GREY),
           ("Total", None)]
    ws.merge_range(8, 0, 8, 1, "Allocated person", f_thl)
    for j, (h, _) in enumerate(hdr[1:], start=2):
        ws.write(8, j, h, f_th)
    shown = persons[:20]
    r0 = 9
    asum = {j: 0 for j in range(2, 7)}
    for i, p in enumerate(shown):
        r = r0 + i
        ws.merge_range(r, 0, r, 1, p, f_txt)
        ref = (f"$A{r + 1}", p)
        for j, g in zip(range(2, 6), ["Open", "Success", "Reject", "Canceled"]):
            fv = calc.count(person=ref, group=g)
            ws.write_formula(r, j, fv[0], f_num, fv[1])
            asum[j] += fv[1]
        rowtot = sum(calc.count(person=ref, group=g)[1] for g in ["Open", "Success", "Reject", "Canceled"])
        ws.write_formula(r, 6, f"=SUM(C{r + 1}:F{r + 1})", f_num, rowtot)
        asum[6] += rowtot
    end = r0 + len(shown)
    if len(persons) > len(shown):
        ws.merge_range(end, 0, end, 1, f"All others ({len(persons) - len(shown)} people)", f_ital)
        oth_row = 0
        for j, g in zip(range(2, 6), ["Open", "Success", "Reject", "Canceled"]):
            L = xl_col_to_name(j)
            tot = calc.count(group=g)
            oth = tot[1] - asum[j]
            ws.write_formula(end, j, f"={tot[0][1:]}-SUM({L}{r0 + 1}:{L}{end})", f_inum, oth)
            asum[j] += oth
            oth_row += oth
        ws.write_formula(end, 6, f"=SUM(C{end + 1}:F{end + 1})", f_inum, oth_row)
        asum[6] += oth_row
        end += 1
    ws.merge_range(end, 0, end, 1, "Total", f_totl)
    for j in range(2, 7):
        L = xl_col_to_name(j)
        ws.write_formula(end, j, f"=SUM({L}{r0 + 1}:{L}{end})", f_tot, asum[j])
    left_end = end
    for j, colr in zip(range(2, 6), [AMBER, GREEN, RED, GREY]):
        ws.conditional_format(r0, j, r0 + len(shown) - 1, j,
                              {"type": "data_bar", "bar_color": colr, "bar_solid": False, "min_type": "num",
                               "min_value": 0})
    ws.conditional_format(r0, 0, r0 + len(shown) - 1, 6, {"type": "formula", "criteria": "=$A10=Lists!$B$2",
                                                          "format": wb.add_format({"bold": True, "bg_color": "#DCE9F7"})})

    # --- filter panel (right)
    ws.merge_range(7, 8, 7, 11, "FILTERS", f_sec)
    ws.merge_range(8, 8, 8, 11, "Change the yellow cells", f_th)
    ws.merge_range(9, 8, 9, 9, "Allocated person", f_flab)
    ws.merge_range(9, 10, 9, 11, init["person"] or ALL, f_in)
    ws.merge_range(10, 8, 10, 9, "Created by", f_flab)
    ws.merge_range(10, 10, 10, 11, init.get("creator") or ALL, f_in)
    for r in (9, 10):
        ws.data_validation(r, 10, r, 10, {"validate": "list", "source": f"=Lists!$A$2:$A${len(everyone) + 3}",
                                          "error_title": "Unknown user", "error_message": "Pick a name from the list."})

    # --- requests created by user (right)
    ws.write(14, 8, "REQUESTS CREATED | BY USER", f_sec)
    ws.merge_range(15, 8, 15, 9, "Created by", f_thl)
    ws.write(15, 10, "Requests", f_th)
    ws.write(15, 11, "Tasks", f_th)
    ctab = rq["created_by"].value_counts() if nr else df["created_by"].value_counts()
    cshow = [u for u in ctab.index if is_human(u)][:13]
    rr0 = 16
    csum = {10: 0, 11: 0}
    for i, u in enumerate(cshow):
        r = rr0 + i
        ws.merge_range(r, 8, r, 9, u, f_txt if u != SYSTEM_LABEL else f_ital)
        ref = (f"$I{r + 1}", u)
        fv = calc.rcount(creator=ref)
        ws.write_formula(r, 10, fv[0], f_num, fv[1])
        csum[10] += fv[1]
        fv = calc.count(creator=ref)
        ws.write_formula(r, 11, fv[0], f_num, fv[1])
        csum[11] += fv[1]
    rend = rr0 + len(cshow)
    if len(ctab) > len(cshow):
        ws.merge_range(rend, 8, rend, 9, "System accounts and other users", f_ital)
        for j, tot in ((10, calc.rcount()), (11, calc.count())):
            L = xl_col_to_name(j)
            oth = tot[1] - csum[j]
            ws.write_formula(rend, j, f"={tot[0][1:]}-SUM({L}{rr0 + 1}:{L}{rend})", f_inum, oth)
            csum[j] += oth
        rend += 1
    ws.merge_range(rend, 8, rend, 9, "Total", f_totl)
    for j in (10, 11):
        L = xl_col_to_name(j)
        ws.write_formula(rend, j, f"=SUM({L}{rr0 + 1}:{L}{rend})", f_tot, csum[j])
    ws.conditional_format(rr0, 10, rr0 + len(cshow) - 1, 10, {"type": "data_bar", "bar_color": NAVY, "bar_solid": False,
                                                              "min_type": "num", "min_value": 0})
    right_end = rend

    # --- timing + breakdown
    r = max(left_end, right_end) + 2
    ws.write(r, 0, "REVIEW TIMING | CURRENT SELECTION", f_sec)
    ws.write(r, 8, "REJECTION BREAKDOWN | CURRENT SELECTION", f_sec)
    timing_rows = [("Open - review started", dict(group="Open", timing="Started")),
                   ("Open - not started", dict(group="Open", timing="Not started")),
                   ("Valid start / end pairs", dict(timing="Valid")),
                   ("Average recorded duration", None),
                   ("Closed - missing timestamps", dict(timing="Missing")),
                   ("Invalid timestamp pairs", dict(timing="Invalid"))]
    for i, (lab, kw) in enumerate(timing_rows, start=1):
        ws.merge_range(r + i, 0, r + i, 4, lab, f_lbl)
        ws.merge_range(r + i, 5, r + i, 6, "", f_val)
        if kw is None:
            fv = calc.avg_duration(person=SEL)
            ws.write_formula(r + i, 5, fv[0], f_dur, fv[1])
        else:
            kw = {k: v for k, v in kw.items()}
            fv = calc.count(person=SEL, **kw)
            ws.write_formula(r + i, 5, fv[0], f_val, fv[1])
    brk = [("Rejected", dict(decision="Rejected")), ("Marked invalid", dict(decision="Marked invalid")),
           ("Sent for reprocess", dict(decision="Sent for reprocess")), ("Canceled", dict(group="Canceled"))]
    for i, (lab, kw) in enumerate(brk, start=1):
        ws.merge_range(r + i, 8, r + i, 10, lab, f_lbl)
        fv = calc.count(person=SEL, **kw)
        ws.write_formula(r + i, 11, fv[0], f_val, fv[1])

    # chart feeder data on Lists
    wl.write(0, 3, "Step")
    wl.write(0, 4, "Count")
    life = [("Requests created", calc.rcount(creator=SELC)), ("Tasks allocated", calc.count(person=SEL)),
            ("Pending", calc.count(person=SEL, group="Open")), ("Accepted", calc.count(person=SEL, group="Success")),
            ("Rejected", calc.count(person=SEL, group="Reject")),
            ("Canceled", calc.count(person=SEL, group="Canceled"))]
    for i, (lab, fv) in enumerate(life, start=1):
        wl.write(i, 3, lab)
        wl.write_formula(i, 4, fv[0], None, fv[1])
    latest = df["created"].max()
    months = [(pd.Timestamp(latest.year, latest.month, 1) - pd.DateOffset(months=11 - i)) for i in range(12)]
    wl.write_row(0, 6, ["Month", "Month start", "Open", "Success", "Reject", "Canceled"])
    for i, m in enumerate(months, start=1):
        wl.write(i, 6, m.strftime("%b %y"))
        wl.write_datetime(i, 7, m.to_pydatetime(), wb.add_format({"num_format": "dd-mmm-yyyy"}))
        for j, g in zip(range(8, 12), ["Open", "Success", "Reject", "Canceled"]):
            fv = calc.count(person=SEL, group=g, month=(f"$H{i + 1}", m), dated=False)
            wl.write_formula(i, j, fv[0], None, fv[1])

    # charts
    cr0 = r + 9
    ws.write(cr0, 0, "Selected user | created > allocated > accepted / rejected", F(bold=True, font_color=GREY, font_size=11))
    ws.write(cr0, 6, "Selected user | tasks per month", F(bold=True, font_color=GREY, font_size=11))
    ch = wb.add_chart({"type": "bar"})
    ch.add_series({"name": "Count", "categories": "=Lists!$D$2:$D$7", "values": "=Lists!$E$2:$E$7",
                   "points": [{"fill": {"color": c}} for c in (NAVY, TEAL, AMBER, GREEN, RED, GREY)],
                   "data_labels": {"value": True, "font": {"bold": True}}, "gap": 50})
    ch.set_legend({"none": True})
    ch.set_title({"none": True})
    ch.set_y_axis({"reverse": True})
    ch.set_x_axis({"major_gridlines": {"visible": True, "line": {"color": "#E3E8EF"}}, "num_format": "#,##0"})
    ch.set_size({"width": 800, "height": 300})
    ch.set_chartarea({"border": {"color": LINE}})
    ws.insert_chart(cr0 + 1, 0, ch, {"x_offset": 2, "y_offset": 4})
    c2 = wb.add_chart({"type": "column", "subtype": "stacked"})
    for j, (g, colr) in zip(range(8, 12), [("Open", AMBER), ("Success", GREEN), ("Reject", RED), ("Canceled", GREY)]):
        L = xl_col_to_name(j)
        c2.add_series({"name": g, "categories": "=Lists!$G$2:$G$13", "values": f"=Lists!${L}$2:${L}$13",
                       "fill": {"color": colr}, "gap": 60})
    c2.set_legend({"position": "bottom"})
    c2.set_title({"none": True})
    c2.set_y_axis({"major_gridlines": {"visible": True, "line": {"color": "#E3E8EF"}}})
    c2.set_size({"width": 780, "height": 300})
    c2.set_chartarea({"border": {"color": LINE}})
    ws.insert_chart(cr0 + 1, 6, c2, {"x_offset": 20, "y_offset": 4})

    # ---- full activity of the selected user (Excel 365 dynamic arrays spill downwards from the first data row)
    det0 = cr0 + 18
    f_cell = F(font_size=9, shrink=True, bottom=1, bottom_color="#EEF1F5")
    f_celld = F(font_size=9, shrink=True, bottom=1, bottom_color="#EEF1F5", num_format="dd-mmm-yy hh:mm", align="left")
    f_dthead = F(bold=True, font_color="#FFFFFF", bg_color=NAVY, font_size=9, border=1, border_color=NAVY)
    f_dtitle = F(bold=True, font_size=12, font_color=NAVY)
    SPILL_ROWS = 1500
    uname = lambda v: (v or "*")
    cu, pe = uname(ic or ip), uname(ip or ic)
    # Dates are turned into text inside the formula (TEXT), so they show as dates whatever the cell format is;
    # the rows are sorted newest first by the real date value.
    DT = '"dd-mmm-yy hh:mm"'
    tables = [
        (0, "f_creator", "REQUESTS CREATED", "REQUESTS CREATED BY ", cu,
         ["Reference", "Process", "Status", "Created"],
         ('r_ref&"",r_process&"",r_status&"",TEXT(r_created,' + DT + ')', "r_creator=f_creator", "r_created",
          "No requests"), (3,)),
        (4, "f_person", "TASKS ALLOCATED", "TASKS ALLOCATED TO ", pe,
         ["Reference", "Status", "Outcome", "Created"],
         ('d_ref&"",d_group&"",d_decision&"",IF(d_created="","",TEXT(d_created,' + DT + '))', "d_person=f_person",
          "d_created", "No tasks"), (3,)),
        (8, "f_person", "TASKS ACCEPTED / REJECTED", "TASKS ACCEPTED / REJECTED BY ", pe,
         ["Reference", "Result", "Reason", "Updated"],
         ('d_ref&"",d_group&"",d_reason&"",IF(d_updated="","",TEXT(d_updated,' + DT + '))',
          '(d_acted=f_person)*(d_group<>"Open")', "d_updated", "No tasks"), (3,)),
    ]
    for c0, fname, none_title, set_title, cur, heads_, spec, datecols in tables:
        cols_expr, cond, by, empty_msg = spec
        body = (f'IFERROR(_xlfn.SORTBY(_xlfn._xlws.FILTER(CHOOSE({{1,2,3,4}},{cols_expr}),{cond}),'
                f'_xlfn._xlws.FILTER({by},{cond}),-1),"{empty_msg}")')
        ws.merge_range(det0, c0, det0, c0 + 3, "", f_dtitle)
        ws.write_formula(det0, c0, f'=IF({fname}="*","{none_title} | pick a user above","{set_title}"&UPPER({fname}))',
                         f_dtitle, f"{none_title} | pick a user above" if cur == "*" else f"{set_title}{cur.upper()}")
        for j, h in enumerate(heads_):
            ws.write(det0 + 1, c0 + j, h, f_dthead)
        ws.write_dynamic_array_formula(det0 + 2, c0, det0 + 2, c0, f'=IF({fname}="*","",{body})', f_cell, "")
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.activate()

    # ---------------------------------------------------------- User Activity sheet
    people = sorted({*(p for p in df["person"] if is_human(p)), *(c for c in df["created_by"] if is_human(c)),
                     *(a for a in df["acted_by"] if is_human(a)), *(c for c in rq["created_by"] if is_human(c))})
    heads = ["User", "Requests created", "Tasks created", "Tasks allocated", "Open", "Success", "Reject", "Canceled",
             "Processed (acted)", "Accepted (acted)", "Rejected (acted)"]
    wu.write(0, 0, "USER ACTIVITY | every user", f_sec)
    for j, h in enumerate(heads):
        wu.write(2, j, h, f_dh)
    wu.set_row(2, 32)
    wu.set_column(0, 0, 28)
    wu.set_column(1, len(heads) - 1, 14)
    calc_u = Calc(df, None, init["dfrom"], init["dto"], rq)
    usum = [0] * len(heads)
    for i, u in enumerate(people):
        r = 3 + i
        a = f"$A{r + 1}"
        ref = (a, u)
        wu.write_string(r, 0, u, f_txt)
        cells = [calc_u.rcount(creator=ref), calc_u.count(creator=ref), calc_u.count(person=ref),
                 calc_u.count(person=ref, group="Open"), calc_u.count(person=ref, group="Success"),
                 calc_u.count(person=ref, group="Reject"), calc_u.count(person=ref, group="Canceled")]
        for j, fv in enumerate(cells, start=1):
            wu.write_formula(r, j, fv[0], f_num, fv[1])
            usum[j] += fv[1]
        ps, rj = calc_u.count(acted=ref, group="Success"), calc_u.count(acted=ref, group="Reject")
        wu.write_formula(r, 9, ps[0], f_num, ps[1])
        wu.write_formula(r, 10, rj[0], f_num, rj[1])
        wu.write_formula(r, 8, f"=J{r + 1}+K{r + 1}", f_num, ps[1] + rj[1])
        usum[8] += ps[1] + rj[1]
        usum[9] += ps[1]
        usum[10] += rj[1]
    tr = 3 + len(people)
    wu.write(tr, 0, "Total", f_totl)
    for j in range(1, len(heads)):
        L = xl_col_to_name(j)
        wu.write_formula(tr, j, f"=SUM({L}4:{L}{tr})", f_tot, usum[j])
    wu.freeze_panes(3, 1)
    if people:
        wu.autofilter(2, 0, tr - 1, len(heads) - 1)
        for j in (1, 4, 5, 6):
            wu.conditional_format(3, j, tr - 1, j, {"type": "data_bar", "bar_color": [NAVY, AMBER, GREEN, RED][
                [1, 4, 5, 6].index(j)], "bar_solid": False, "min_type": "num", "min_value": 0})

    # ---------------------------------------------------------- Pending Queue
    pend = df[df["group"] == "Open"].sort_values("created")
    ph = ["Request Ref", "Request Title", "Task Ref", "Created By", "Allocated Person", "Task Created", "Opened On",
          "Review", "Age (days)"]
    for j, h in enumerate(ph):
        wp.write(0, j, h, f_dh)
    wp.set_row(0, 30)
    for i, row in enumerate(pend.itertuples(index=False), start=1):
        wp.write_string(i, 0, str(row.reference))
        wp.write_string(i, 1, row.title or "")
        wp.write_string(i, 2, row.task_ref or "")
        wp.write_string(i, 3, row.created_by)
        wp.write_string(i, 4, row.person)
        wp.write_datetime(i, 5, row.created.to_pydatetime(), f_dd)
        if pd.notna(row.opened):
            wp.write_datetime(i, 6, row.opened.to_pydatetime(), f_dd)
        wp.write_string(i, 7, row.timing)
        wp.write_formula(i, 8, f"=INT(NOW()-F{i + 1})", wb.add_format({"num_format": "0"}),
                         (datetime.now() - row.created.to_pydatetime()).days)
    for j, w in enumerate([13, 34, 14, 22, 22, 17, 17, 12, 10]):
        wp.set_column(j, j, w)
    wp.freeze_panes(1, 0)
    wp.autofilter(0, 0, max(len(pend), 1), len(ph) - 1)

    # ---------------------------------------------------------- Rejections
    rej = df[df["group"].isin(["Reject", "Canceled"])].sort_values("updated", ascending=False)
    rh = ["Request Ref", "Request Title", "Outcome", "Decision Detail", "Reason", "Acted By", "Allocated Person",
          "Created By", "Task Created", "Task Updated", "Decided By Field"]
    for j, h in enumerate(rh):
        wr.write(0, j, h, f_dh)
    wr.set_row(0, 30)
    for i, row in enumerate(rej.itertuples(index=False), start=1):
        vals = [str(row.reference), row.title or "", row.group, row.decision or "", row.reason or "", row.acted_by,
                row.person, row.created_by]
        for j, v in enumerate(vals):
            wr.write_string(i, j, v)
        wr.write_datetime(i, 8, row.created.to_pydatetime(), f_dd)
        if pd.notna(row.updated):
            wr.write_datetime(i, 9, row.updated.to_pydatetime(), f_dd)
        wr.write_string(i, 10, row.decision_field or "")
    for j, w in enumerate([13, 34, 11, 20, 32, 22, 22, 22, 17, 17, 22]):
        wr.set_column(j, j, w)
    wr.freeze_panes(1, 0)
    wr.autofilter(0, 0, max(len(rej), 1), len(rh) - 1)

    wb.close()
    return {"tasks": n, "people": len(people)}


# ============================================================== main
def parse_date(s):
    try:
        return pd.Timestamp(datetime.strptime(s, "%Y-%m-%d"))
    except ValueError:
        sys.exit(f"Date '{s}' must look like 2026-11-18")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--url", help="Control Room URL (overrides config.json / env DA_URL)")
    ap.add_argument("--db", help="SQLite file (overrides config.json db_path)")
    ap.add_argument("--out", help="output .xlsx path")
    ap.add_argument("--offline", action="store_true", help="build from the local database only")
    ap.add_argument("--mock", action="store_true", help="generated demo data in da_dashboard_mock.db")
    ap.add_argument("--no-details", action="store_true", help="skip per-task detail calls")
    ap.add_argument("--refresh-details", action="store_true", help="re-read details of ALL submitted tasks")
    ap.add_argument("--workers", type=int, default=8)
    ap.add_argument("--every", type=int, metavar="MIN", help="keep running and refresh every MIN minutes")
    ap.add_argument("--person", help="initial allocated-person filter")
    ap.add_argument("--creator", help="initial created-by filter (requests)")
    ap.add_argument("--tz", help="show times in this timezone, e.g. Asia/Kolkata (default UTC)")
    a = ap.parse_args()

    cfg = json.loads(json.dumps(DEFAULT_CFG))
    if os.path.exists(a.config):
        try:
            user_cfg = json.load(open(a.config, encoding="utf-8"))
        except Exception as e:
            sys.exit(f"Cannot read {a.config}: {e}")
        for k in ("verify_ssl", "db_path"):
            if k in user_cfg:
                cfg[k] = user_cfg[k]
        if isinstance(user_cfg.get("tasks_api"), dict):
            cfg["tasks_api"].update({k: v for k, v in user_cfg["tasks_api"].items() if k in ("path", "body", "page_size", "results_path")})
    cfg["control_room_url"] = a.url or CR_URL or os.environ.get("DA_URL") or cfg["control_room_url"]
    if a.tz:
        try:
            pd.Timestamp.now(tz=a.tz)
        except Exception:
            sys.exit(f"Unknown timezone '{a.tz}' (try: pip install tzdata)")
    if not a.mock and not a.offline and (not cfg["control_room_url"] or "your-company" in cfg["control_room_url"]):
        sys.exit("Set CR_URL at the top of the script, pass --url, or set env DA_URL.")
    if a.every is not None and a.every < 1:
        sys.exit("--every must be at least 1 minute")
    cr = MockClient() if a.mock else (None if a.offline else ControlRoom(cfg))
    if a.mock:
        print("Demo mode: generated data")

    if a.every:
        log(f"Scheduled mode: refreshing every {a.every} minute(s). Press Ctrl+C to stop.")
        while True:
            started = time.time()
            run_once(a, cfg, cr)
            time.sleep(max(5, a.every * 60 - (time.time() - started)))
    else:
        sys.exit(0 if run_once(a, cfg, cr) else 1)


def run_once(a, cfg, cr):
    """One full cycle: lock -> sync -> reconcile -> build temp file -> swap in. Returns True on success.
    Any failure before the swap leaves the previous workbook and the database untouched."""
    if not acquire_lock():
        log("Another run is still in progress (lock file present) - skipping this cycle.")
        return False
    t0 = time.time()
    conn = None
    try:
        db_path = "da_dashboard_mock.db" if a.mock else (a.db or cfg["db_path"])
        if a.mock and os.path.exists(db_path):
            os.remove(db_path)
        conn = connect(db_path)
        if cr is not None:
            log("Syncing from Control Room ...")
            sync_task_list(cr, cfg, conn)
            sync_requests(cr, cfg, conn)
            sync_users(cr, conn)
            if not a.no_details:
                sync_details(cr, conn, max(1, a.workers), refresh=a.refresh_details)
        df, meta = prepare(conn, a.tz)
        if a.person and a.person not in set(df["person"]):
            sys.exit(f"--person '{a.person}' not found. Known: {', '.join(sorted(set(df['person']))[:15])} ...")
        rq = prepare_requests(conn, a.tz)
        if a.creator and a.creator not in set(rq["created_by"]):
            sys.exit(f"--creator '{a.creator}' not found among request creators.")
        init = {"person": a.person, "creator": a.creator, "dfrom": None, "dto": None, "tz": a.tz}
        out = a.out or "DA_Request_Dashboard.xlsx"
        tmp = out[:-5] + ".tmp.xlsx" if out.lower().endswith(".xlsx") else out + ".tmp.xlsx"
        stats = build_workbook(df, rq, meta, tmp, init)
        if publish(tmp, out, retries=1 if a.every else 3, wait=10):
            where = f"Dashboard written: {os.path.abspath(out)}"
        else:
            where = (f"WAITING: {out} is open in Excel, so the new version is saved as {tmp} and will replace it "
                     "on a later run once the file is closed.")
        g = df["group"].value_counts()
        rs = rq["status"].value_counts()
        log(f"{where}\n  {len(rq)} requests (Open {rs.get('Open', 0)} | Success {rs.get('Success', 0)} | "
            f"Failed {rs.get('Failed', 0)})\n  {stats['tasks']} tasks | Open {g.get('Open', 0)} | "
            f"Success {g.get('Success', 0)} | Reject {g.get('Reject', 0)} | Canceled {g.get('Canceled', 0)} | "
            f"users {stats['people']} | "
            f"{time.time() - t0:.0f}s")
        return True
    except SystemExit:
        raise
    except Exception as e:
        log(f"RUN FAILED (previous dashboard kept): {type(e).__name__}: {e}")
        return False
    finally:
        if conn is not None:
            conn.close()
        release_lock()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nStopped.")
