# How to Create a Per-User Activity Dashboard for Document Automation Requests and Tasks Using Control Room APIs and Excel

Builds an Excel dashboard of Document Automation requests and human tasks per user, using the Automation 360 Control
Room APIs.


*Sample data.*

## Features

- Requests created per user, with status (Open, Success, Failed)
- Tasks allocated to each user, with accepted, rejected and canceled counts and rejection reasons
- Queue view: open and unassigned tasks with their age
- Filter by **Allocated person** or **Created by** to see everything one user did
- Local SQLite copy of the data, so later runs are fast
- Optional refresh every N minutes

## Requirements

- Python 3.10+
- Microsoft Excel 365 or 2021
- A Control Room user that can view Automation Co-Pilot requests and tasks (read-only service account recommended)

## Installation

```bash
git clone <repository-url>
cd <repository-folder>
pip install requests pandas xlsxwriter tzdata
```

## Configuration

Edit the block at the top of `da_dashboard.py`:

```python
CR_URL = "https://your-company.my.automationanywhere.digital"
CR_USERNAME = "YOUR_USERNAME"
CR_PASSWORD = "YOUR_PASSWORD"
```

If the username and password are left as placeholders, the script prompts for them at run time. The environment
variables `DA_USER`, `DA_PASSWORD` and `DA_URL` are also read. Never commit real credentials.

## Usage

```bash
python da_dashboard.py                # sync and build DA_Request_Dashboard.xlsx
python da_dashboard.py --every 15     # refresh every 15 minutes
python da_dashboard.py --offline      # rebuild from the local database
python da_dashboard.py --mock         # demo workbook from generated data
```

The first run fetches the detail of every submitted task and takes longer. Later runs only fetch new or changed records.

| Option | Purpose |
|---|---|
| `--person NAME` / `--creator NAME` | Open with a user pre-selected |
| `--out FILE.xlsx` | Output workbook path |
| `--tz ZONE` | Time zone for displayed times, e.g. `Asia/Kolkata` (default UTC) |
| `--no-details` | Skip per-task detail calls |
| `--refresh-details` | Re-read details of all submitted tasks |
| `--workers N` | Parallel detail calls (default 8) |

## Output

| File | Content |
|---|---|
| `DA_Request_Dashboard.xlsx` | Dashboard, User Activity, Pending Queue, Rejections, Requests and Data sheets |
| `da_tracker.db` | Local SQLite copy of requests and tasks |
| `da_dashboard.log` | Run log |

If the workbook is open in Excel during a refresh, the new version is saved as `DA_Request_Dashboard.tmp.xlsx` and
replaces the old one on a later run. Close and reopen the workbook to see updates.

### Sample output

The [`output/`](output/) folder holds the three files produced by `python da_dashboard.py --mock`. They contain
generated demo data only, so you can see the result without a Control Room.

| Sample file | What to look at |
|---|---|
| [`output/DA_Request_Dashboard.xlsx`](output/DA_Request_Dashboard.xlsx) | The dashboard workbook. Pick a user in the yellow filter cells. |
| [`output/da_dashboard.log`](output/da_dashboard.log) | A run log with the counts printed at the end of each run |
| [`output/da_tracker.db`](output/da_tracker.db) | The SQLite database. Open it with any SQLite viewer; the main tables are `tasks` and `requests`. |

## APIs used

All calls are read-only.

| Method | Endpoint |
|---|---|
| POST | `/v2/authentication` |
| POST | `/aari/v3/requests/list` |
| POST | `/aari/v2/tasks/humans/list` |
| GET | `/aari/v2/tasks/humans/{id}` |
| POST | `/v2/usermanagement/users/list` (optional) |

## Limitations

- The Control Room does not always record which user acted on a task. Where it did not, the allocated person is shown.
- Reviewers are shown only when the process uses approval steps.
- Processing phase and learning instance are not part of the Co-Pilot request and task data.
- The per-user lists need Excel 365 or 2021.

## Security

`da_tracker.db`, `da_dashboard.log` and the generated workbook contain user and request data. Keep them out of version
control:

```gitignore
*.db
*.log
*.xlsx
*.lock
!output/*
```

The last line keeps the demo files in `output/` tracked.

## Documentation

See `KB_DA_User_Activity_Dashboard.docx` for the full step-by-step guide.
