# ProjectPulse — Local Excel Sync

This version keeps the existing ProjectPulse scheduler and Gantt UI, but adds a small local Python bridge so the browser can read and write a real `.xlsx` file.

## 1. Set the Excel file path

Open `server.py` and find:

```python
EXCEL_FILE = Path(
    os.environ.get("PROJECTPULSE_EXCEL", str(BASE_DIR / "Project_data.xlsx"))
)
```

If your workbook is somewhere else, replace the default path with your real path, for example:

```python
EXCEL_FILE = Path(r"C:\Users\Akanksh\Documents\ProjectPulse\Project_data.xlsx")
```

The `r` before the Windows path is intentional.

## 2. Install the one Python dependency

Open Command Prompt in this folder and run:

```text
pip install -r requirements.txt
```

## 3. Start ProjectPulse

Either double-click:

```text
run_projectpulse.bat
```

or run:

```text
python server.py
```

Then open:

```text
http://127.0.0.1:8000
```

## 4. What happens now

When the app starts:

```text
ProjectPulse → /api/tasks → server.py → Project_data.xlsx
```

The workbook is loaded automatically. You do not need to upload it manually.

When you change a task's:

- Start Date
- Duration
- Dependency
- Status / Progress

the existing scheduler runs first, the Gantt is redrawn, and the current task data is then sent back to `server.py`, which updates the matching Excel rows.

The Excel workbook remains the source file on disk.

## 5. Live Server in VS Code

You can still use VS Code Live Server if you prefer, but **`server.py` must also be running** because Live Server only serves the web files; it does not have permission to edit an arbitrary Excel file.

If the UI is running on `127.0.0.1:5500`, this version automatically sends Excel requests to:

```text
http://127.0.0.1:8000
```

So the recommended setup is simply:

```text
Terminal 1: python server.py
Browser:   http://127.0.0.1:8000
```

## 6. Excel column names

The bridge is intentionally flexible. It recognizes common variants of:

- `ID`
- `Activity`
- `Start Date`
- `End Date`
- `Duration`, `Duration (Weeks)`, or `Duration (Days)`
- `Depends On` / `Dependency`
- `Progress`
- `Status`

For `Duration (Days)`, ProjectPulse's internal duration in weeks is converted to days when it is written back to Excel.

The bridge also looks for the header row instead of assuming the headers are in row 1, which lets workbooks have project titles or blank rows above the table.

## Important

Do not keep the Excel workbook open in a way that locks the file while ProjectPulse is trying to save it. If Excel reports the workbook as locked, close the workbook and make the change again in ProjectPulse.

## Shared Excel + ProjectPulse editing

ProjectPulse now uses a durable local synchronization queue when Microsoft Excel has the workbook locked.

- Browser edits are accepted immediately.
- Only tasks changed by ProjectPulse are sent to the server, avoiding stale full-project overwrites.
- If Excel is writable, the changed task rows are saved immediately.
- If Excel is locked/open, the changes are stored in `.projectpulse_pending_sync.json` next to the application.
- A background Python worker retries every 3 seconds.
- When Excel becomes writable, queued task changes are merged into the current workbook and the queue is cleared.
- The scheduler, dependency calculations, milestones, delays, Gantt, dashboard, and local Excel loading remain unchanged.

The queue is intentionally local to the ProjectPulse machine. It is a persistence/retry mechanism for Excel locking; it is not a replacement for true multi-user Excel collaboration.

## 7. Downtime events

Downtime is implemented as a real scheduled task, not as a visual annotation.

From the Gantt page, hover over any normal activity row and click the small `+` button to add a downtime immediately after that activity.

A downtime has:

- its own Task ID
- a downtime name (shown in the Activity column)
- a reason
- a Start Date
- a Duration (using the same week-based duration model as the current Gantt)
- an automatically calculated End Date
- a dependency relationship

The downtime is rendered as a red Gantt bar. Hovering over the red bar shows its name, reason, start date, end date, and duration.

Most importantly, downtime participates in the scheduler. The activity immediately following the downtime is made dependent on the downtime, so extending or moving the downtime automatically cascades through downstream activities.

When a downtime is first created, ProjectPulse inserts it into the `Project_Data` task table at the correct position. If the workbook does not already contain the downtime metadata columns, the local Excel bridge creates:

- `Type`
- `Reason`
- `Parent Activity ID`

Downtime rows use `Type = Downtime`.

The Dashboard no longer has an `Upload Excel` control. The configured workbook is loaded automatically by `server.py`.


## Downtime persistence (2026-08-20)

Downtime rows are no longer stored in `Project_Data.xlsx`. ProjectPulse stores them in the local `downtime_data.json` file beside `server.py`. The browser combines the Excel activities and JSON downtime events at runtime, assigns continuous visible IDs (1..N), and uses the downtime rows in the dependency scheduler. Normal activities continue to be written to Excel using their stable Excel IDs, so virtual downtime IDs are never written into the workbook. Existing workbooks containing `DT_*` rows are automatically migrated out of Excel on first load.
