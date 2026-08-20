"""
ProjectPulse local Excel bridge.

Run this file instead of opening index.html directly:
    python server.py

Then open:
    http://127.0.0.1:8000

IMPORTANT: Set EXCEL_FILE below to the Excel file you want ProjectPulse
            to load and update automatically.
"""

from __future__ import annotations

import json
import os
import threading
import time
from datetime import date, datetime
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from copy import copy
from urllib.parse import urlparse

from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent

# ============================================================
# CHANGE THIS ONE LINE if your Excel file is somewhere else.
# Example:
# EXCEL_FILE = Path(r"C:\Users\Akanksh\Documents\ProjectPulse\Project_data.xlsx")
# ============================================================
EXCEL_FILE = Path(r"C:\Users\CAK6BAN\OneDrive - Bosch Group\ProjectPulse\Project_data(synthetic).xlsx")

HOST = "127.0.0.1"
PORT = 8000

_write_lock = threading.Lock()
_queue_lock = threading.Lock()
SYNC_INTERVAL_SECONDS = 3
PENDING_QUEUE_FILE = BASE_DIR / ".projectpulse_pending_sync.json"
DOWNTIME_FILE = BASE_DIR / "downtime_data.json"


def normalize_header(value) -> str:
    return "".join(ch.lower() for ch in str(value or "") if ch.isalnum())


def find_header_row(ws) -> int:
    """Find the row containing the real table headers.

    ProjectPulse sheets can have title/blank rows above the table, so the
    headers are not assumed to be row 1.
    """
    for row_number in range(1, min(ws.max_row, 30) + 1):
        values = [ws.cell(row_number, col).value for col in range(1, ws.max_column + 1)]
        normalized = {normalize_header(v) for v in values if v not in (None, "")}
        if "activity" in normalized and ("id" in normalized or "startdate" in normalized):
            return row_number
    return 1


def header_map(ws, header_row: int):
    mapping = {}
    for col in range(1, ws.max_column + 1):
        value = ws.cell(header_row, col).value
        if value not in (None, ""):
            mapping[normalize_header(value)] = col
    return mapping


def find_col(headers: dict[str, int], *candidates: str):
    for candidate in candidates:
        col = headers.get(normalize_header(candidate))
        if col:
            return col

    # Prefix fallback for headers such as "Duration (Weeks)".
    normalized_headers = list(headers.items())
    for candidate in candidates:
        prefix = normalize_header(candidate)
        for key, col in normalized_headers:
            if key.startswith(prefix):
                return col
    return None


def serialize_value(value):
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    return value


def _load_downtime_state():
    if not DOWNTIME_FILE.exists():
        return {"downtimes": [], "dependencyOverrides": []}
    try:
        with DOWNTIME_FILE.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
        if not isinstance(data, dict):
            return {"downtimes": [], "dependencyOverrides": []}
        return {
            "downtimes": data.get("downtimes") if isinstance(data.get("downtimes"), list) else [],
            "dependencyOverrides": data.get("dependencyOverrides") if isinstance(data.get("dependencyOverrides"), list) else [],
        }
    except Exception as exc:
        print(f"[Downtime] Could not read {DOWNTIME_FILE}: {exc}")
        return {"downtimes": [], "dependencyOverrides": []}


def _write_downtime_state(state):
    payload = {
        "version": 2,
        "downtimes": state.get("downtimes", []),
        "dependencyOverrides": state.get("dependencyOverrides", []),
    }
    temp = DOWNTIME_FILE.with_suffix(".tmp")
    with temp.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, default=str)
    os.replace(temp, DOWNTIME_FILE)


def _row_is_downtime(row, headers):
    activity = str(row.get("Activity", "")).strip()
    task_type = str(row.get("Type", row.get("Task Type", "")) or "").strip().lower()
    return task_type == "downtime" or bool(__import__("re").match(r"^dt_\d+$", activity, __import__("re").IGNORECASE))


def migrate_excel_downtimes():
    """Move legacy downtime rows out of Excel into downtime_data.json.

    Existing normal activities are renumbered back to a clean 1..N Excel
    sequence and dependencies that pointed at a removed downtime are restored
    to the real activity immediately before that downtime.
    """
    if not EXCEL_FILE.exists():
        return

    with _write_lock:
        wb = load_workbook(EXCEL_FILE)
        sheet_name = "Project_Data" if "Project_Data" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        header_row = find_header_row(ws)
        headers = header_map(ws, header_row)
        id_col = find_col(headers, "ID")
        activity_col = find_col(headers, "Activity")
        type_col = find_col(headers, "Type", "Task Type")
        reason_col = find_col(headers, "Reason", "Downtime Reason")
        start_col = find_col(headers, "Start Date")
        duration_col = find_col(headers, "Duration (Weeks)", "Duration (Days)", "Duration")
        depends_col = find_col(headers, "Depends On", "Dependency")
        parent_col = find_col(headers, "Parent Activity ID", "Parent ID")

        if not id_col or not activity_col:
            wb.close()
            return

        rows = []
        for row_number in range(header_row + 1, ws.max_row + 1):
            if all(ws.cell(row_number, c).value in (None, "") for c in range(1, ws.max_column + 1)):
                continue
            rows.append((row_number, {str(ws.cell(header_row, c).value): ws.cell(row_number, c).value for c in range(1, ws.max_column + 1) if ws.cell(header_row, c).value not in (None, "")}))

        downtime_rows = [(rn, row) for rn, row in rows if _row_is_downtime(row, headers)]
        if not downtime_rows:
            wb.close()
            return

        state = _load_downtime_state()
        existing_uids = {str(item.get("uid")) for item in state["downtimes"] if item.get("uid")}
        migrated = list(state["downtimes"])

        # Determine each downtime's nearest preceding real activity.
        previous_real_id = None
        downtime_info = {}
        for rn, row in rows:
            old_id = parse_id(row.get(str(ws.cell(header_row, id_col).value)))
            if any(rn == drn for drn, _ in downtime_rows):
                uid = f"dt-migrated-{old_id}-{rn}"
                if uid not in existing_uids:
                    raw_duration = row.get(str(ws.cell(header_row, duration_col).value)) if duration_col else 1
                    try:
                        duration = max(1, int(float(raw_duration)))
                    except (TypeError, ValueError):
                        duration = 1
                    if duration_col and "day" in str(ws.cell(header_row, duration_col).value or "").lower():
                        duration = max(1, int(round(duration / 7)))
                    start = serialize_value(row.get(str(ws.cell(header_row, start_col).value))) if start_col else None
                    migrated.append({
                        "uid": uid,
                        "name": str(row.get(str(ws.cell(header_row, activity_col).value), "Downtime") or "Downtime"),
                        "reason": str(row.get(str(ws.cell(header_row, reason_col).value), "") or ""),
                        "startDate": start,
                        "duration": duration,
                        "afterOriginalId": previous_real_id,
                        "sequence": sum(1 for item in migrated if parse_id(item.get("afterOriginalId")) == parse_id(previous_real_id)),
                    })
                downtime_info[old_id] = previous_real_id
            else:
                previous_real_id = old_id

        # Remove downtime rows from the workbook, then renumber remaining
        # physical rows and remap dependencies to the new stable IDs.
        for rn, _ in reversed(downtime_rows):
            ws.delete_rows(rn, 1)

        remaining = []
        for rn in range(header_row + 1, ws.max_row + 1):
            value = ws.cell(rn, id_col).value
            if value in (None, ""):
                continue
            remaining.append((rn, parse_id(value)))

        old_to_new = {old_id: index + 1 for index, (_, old_id) in enumerate(remaining)}

        for index, (rn, old_id) in enumerate(remaining):
            new_id = index + 1
            ws.cell(rn, id_col).value = new_id
            if depends_col:
                raw = ws.cell(rn, depends_col).value
                dep_values = []
                for token in str(raw or "").replace(";", ",").split(","):
                    token = token.strip()
                    if not token:
                        continue
                    dep = parse_id(token)
                    if dep in downtime_info:
                        dep = downtime_info[dep]
                    if dep in old_to_new:
                        dep_values.append(old_to_new[dep])
                ws.cell(rn, depends_col).value = ",".join(str(x) for x in dict.fromkeys(dep_values))

        wb.save(EXCEL_FILE)
        wb.close()
        _write_downtime_state({"downtimes": migrated, "dependencyOverrides": state["dependencyOverrides"]})
        print(f"[Downtime] Migrated {len(downtime_rows)} downtime row(s) from Excel to {DOWNTIME_FILE.name}.")


def read_downtime_state():
    migrate_excel_downtimes()
    return _load_downtime_state()

def read_rows():
    migrate_excel_downtimes()
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}\n"
            "Set EXCEL_FILE in server.py to your real Project_Data.xlsx path."
        )

    wb = load_workbook(EXCEL_FILE, data_only=True)
    sheet_name = "Project_Data" if "Project_Data" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    header_row = find_header_row(ws)
    headers = header_map(ws, header_row)

    id_col = find_col(headers, "ID")
    if not id_col:
        raise ValueError(f'No "ID" column found in worksheet "{sheet_name}".')

    rows = []
    for row_number in range(header_row + 1, ws.max_row + 1):
        row_values = [ws.cell(row_number, col).value for col in range(1, ws.max_column + 1)]
        if all(value in (None, "") for value in row_values):
            continue

        row = {}
        for col in range(1, ws.max_column + 1):
            header = ws.cell(header_row, col).value
            if header in (None, ""):
                continue
            row[str(header)] = serialize_value(ws.cell(row_number, col).value)

        if row.get(ws.cell(header_row, id_col).value) in (None, ""):
            continue
        rows.append(row)

    return {
        "file": str(EXCEL_FILE),
        "sheet": sheet_name,
        "headerRow": header_row,
        "rows": rows,
    }


def parse_id(value):
    try:
        number = float(value)
        return int(number) if number.is_integer() else number
    except (TypeError, ValueError):
        return value


def task_map_by_id(tasks):
    return {parse_id(task.get("id")): task for task in tasks}


def set_date_cell(cell, value):
    if not value:
        return
    try:
        year, month, day = [int(part) for part in str(value)[:10].split("-")]
        cell.value = datetime(year, month, day)
        # Preserve an existing date format where possible.
        if not cell.number_format or cell.number_format == "General":
            cell.number_format = "dd-mmm-yy"
    except Exception:
        # Do not destroy an existing cell with an invalid client value.
        pass


def _is_excel_lock_error(exc: Exception) -> bool:
    """Return True for the common Windows errors caused by Excel locking the workbook."""
    if isinstance(exc, PermissionError):
        return True
    text = str(exc).lower()
    return any(token in text for token in ("permission denied", "being used by another process", "access is denied", "winerror 32"))


def _load_pending_queue():
    if not PENDING_QUEUE_FILE.exists():
        return []
    try:
        with PENDING_QUEUE_FILE.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
        return data if isinstance(data, list) else []
    except Exception as exc:
        print(f"[Sync] Could not read pending queue: {exc}")
        return []


def _write_pending_queue(queue):
    temp = PENDING_QUEUE_FILE.with_suffix(".tmp")
    with temp.open("w", encoding="utf-8") as handle:
        json.dump(queue, handle, ensure_ascii=False, indent=2, default=str)
    os.replace(temp, PENDING_QUEUE_FILE)


def _merge_pending_patch(patch):
    """Keep the newest patch for each task ID, preserving other pending tasks."""
    task_id = parse_id(patch.get("id"))
    queue = _load_pending_queue()
    replaced = False
    for index, existing in enumerate(queue):
        if parse_id(existing.get("id")) == task_id:
            queue[index] = patch
            replaced = True
            break
    if not replaced:
        queue.append(patch)
    _write_pending_queue(queue)
    return queue


def _task_patch(task):
    return {
        "id": task.get("id"),
        "activity": task.get("activity", ""),
        "startDate": task.get("startDate"),
        "endDate": task.get("endDate"),
        "duration": task.get("duration"),
        "dependencies": task.get("dependencies") or [],
        "status": task.get("status") or "Not Started",
        "isDowntime": bool(task.get("isDowntime")),
        "reason": task.get("reason", ""),
        "parentActivityId": task.get("parentActivityId"),
        "insertAfterId": task.get("insertAfterId"),
        "originalId": task.get("originalId"),
        "insertAfterOriginalId": task.get("insertAfterOriginalId"),
    }


def _ensure_column(ws, header_row: int, headers: dict[str, int], label: str) -> int:
    """Return a column for a header, appending the header if necessary."""
    existing = find_col(headers, label)
    if existing:
        return existing

    new_col = ws.max_column + 1
    ws.cell(header_row, new_col).value = label
    headers[normalize_header(label)] = new_col
    return new_col


def _copy_row_style(ws, source_row: int, target_row: int) -> None:
    """Copy the parent row's presentation to a newly inserted downtime row."""
    for col in range(1, ws.max_column + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.alignment:
            target.alignment = copy(source.alignment)
        if source.protection:
            target.protection = copy(source.protection)
    if source_row in ws.row_dimensions:
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def _write_task_row(ws, header_row: int, headers: dict[str, int], row_number: int, task: dict) -> None:
    """Write a ProjectPulse task/downtime patch into one Excel row."""
    id_col = find_col(headers, "ID")
    activity_col = find_col(headers, "Activity")
    start_col = find_col(headers, "Start Date")
    end_col = find_col(headers, "End Date")
    duration_col = find_col(headers, "Duration (Days)", "Duration (Weeks)", "Duration")
    depends_col = find_col(headers, "Depends On", "Dependency")
    progress_col = find_col(headers, "Progress")
    status_col = find_col(headers, "Status")

    if id_col:
        ws.cell(row_number, id_col).value = task.get("excelId", task.get("originalId", task.get("id")))
    if activity_col:
        ws.cell(row_number, activity_col).value = task.get("activity", "")
    if start_col:
        set_date_cell(ws.cell(row_number, start_col), task.get("startDate"))
    if end_col:
        set_date_cell(ws.cell(row_number, end_col), task.get("endDate"))
    if duration_col and task.get("duration") is not None:
        duration_weeks = max(1, int(round(float(task.get("duration", 1)))))
        duration_header = str(ws.cell(header_row, duration_col).value or "").lower()
        ws.cell(row_number, duration_col).value = duration_weeks * 7 if "day" in duration_header else duration_weeks
    if depends_col:
        deps = task.get("dependencies") or []
        ws.cell(row_number, depends_col).value = ",".join(str(int(dep)) for dep in deps)
    if progress_col:
        ws.cell(row_number, progress_col).value = task.get("status", "Not Started")
    if status_col:
        ws.cell(row_number, status_col).value = task.get("status", "Not Started")

    if task.get("isDowntime"):
        type_col = _ensure_column(ws, header_row, headers, "Type")
        reason_col = _ensure_column(ws, header_row, headers, "Reason")
        parent_col = _ensure_column(ws, header_row, headers, "Parent Activity ID")
        ws.cell(row_number, type_col).value = "Downtime"
        ws.cell(row_number, reason_col).value = task.get("reason", "")
        ws.cell(row_number, parent_col).value = task.get("parentActivityId")



def _apply_task_patches_to_workbook(patches):
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}\n"
            "Set EXCEL_FILE in server.py to your real Project_Data.xlsx path."
        )

    with _write_lock:
        wb = load_workbook(EXCEL_FILE)
        sheet_name = "Project_Data" if "Project_Data" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        header_row = find_header_row(ws)
        headers = header_map(ws, header_row)

        id_col = find_col(headers, "ID")
        if not id_col:
            wb.close()
            raise ValueError(f'No "ID" column found in worksheet "{sheet_name}".')

        excel_rows = {}
        for row_number in range(header_row + 1, ws.max_row + 1):
            value = ws.cell(row_number, id_col).value
            if value not in (None, ""):
                excel_rows[parse_id(value)] = row_number

        updated = 0

        for task in patches:
            if task.get("isDowntime"):
                continue
            task_id = parse_id(task.get("originalId", task.get("id")))
            row_number = excel_rows.get(task_id)

            # A new downtime is inserted into the same task table. Its
            # insertAfterId is persisted only for the creation operation;
            # once the row exists, subsequent edits behave like normal task
            # patches.
            if row_number is None:
                if not task.get("isDowntime"):
                    continue

                insert_after_id = parse_id(task.get("insertAfterId"))
                insert_after_row = excel_rows.get(insert_after_id)

                if insert_after_row is None:
                    row_number = ws.max_row + 1
                else:
                    row_number = insert_after_row + 1
                    ws.insert_rows(row_number, 1)
                    _copy_row_style(ws, row_number - 1, row_number)

                    # Excel row numbers after the insertion move down by one.
                    for existing_id, existing_row in list(excel_rows.items()):
                        if existing_row >= row_number:
                            excel_rows[existing_id] = existing_row + 1

                excel_rows[task_id] = row_number

            _write_task_row(ws, header_row, headers, row_number, task)
            updated += 1

        if updated:
            wb.save(EXCEL_FILE)
        wb.close()

    return {"updatedRows": updated, "file": str(EXCEL_FILE), "sheet": sheet_name}


def _is_legacy_placeholder_row(ws, header_row, headers, row_number):
    """Identify the old sample 'Downtime' row that predates real downtime support."""
    activity_col = find_col(headers, "Activity")
    type_col = find_col(headers, "Type", "Task Type")
    reason_col = find_col(headers, "Reason", "Downtime Reason")
    parent_col = find_col(headers, "Parent Activity ID", "Parent ID")

    activity = str(ws.cell(row_number, activity_col).value or "").strip().lower() if activity_col else ""
    task_type = str(ws.cell(row_number, type_col).value or "").strip() if type_col else ""
    reason = str(ws.cell(row_number, reason_col).value or "").strip() if reason_col else ""
    parent = ws.cell(row_number, parent_col).value if parent_col else None

    return (
        activity == "downtime"
        and not task_type
        and not reason
        and parent in (None, "")
    )


def _apply_ordered_tasks_to_workbook(tasks):
    """Persist the complete ordered Project.tasks list to Excel.

    ``originalId`` is the stable physical-row identity.  The visible ``id``
    is allowed to change whenever a downtime is inserted/removed because the
    UI uses continuous 1..N IDs.  Using originalId lets us safely:
      * insert a new downtime at the correct row,
      * update every existing row after IDs are renumbered, and
      * DELETE a downtime row that is no longer present in Project.tasks,
        including after a server restart.
    """
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}\n"
            "Set EXCEL_FILE in server.py to your real Project_Data.xlsx path."
        )

    with _write_lock:
        wb = load_workbook(EXCEL_FILE)
        sheet_name = "Project_Data" if "Project_Data" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        header_row = find_header_row(ws)
        headers = header_map(ws, header_row)
        id_col = find_col(headers, "ID")

        if not id_col:
            wb.close()
            raise ValueError(f'No "ID" column found in worksheet "{sheet_name}".')

        # -------------------------------------------------------------
        # 1. Remove the obsolete placeholder downtime.
        # -------------------------------------------------------------
        legacy_rows = []
        for row_number in range(ws.max_row, header_row, -1):
            if _is_legacy_placeholder_row(ws, header_row, headers, row_number):
                legacy_rows.append(row_number)
                ws.delete_rows(row_number, 1)

        # -------------------------------------------------------------
        # 2. Build the set of Excel rows that are still part of the
        #    current Project.tasks collection.
        #
        #    This is the critical deletion step.  After a restart,
        #    Project.tasks is rebuilt from Excel and every loaded task has
        #    _excelOriginalId = its physical Excel ID.  If a downtime is
        #    deleted in the UI, its original ID is absent from this set,
        #    so its physical Excel row is removed here.
        # -------------------------------------------------------------
        submitted_original_ids = {
            parse_id(task.get("originalId"))
            for task in tasks
            if task.get("originalId") not in (None, "")
        }

        removed_rows = 0
        for row_number in range(ws.max_row, header_row, -1):
            value = ws.cell(row_number, id_col).value
            if value in (None, ""):
                continue

            excel_id = parse_id(value)

            if excel_id not in submitted_original_ids:
                ws.delete_rows(row_number, 1)
                removed_rows += 1

        # Rebuild physical Excel-ID -> row-number map after deletions.
        excel_rows = {}
        for row_number in range(header_row + 1, ws.max_row + 1):
            value = ws.cell(row_number, id_col).value
            if value not in (None, ""):
                excel_rows[parse_id(value)] = row_number

        # -------------------------------------------------------------
        # 3. Insert any newly-created downtime rows.
        #
        #    New rows have originalId == null.  They are inserted according
        #    to their Project.tasks index, which is the exact Gantt/Excel
        #    order.  Existing rows are never overwritten.
        # -------------------------------------------------------------
        for index, task in enumerate(tasks):
            original_id = task.get("originalId")

            if original_id not in (None, ""):
                task_id = parse_id(original_id)
                row_number = excel_rows.get(task_id)

                if row_number is None:
                    # A supposedly-existing row disappeared. Insert it at
                    # the requested project position instead of writing into
                    # an unrelated physical row.
                    target_row = header_row + 1 + index
                    ws.insert_rows(target_row, 1)

                    if target_row > header_row + 1:
                        _copy_row_style(ws, target_row - 1, target_row)

                    task["_resolvedRow"] = target_row

                    for existing_id, existing_row in list(excel_rows.items()):
                        if existing_row >= target_row:
                            excel_rows[existing_id] = existing_row + 1
                else:
                    task["_resolvedRow"] = row_number

                continue

            # New downtime: insert it, never overwrite an existing task.
            target_row = header_row + 1 + index
            ws.insert_rows(target_row, 1)

            if target_row > header_row + 1:
                _copy_row_style(ws, target_row - 1, target_row)

            task["_resolvedRow"] = target_row

            for existing_id, existing_row in list(excel_rows.items()):
                if existing_row >= target_row:
                    excel_rows[existing_id] = existing_row + 1

        # -------------------------------------------------------------
        # 4. Rebuild the map one more time using the stable ORIGINAL IDs.
        #    Visible IDs have not been rewritten yet.
        # -------------------------------------------------------------
        refreshed = {}
        for row_number in range(header_row + 1, ws.max_row + 1):
            value = ws.cell(row_number, id_col).value
            if value not in (None, ""):
                refreshed[parse_id(value)] = row_number

        # -------------------------------------------------------------
        # 5. Write the complete Project.tasks list in its exact order.
        # -------------------------------------------------------------
        for task in tasks:
            original_id = task.get("originalId")

            if original_id not in (None, ""):
                row_number = refreshed.get(
                    parse_id(original_id),
                    task.get("_resolvedRow")
                )
            else:
                row_number = task.get("_resolvedRow")

            if not row_number:
                continue

            _write_task_row(
                ws,
                header_row,
                headers,
                int(row_number),
                task
            )

        # Helper values never go into the workbook.
        for task in tasks:
            task.pop("_resolvedRow", None)

        wb.save(EXCEL_FILE)
        wb.close()

    return {
        "updatedRows": len(tasks),
        "removedRows": removed_rows,
        "removedLegacyRows": len(legacy_rows),
        "file": str(EXCEL_FILE),
        "sheet": sheet_name,
    }


def save_ordered_tasks(tasks):
    """Persist a full ordered task list; queue the whole operation if locked.

    An empty list is valid: it means the last task/downtime was deleted and
    the corresponding Excel rows should be removed.
    """
    if tasks is None:
        tasks = []

    operation = {
        "__operation": "reorder",
        "tasks": tasks,
    }

    try:
        result = _apply_ordered_tasks_to_workbook(tasks)
        with _queue_lock:
            queue = [item for item in _load_pending_queue()
                     if item.get("__operation") != "reorder"]
            _write_pending_queue(queue)
        return {"queued": False, "pendingRows": 0, **result}
    except Exception as exc:
        if not _is_excel_lock_error(exc):
            raise

        with _queue_lock:
            queue = [item for item in _load_pending_queue()
                     if item.get("__operation") != "reorder"]
            queue.append(operation)
            _write_pending_queue(queue)

        print(f"[Sync] Excel is locked; queued ordered project ({len(tasks)} task(s)).")
        return {
            "updatedRows": 0,
            "queued": True,
            "pendingRows": len(tasks),
            "file": str(EXCEL_FILE),
            "message": "Excel is currently open/locked. The ordered project is queued and will retry automatically.",
        }


def save_tasks(tasks):
    """Try to write immediately; if Excel is locked, persist patches for retry."""
    patches = [_task_patch(task) for task in tasks if task.get("id") not in (None, "")]
    if not patches:
        return {"updatedRows": 0, "queued": False, "pendingRows": len(_load_pending_queue()), "file": str(EXCEL_FILE)}

    # Merge with any earlier queued edits first. This gives us a durable queue
    # and ensures a newer ProjectPulse edit replaces an older queued edit for
    # the same activity.
    with _queue_lock:
        queue = _load_pending_queue()
        by_id = {parse_id(item.get("id")): item for item in queue}
        for patch in patches:
            by_id[parse_id(patch.get("id"))] = patch
        merged = list(by_id.values())

    try:
        result = _apply_task_patches_to_workbook(merged)
        with _queue_lock:
            _write_pending_queue([])
        result.update({"queued": False, "pendingRows": 0})
        return result
    except Exception as exc:
        if not _is_excel_lock_error(exc):
            raise
        with _queue_lock:
            _write_pending_queue(merged)
        print(f"[Sync] Excel is locked; queued {len(merged)} task patch(es).")
        return {
            "updatedRows": 0,
            "queued": True,
            "pendingRows": len(merged),
            "file": str(EXCEL_FILE),
            "message": "Excel is currently open/locked. Changes are safely queued and will retry automatically."
        }


def _sync_pending_queue_once():
    with _queue_lock:
        queue = _load_pending_queue()
    if not queue:
        return
    try:
        ordered = next((item for item in queue if item.get("__operation") == "reorder"), None)
        if ordered:
            result = _apply_ordered_tasks_to_workbook(ordered.get("tasks", []))
        else:
            result = _apply_task_patches_to_workbook(queue)
        with _queue_lock:
            _write_pending_queue([])
        print(f"[Sync] Flushed {result['updatedRows']} queued task change(s) to Excel.")
    except Exception as exc:
        if not _is_excel_lock_error(exc):
            print(f"[Sync] Pending queue error: {exc}")


def _sync_worker():
    while True:
        time.sleep(SYNC_INTERVAL_SECONDS)
        try:
            _sync_pending_queue_once()
        except Exception as exc:
            print(f"[Sync] Worker error: {exc}")


def pending_sync_status():
    queue = _load_pending_queue()
    return {"pendingRows": len(queue), "queued": bool(queue), "retrySeconds": SYNC_INTERVAL_SECONDS}


class ProjectPulseHandler(SimpleHTTPRequestHandler):
    def end_headers(self):
        # The frontend is edited frequently during development. Prevent the
        # browser from reusing an older JS/CSS file after ProjectPulse is
        # restarted. The version query strings in index.html provide an
        # additional cache-busting layer.
        request_path = urlparse(self.path).path
        if request_path.endswith((".js", ".css", ".html")):
            self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
            self.send_header("Pragma", "no-cache")
            self.send_header("Expires", "0")
        # Helpful if the user accidentally runs the UI from VS Code Live Server
        # on another local port while this API is running on 8000.
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(204)
        self.end_headers()

    def send_json(self, status_code, payload):
        body = json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urlparse(self.path).path

        if path == "/api/tasks":
            try:
                payload = read_rows()
                downtime_state = read_downtime_state()
                payload["downtimes"] = downtime_state["downtimes"]
                payload["dependencyOverrides"] = downtime_state["dependencyOverrides"]
                payload["sync"] = pending_sync_status()
                self.send_json(200, payload)
            except Exception as exc:
                self.send_json(500, {"error": str(exc)})
            return

        if path == "/api/sync-status":
            self.send_json(200, pending_sync_status())
            return

        if path == "/api/downtimes":
            try:
                self.send_json(200, read_downtime_state())
            except Exception as exc:
                self.send_json(500, {"error": str(exc)})
            return

        super().do_GET()

    def do_POST(self):
        path = urlparse(self.path).path

        if path not in ("/api/save", "/api/reorder", "/api/downtimes"):
            self.send_json(404, {"error": "Not found"})
            return

        try:
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            if path == "/api/reorder":
                # Kept for backward compatibility. New frontend versions do
                # not use this endpoint for downtime persistence.
                result = save_ordered_tasks(payload.get("tasks", []))
            elif path == "/api/downtimes":
                state = {
                    "downtimes": payload.get("downtimes", []) if isinstance(payload.get("downtimes", []), list) else [],
                    "dependencyOverrides": payload.get("dependencyOverrides", []) if isinstance(payload.get("dependencyOverrides", []), list) else [],
                }
                _write_downtime_state(state)
                result = {"updated": True, "file": str(DOWNTIME_FILE)}
            else:
                result = save_tasks(payload.get("tasks", []))
            self.send_json(200, {"ok": True, **result})
        except Exception as exc:
            self.send_json(500, {"ok": False, "error": str(exc)})

    def log_message(self, format, *args):
        print(f"[HTTP] {self.address_string()} - {format % args}")


if __name__ == "__main__":
    os.chdir(BASE_DIR)
    server = ThreadingHTTPServer((HOST, PORT), ProjectPulseHandler)
    threading.Thread(target=_sync_worker, daemon=True, name="ProjectPulseExcelSync").start()
    print("=" * 60)
    print("ProjectPulse local server")
    print(f"Web app : http://{HOST}:{PORT}")
    print(f"Excel   : {EXCEL_FILE}")
    print("Press Ctrl+C to stop.")
    print("=" * 60)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping ProjectPulse server...")
    finally:
        server.server_close()
