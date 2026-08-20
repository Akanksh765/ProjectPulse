/*=========================================================
    Parse Raw Rows Into Project
==========================================================*/
// Shared by the local Excel-server path. This keeps the scheduler/Gantt
// pipeline independent from how the backend obtained the workbook rows.

function normalizeLoadedTaskIds() {
    // Excel IDs are the stable physical-row references, while ProjectPulse
    // displays continuous 1..N IDs.  Older versions could leave a duplicate
    // or skipped ID after inserting a downtime. Normalize strictly by row
    // order when loading so the Gantt and scheduler always see:
    //
    // 1, 2, 3, ... N
    //
    // Dependencies are remapped using the original Excel ID. When an old ID
    // occurs more than once, prefer the latest occurrence above the dependent
    // row because dependency chains in ProjectPulse point backwards.
    const tasks = Project.tasks;
    const oldIds = tasks.map(task => Number(task._excelOriginalId));

    let changed = false;

    const candidatesForOriginalId = (originalId, beforeIndex = tasks.length) => {
        const numericId = Number(originalId);
        const candidates = [];

        for (let i = 0; i < beforeIndex; i++) {
            if (Number(tasks[i]._excelOriginalId) === numericId) {
                candidates.push(tasks[i]);
            }
        }

        return candidates;
    };

    tasks.forEach((task, index) => {
        const newId = index + 1;
        if (Number(task.id) !== newId) {
            changed = true;
        }
        task.id = newId;
    });

    tasks.forEach((task, index) => {
        const originalDependencies = Array.isArray(task.dependencies)
            ? task.dependencies.map(Number)
            : [];

        const remappedDependencies = originalDependencies
            .map(oldDependencyId => {
                const previousMatches = candidatesForOriginalId(
                    oldDependencyId,
                    index
                );

                if (previousMatches.length > 0) {
                    return previousMatches[previousMatches.length - 1].id;
                }

                const anyMatch = candidatesForOriginalId(
                    oldDependencyId,
                    tasks.length
                );

                return anyMatch.length
                    ? anyMatch[anyMatch.length - 1].id
                    : null;
            })
            .filter(id => Number.isFinite(id));

        const uniqueDependencies = [...new Set(remappedDependencies)];

        if (
            JSON.stringify(uniqueDependencies) !==
            JSON.stringify(originalDependencies)
        ) {
            changed = true;
        }

        task.dependencies = uniqueDependencies;

        if (task.parentActivityId != null) {
            const parentMatches = candidatesForOriginalId(
                task.parentActivityId,
                index
            );

            const anyParentMatch = parentMatches.length
                ? parentMatches[parentMatches.length - 1]
                : candidatesForOriginalId(
                    task.parentActivityId,
                    tasks.length
                ).slice(-1)[0];

            if (anyParentMatch) {
                const newParentId = anyParentMatch.id;
                if (Number(task.parentActivityId) !== Number(newParentId)) {
                    changed = true;
                }
                task.parentActivityId = newParentId;
            }
        }
    });

    // oldIds is intentionally read before mutation; this detects gaps even
    // when the IDs happen to be unique.
    const expected = tasks.map((_, index) => index + 1);
    if (JSON.stringify(oldIds) !== JSON.stringify(expected)) {
        changed = true;
    }

    return changed;
}


function parseRowsIntoProject(rows) {

    Project.tasks = [];

    rows.forEach(row => {

        const activity = String(getField(row, ["Activity"]) || "").trim();
        const type = String(getField(row, ["Type", "Task Type"]) || "").trim();
        const reason = String(getField(row, ["Reason", "Downtime Reason"]) || "").trim();
        const parentId = getField(row, ["Parent Activity ID", "Parent ID"]);

        // Ignore the old placeholder row that existed in the original
        // workbook. A real downtime must be explicitly marked as Type =
        // Downtime (or named DT_#). This prevents one manually-created
        // DT_1 from being accompanied by the old sample "Downtime" row.
        if (
            activity.toLowerCase() === "downtime" &&
            !type &&
            !reason &&
            (parentId === "" || parentId == null)
        ) {
            console.warn("[Excel] Ignoring legacy placeholder Downtime row.");
            return;
        }

        const task = createTask(row);

        if (task)
            Project.tasks.push(task);

    });

    const idsWereNormalized = normalizeLoadedTaskIds();

    validateTasks();

    runScheduler();

    refreshProjectUI();

    // If an older workbook contained skipped/duplicate IDs, immediately
    // persist the normalized continuous order back to Excel.  The stable
    // _excelOriginalId values still identify the physical rows correctly.
    if (idsWereNormalized && typeof saveProjectToExcel === "function") {
        saveProjectToExcel();
    }

}


/*=========================================================
    Local Excel Auto-Load
==========================================================*/
// server.py serves /api/tasks from the Excel file configured there.
// ProjectPulse loads it automatically when the app starts, so the user
// no longer needs to select the workbook manually.

const PROJECTPULSE_API_BASE = (window.PROJECTPULSE_API_BASE || "http://127.0.0.1:8000").replace(/\/$/, "");

async function loadWorkbookFromServer() {

    try {

        const response = await fetch(PROJECTPULSE_API_BASE + "/api/tasks", {
            cache: "no-store"
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({}));
            throw new Error(error.error || `Server returned HTTP ${response.status}`);
        }

        const payload = await response.json();

        if (!Array.isArray(payload.rows)) {
            throw new Error("The local Excel server returned no task rows.");
        }

        Project.excel = {
            file: payload.file || "",
            sheet: payload.sheet || "",
            source: "local-server"
        };

        console.log("[Excel] Automatically loaded:", payload.file);
        console.log("[Excel] Sheet:", payload.sheet);

        parseRowsIntoProject(payload.rows);

        if (typeof applyLoadedDowntimes === "function") {
            applyLoadedDowntimes(payload.downtimes || [], payload.dependencyOverrides || []);
        }

        setExcelSyncStatus("Excel connected");
        startExcelSyncStatusPolling();

    } catch (error) {

        console.warn("[Excel] Automatic local load unavailable:", error.message);
        setExcelSyncStatus("Excel not connected");

    }

}


/*=========================================================
    Save Current Project To Local Excel
==========================================================*/

async function saveProjectToExcel(taskIds = null) {

    // Do nothing when the app is being opened directly from file:// or
    // from a server that does not expose the ProjectPulse API.
    if (window.location.protocol === "file:")
        return false;

    const selectedIds = Array.isArray(taskIds) && taskIds.length
        ? new Set(taskIds.map(id => Number(id)))
        : null;

    const tasksToSave = (selectedIds
        ? Project.tasks.filter(task => selectedIds.has(Number(task.id)))
        : Project.tasks).filter(task => !task.isDowntime);

    if (!tasksToSave.length)
        return true;

    try {

        const response = await fetch(PROJECTPULSE_API_BASE + "/api/save", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                // IMPORTANT: send only tasks that actually changed. This
                // prevents a stale browser snapshot from overwriting edits
                // made directly in Excel by another person.
                tasks: tasksToSave.map(task => ({
                    id: task.id,
                    originalId: task._excelOriginalId ?? task.id,
                    activity: task.activity,
                    startDate: task.startDate instanceof Date ? toApiDate(task.startDate) : null,
                    endDate: task.endDate instanceof Date ? toApiDate(task.endDate) : null,
                    duration: task.duration,
                    // Never write a virtual downtime ID into Excel.
                    dependencies: task._excelDependencies || [],
                    status: task.status || "Not Started",
                    isDowntime: false,
                    reason: task.reason || "",
                    parentActivityId: task.parentActivityId ?? null,
                    insertAfterId: task.insertAfterId ?? null
                }))
            })
        });

        const payload = await response.json().catch(() => ({}));

        if (!response.ok || payload.ok === false) {
            throw new Error(payload.error || `Server returned HTTP ${response.status}`);
        }

        if (payload.queued) {
            console.warn("[Excel] Workbook is locked; changes queued for automatic retry.");
            setExcelSyncStatus(`Excel open — ${payload.pendingRows || tasksToSave.length} change(s) queued`);
            return true;
        }

        console.log("[Excel] Saved changes:", payload.updatedRows, "row(s)");
        setExcelSyncStatus("Saved to Excel");
        return true;

    } catch (error) {

        console.error("[Excel] Could not save to local workbook:", error);
        setExcelSyncStatus("Excel save failed — retrying");
        return false;

    }

}

/*=========================================================
    Save Downtime State To Local JSON
==========================================================*/
// Used when a downtime is inserted. Inserting a downtime changes the
// visible IDs of every task below it, so saving only "changed IDs" is not
// safe. This endpoint uses each task's original Excel ID to update the
// existing row while keeping the worksheet in exactly the same order as
// Project.tasks.
async function saveDowntimeState() {

    if (window.location.protocol === "file:")
        return false;

    const downtimes = Project.tasks
        .filter(task => task.isDowntime)
        .map((task, index) => ({
            uid: task._downtimeUid,
            name: task.activity,
            reason: task.reason || "",
            startDate: task.startDate instanceof Date ? toApiDate(task.startDate) : null,
            duration: Number(task.duration) || 1,
            afterOriginalId: task._afterOriginalId ?? null,
            sequence: Number(task._downtimeSequence) || index
        }));

    const dependencyOverrides = Project.tasks
        .filter(task => !task.isDowntime && task._virtualDependencyUid)
        .map(task => ({
            taskOriginalId: task._excelOriginalId,
            downtimeUid: task._virtualDependencyUid
        }));

    try {
        const response = await fetch(PROJECTPULSE_API_BASE + "/api/downtimes", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ downtimes, dependencyOverrides })
        });

        const payload = await response.json().catch(() => ({}));
        if (!response.ok || payload.ok === false) {
            throw new Error(payload.error || `Server returned HTTP ${response.status}`);
        }

        setExcelSyncStatus("Schedule saved");
        return true;
    } catch (error) {
        console.error("[Downtime] Could not save downtime state:", error);
        setExcelSyncStatus("Downtime save failed");
        return false;
    }
}

/*=========================================================
    Rebuild Virtual Downtime Rows From JSON
=========================================================*/
function applyLoadedDowntimes(records, dependencyOverrides = []) {

    if (!Array.isArray(records) || !records.length)
        return;

    const normalTasks = Project.tasks.filter(task => !task.isDowntime);
    const byOriginalId = new Map(
        normalTasks.map(task => [Number(task._excelOriginalId), task])
    );

    const grouped = new Map();
    records.forEach((record, index) => {
        const anchor = Number(record.afterOriginalId);
        const key = Number.isFinite(anchor) ? anchor : "__append__";
        if (!grouped.has(key)) grouped.set(key, []);
        grouped.get(key).push({ ...record, _recordIndex: index });
    });

    grouped.forEach(list => {
        list.sort((a, b) => {
            const sa = Number(a.sequence);
            const sb = Number(b.sequence);
            return (Number.isFinite(sa) ? sa : 0) - (Number.isFinite(sb) ? sb : 0) ||
                a._recordIndex - b._recordIndex;
        });
    });

    const ordered = [];
    normalTasks.forEach(task => {
        ordered.push(task);
        const recordsForTask = grouped.get(Number(task._excelOriginalId)) || [];
        recordsForTask.forEach(record => {
            ordered.push(makeDowntimeTaskFromRecord(record, task._excelOriginalId));
        });
    });

    const appended = grouped.get("__append__") || [];
    appended.forEach(record => {
        ordered.push(makeDowntimeTaskFromRecord(record, null));
    });

    Project.tasks = ordered;
    renumberTasksContinuously();

    // Rebuild the virtual dependency chain after IDs are continuous.
    Project.tasks.forEach((task, index) => {
        if (!task.isDowntime) return;
        const previous = Project.tasks[index - 1] || null;
        task.dependencies = previous ? [previous.id] : [];
        task.parentActivityId = previous ? previous.id : null;
        task.insertAfterId = previous ? previous.id : null;
    });

    // By default the real task immediately after a downtime depends on that
    // downtime. Then restore any explicit virtual dependency choices saved by
    // the user.
    Project.tasks.forEach((task, index) => {
        if (task.isDowntime) return;
        const previous = Project.tasks[index - 1] || null;
        if (previous && previous.isDowntime) {
            task.dependencies = [previous.id];
            task._virtualDependencyUid = previous._downtimeUid;
        }
    });

    const overrides = new Map(
        (Array.isArray(dependencyOverrides) ? dependencyOverrides : [])
            .map(item => [Number(item.taskOriginalId), item.downtimeUid])
    );

    Project.tasks.forEach(task => {
        if (task.isDowntime || !overrides.has(Number(task._excelOriginalId))) return;
        const uid = overrides.get(Number(task._excelOriginalId));
        const downtime = Project.tasks.find(item => item.isDowntime && item._downtimeUid === uid);
        if (!downtime) return;
        task.dependencies = [downtime.id];
        task._virtualDependencyUid = uid;
    });

    runScheduler();
}

function makeDowntimeTaskFromRecord(record, afterOriginalId) {
    const uid = String(record.uid || `dt-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`);
    const start = parseExcelDate(record.startDate);
    const duration = Math.max(1, Math.round(Number(record.duration) || 1));

    return {
        id: 0,
        activity: String(record.name || "Downtime"),
        isDowntime: true,
        reason: String(record.reason || ""),
        parentActivityId: null,
        duration,
        startDate: start instanceof Date ? new Date(start) : new Date(),
        manualStartDate: start instanceof Date ? new Date(start) : new Date(),
        endDate: null,
        dependencies: [],
        owner: "",
        status: "Not Started",
        progress: 0,
        priority: "Medium",
        insertAfterId: null,
        _excelOriginalId: null,
        _afterOriginalId: afterOriginalId,
        _downtimeUid: uid,
        _downtimeSequence: Number(record.sequence) || 0,
        _excelDependencies: []
    };
}

/*=========================================================
    Excel Sync Status Polling
==========================================================*/
// The Python server retries queued changes every few seconds while Excel is
// locked. Polling here keeps the UI status accurate without changing the
// scheduler or Gantt behaviour.

let excelSyncStatusTimer = null;

async function refreshExcelSyncStatus() {
    if (window.location.protocol === "file:") return;

    try {
        const response = await fetch(PROJECTPULSE_API_BASE + "/api/sync-status", { cache: "no-store" });
        if (!response.ok) return;
        const status = await response.json();

        if (status.queued) {
            setExcelSyncStatus(`Excel open — ${status.pendingRows} change(s) queued`);
        } else {
            setExcelSyncStatus("Excel synchronized");
        }
    } catch (_) {
        // Keep the last useful status; the normal save/load error handling
        // will report a genuine server outage.
    }
}

function startExcelSyncStatusPolling() {
    if (excelSyncStatusTimer) return;
    refreshExcelSyncStatus();
    excelSyncStatusTimer = setInterval(refreshExcelSyncStatus, 3000);
}

function toApiDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

function refreshProjectUI() {

    if (typeof refreshCurrentPage === "function") {
        refreshCurrentPage();
    } else {
        updateDashboard();
        if (typeof drawGantt === "function")
            drawGantt();
    }

}

function setExcelSyncStatus(message) {

    const status = document.getElementById("excelSyncStatus");

    if (!status) return;

    status.textContent = message;
    status.title = message;

}


/*=========================================================
    Flexible Column Lookup
==========================================================*/
// BUGFIX: createTask() used to read row.Duration, row["Start Date"],
// etc. by an exact, hardcoded header name. Real-world sheets don't
// always match that exactly - e.g. this project's own sample sheet
// uses "Duration (Weeks)" instead of "Duration", which meant
// row.Duration was always undefined and every single task silently
// fell back to the 1-week default (regardless of what the sheet
// actually said). This looks a column up by trying each candidate
// name, ignoring case/spacing/punctuation, and finally falling back
// to a prefix match so "Duration (Weeks)" is still found under the
// candidate "Duration".

function normalizeHeader(header) {

    return header
        .toString()
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");

}

function getField(row, candidates) {

    const keys = Object.keys(row);
    const normalizedKeys = keys.map(normalizeHeader);

    for (const candidate of candidates) {

        const normCandidate = normalizeHeader(candidate);
        const index = normalizedKeys.indexOf(normCandidate);

        if (index !== -1)
            return row[keys[index]];

    }

    // Fallback: prefix match, so "Duration (Weeks)" is found via
    // the candidate "Duration", "Start Date " (trailing space) is
    // found via "Start Date", etc.
    for (const candidate of candidates) {

        const normCandidate = normalizeHeader(candidate);
        const index = normalizedKeys.findIndex(k => k.startsWith(normCandidate));

        if (index !== -1)
            return row[keys[index]];

    }

    return "";

}


/*=========================================================
    Create Task
==========================================================*/

function createTask(row) {

    const idValue = getField(row, ["ID"]);

    if (idValue === "" || idValue == null) {

        console.error(
            "[Excel] Row skipped: missing ID value.", row
        );

        return null;
    }

    const activity = getField(row, ["Activity"]) || "";
    const taskType = String(getField(row, ["Type", "Task Type"]) || "").trim().toLowerCase();
    // Backward compatibility: older ProjectPulse workbooks may already
    // contain a row literally named "Downtime" without the newer Type
    // column. Treat that row as a downtime so it is rendered red and
    // participates in the scheduler just like newly-created downtime rows.
    const normalizedActivity = String(activity).trim().toLowerCase();
    // Type is the authoritative marker for a downtime. For convenience,
    // rows named DT_1, DT_2, ... are also treated as downtime rows when
    // importing older workbooks that do not yet have the Type column.
    const isDowntime =
        taskType === "downtime" ||
        /^dt_\\d+$/i.test(String(activity).trim());
    const reason = getField(row, ["Reason", "Downtime Reason"]) || "";
    const parentActivityIdRaw = getField(row, ["Parent Activity ID", "Parent ID"]);
    const parentActivityId = parentActivityIdRaw === "" || parentActivityIdRaw == null
        ? null
        : Number(parentActivityIdRaw);

    if (!activity) {

        console.warn(
            `[Excel] Task ${idValue}: missing Activity name.`
        );

    }

    const mapped = mapProgressValue(
        getField(row, ["Progress", "Status"])
    );

    const parsedStartDate = parseExcelDate(
        getField(row, ["Start Date"])
    );

    return {

        id: Number(idValue),

        activity: activity,

        duration: parseDuration(
            getField(row, ["Duration", "Duration (Weeks)"])
        ),

        startDate: parsedStartDate,

        // Keeps the user's explicit Start Date separate from the
        // scheduler's calculated Start Date. A dependency can push the
        // actual start later, but the user's chosen date is still
        // remembered and remains editable from the Gantt page.
        manualStartDate: parsedStartDate instanceof Date && !isNaN(parsedStartDate)
            ? new Date(parsedStartDate)
            : null,

        endDate: null,

        dependencies: parseDependencies(
            getField(row, ["Depends On"])
        ),

        owner: getField(row, ["Owner"]) || "",

        status: mapped.status,

        progress: mapped.progress,

        priority: getField(row, ["Priority"]) || "Medium",

        // Downtime is a first-class scheduled task. Existing rows are
        // normal activities unless the workbook explicitly marks them
        // as Type = Downtime.
        isDowntime: isDowntime,
        reason: String(reason || ""),
        parentActivityId: Number.isFinite(parentActivityId) ? parentActivityId : null,

        // Stable reference to the row's original Excel ID. This survives
        // a continuous-ID renumbering when a downtime is inserted.
        _excelOriginalId: Number(idValue),

        // Dependencies stored in Excel always use stable Excel activity IDs.
        // The visible Gantt IDs may shift when virtual downtime rows are
        // inserted, so keep the workbook dependency graph separately.
        _excelDependencies: parseDependencies(
            getField(row, ["Depends On"])
        )

    };

}


/*=========================================================
    Parse Duration
==========================================================*/

function parseDuration(value) {

    if (value === "" || value == null)
        return 1;

    if (typeof value === "number")
        return value;

    const text = value.toString().trim().toLowerCase();

    const match = text.match(/\d+/);

    if (!match)
        return 1;

    return parseInt(match[0]);

}
/*=========================================================
        Excel Date Converter (Single Source of Truth)
=========================================================*/
// parseExcelDate() is the ONLY date parser in the project. Every
// place that needs to turn a raw Excel cell value into a JS Date
// must call this function - do not add a second implementation
// elsewhere (a duplicate used to live in utils.js; it's gone now).
// Normalizes: Excel serial numbers, JS Date objects, DD/MM/YYYY and
// D/M/YYYY strings, "DD-Mon-YYYY" strings, and ISO strings.

function parseExcelDate(value) {

    if (value === "" || value == null)
        return null;

    // Already a Date object
    if (value instanceof Date)
        return new Date(value);

    // Excel serial number
    if (typeof value === "number") {

        return new Date(
            Math.round((value - 25569) * 86400 * 1000)
        );

    }

    const text = value.toString().trim();

    // BUGFIX: numeric slash/dash dates like "22/06/2026" used to be
    // handed straight to `new Date(text)`, which parses them as
    // MM/DD/YYYY (US format) - NOT the DD/MM/YYYY format this app
    // displays dates in everywhere else. For a day > 12 that fails
    // outright (month 22 doesn't exist) and the task silently loses
    // its date; for a day <= 12 it's worse - it parses "successfully"
    // into the WRONG date with no error at all (e.g. "05/03/2026"
    // silently became 3 May instead of 5 Mar). Parse DD/MM/YYYY
    // explicitly instead of leaving it to the browser's ambiguous
    // built-in guess.
    let match = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);

    if (match) {

        const day = parseInt(match[1], 10);
        const month = parseInt(match[2], 10);
        const year = parseInt(match[3], 10);
        const parsed = new Date(year, month - 1, day);

        // getMonth() check catches e.g. "31/09/2026" (September only
        // has 30 days) - JS would otherwise silently roll that over
        // into October instead of flagging it as bad data.
        if (!isNaN(parsed.getTime()) && parsed.getMonth() === month - 1)
            return parsed;

    }

    // DD-Mon-YYYY / DD Mon YYYY (e.g. "06-Jan-2026", "06 Jan 2026")
    match = text.match(/^(\d{1,2})[\s\-]([A-Za-z]{3,})[\s\-](\d{4})$/);

    if (match) {

        const parsed = new Date(`${match[2]} ${match[1]}, ${match[3]}`);

        if (!isNaN(parsed.getTime()))
            return parsed;

    }

    // Fallback - covers ISO strings ("2026-01-06") and anything else
    // the browser's own parser genuinely handles unambiguously.
    const fallback = new Date(text);

    if (!isNaN(fallback.getTime()))
        return fallback;

    console.error("[Excel] Unable to parse date:", value);

    return null;

}


/*=========================================================
    Progress Mapping
==========================================================*/

function mapProgressValue(value) {

    const text = (value || "")
        .toString()
        .trim()
        .toLowerCase();

    switch (text) {

        case "completed":

            return {
                status: "Completed",
                progress: 100
            };

        case "in progress":

            return {
                status: "In Progress",
                progress: 50
            };

        case "yet to start":

        case "not started":

            return {
                status: "Not Started",
                progress: 0
            };

        default:

            return {
                status: "Not Started",
                progress: 0
            };

    }

}


/*=========================================================
    Dependency Parser
==========================================================*/

function parseDependencies(value) {

    if (
        value === "" ||
        value == null
    ) {
        return [];
    }

    const text = value.toString().trim();

    if (text === "" || text === "-")
        return [];

    // BUGFIX: this used to be `Number(value)` wrapped in a single-
    // element array, which only ever handled ONE dependency ID.
    // "Depends On" cells like "5,7" or "3, 4" became Number("5,7")
    // === NaN, so every multi-dependency task silently ended up with
    // NO dependencies at all. Split on commas and parse each ID
    // individually instead.

    const parts = text.split(",");
    const ids = [];
    const seen = new Set();

    parts.forEach(part => {

        const trimmed = part.trim();

        if (trimmed === "" || trimmed === "-")
            return;

        const id = Number(trimmed);

        if (isNaN(id)) {

            console.error(
                `[Excel] Malformed dependency value "${trimmed}" ` +
                `in "Depends On" = "${text}" - ignored.`
            );

            return;
        }

        if (seen.has(id)) {

            console.warn(
                `[Excel] Duplicate dependency ID ${id} in ` +
                `"Depends On" = "${text}" - ignored.`
            );

            return;
        }

        seen.add(id);
        ids.push(id);

    });

    return ids;

}


/*=========================================================
    Validate Tasks
=========================================================*/

function validateTasks() {

    // -----------------------------
    // Duplicate ID Detection
    // -----------------------------
    // Run first, across the whole set, since it needs to compare
    // tasks against each other rather than validate one in isolation.
    const idCounts = {};

    Project.tasks.forEach(task => {
        idCounts[task.id] = (idCounts[task.id] || 0) + 1;
    });

    Object.keys(idCounts).forEach(id => {

        if (idCounts[id] > 1) {

            console.error(
                `[Excel] Duplicate ID detected: ${id} appears ` +
                `${idCounts[id]} times. All matching rows were kept, ` +
                `but duplicate IDs will break ordering and dependency ` +
                `lookups - fix the source Excel.`
            );

        }

    });

    Project.tasks.forEach(task => {

        // -----------------------------
        // Validate Start Date
        // -----------------------------
        if (!(task.startDate instanceof Date) ||
            isNaN(task.startDate.getTime())) {

            console.error(
                `[Excel] Task ${task.id}: Invalid Start Date`,
                task.startDate
            );

            task.startDate = null;
            task.manualStartDate = null;
        } else if (!(task.manualStartDate instanceof Date) ||
                   isNaN(task.manualStartDate.getTime())) {

            task.manualStartDate = new Date(task.startDate);

        }

        // Scheduler will calculate this later
        task.endDate = null;

        // -----------------------------
        // Validate Duration
        // -----------------------------
        if (isNaN(task.duration) || task.duration <= 0) {

            console.warn(
                `[Excel] Task ${task.id}: Invalid Duration (${task.duration}). Using 1 week.`
            );

            task.duration = 1;
        }

        // -----------------------------
        // Validate Progress
        // -----------------------------
        task.progress = Number(task.progress);

        if (isNaN(task.progress))
            task.progress = 0;

        task.progress = Math.max(0, Math.min(100, task.progress));

        // -----------------------------
        // Self-Dependency Detection
        // -----------------------------
        // Stripped here (rather than left for the circular-dependency
        // detector) so it gets its own clear, specific message instead
        // of showing up as a generic "Task X -> Task X" cycle.
        task.dependencies = task.dependencies.filter(depId => {

            if (depId === task.id) {

                console.error(
                    `[Excel] Task ${task.id}: self-dependency ` +
                    `(depends on itself) - removed.`
                );

                return false;
            }

            return true;

        });

    });

}
