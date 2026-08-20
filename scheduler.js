/*=========================================================
                ProjectPulse V2
             Scheduling Engine V1
=========================================================*/

/*
    Features
    --------
    ✔ Multiple Dependencies
    ✔ Automatic Cascading
    ✔ Circular Dependency Detection
    ✔ Auto End Date Calculation
*/


function runScheduler() {

    console.log("[Scheduler] Scheduling Project...");

    // -----------------------------------------------------
    // Circular Dependency Detection
    // -----------------------------------------------------
    // BUGFIX: this used to alert() and `return` immediately, which
    // aborted scheduling for EVERY task the moment any circular
    // dependency existed anywhere in the sheet - one bad cycle
    // between two tasks used to freeze the whole project's dates.
    // Now: find every cycle, log each clearly, block ONLY the tasks
    // that are part of (or unreachable because of) a cycle, and keep
    // scheduling everything else normally.
    const cycles = findCircularDependencies();

    const cyclicTaskIds = new Set();

    if (cycles.length > 0) {

        cycles.forEach(cyclePath => {

            cyclePath.forEach(id => cyclicTaskIds.add(id));

            console.error(
                "[Scheduler] Circular dependency detected:\n" +
                "Task " + cyclePath.join(" \u2192 Task ")
            );

        });

        alert(
            "Circular dependency detected among " + cyclicTaskIds.size +
            " task(s). Check the console for details - those tasks (and " +
            "anything depending on them) were skipped; all other tasks " +
            "were scheduled normally."
        );

    }

    // -----------------------------------------------------
    // Invalid Dependency References
    // -----------------------------------------------------
    const invalidDependencies = validateDependencies();

    const blockedTaskIds = new Set(cyclicTaskIds);

    invalidDependencies.forEach(entry => {
        blockedTaskIds.add(entry.task.id);
    });

    if (invalidDependencies.length > 0) {

        invalidDependencies.forEach(entry => {

            console.error(
                `[Scheduler] Invalid dependency: Task ${entry.task.id} ("${entry.task.activity}") ` +
                `depends on Task ID ${entry.invalidId}, which does not exist. ` +
                `This task will not be scheduled until the dependency is corrected.`
            );

        });

        alert(
            invalidDependencies.length + " task(s) reference a dependency ID " +
            "that does not exist. Check the console for details - those tasks " +
            "were skipped and kept their previous dates."
        );

    }

    // -----------------------------------------------------
    // Scheduling
    // -----------------------------------------------------
    // computeSchedulingOrder() returns a SEPARATE array used only to
    // decide the order tasks are processed in (so a dependency's
    // dates are always computed before its successor's). It does NOT
    // touch Project.tasks itself - the master collection stays in
    // Excel ID order for display, exactly as uploaded. See the note
    // above computeSchedulingOrder() for why this used to be broken.
    const schedulingOrder = computeSchedulingOrder(blockedTaskIds);

    scheduleTasks(schedulingOrder, blockedTaskIds);

    calculateProjectDates();

    console.log("[Scheduler] Scheduling Complete");

}



/*=========================================================

        Dependency Validation

=========================================================*/

// Confirms every dependency ID referenced by any task actually
// matches an existing task. Returns a list of { task, invalidId }
// entries - one per bad reference - without modifying anything.
// A task with even one invalid dependency ID is reported in full;
// it's up to the caller (runScheduler) to decide what to do about it.
function validateDependencies() {

    const validIds = new Set(
        Project.tasks.map(task => task.id)
    );

    const invalid = [];

    Project.tasks.forEach(task => {

        task.dependencies.forEach(depId => {

            if (!validIds.has(depId)) {

                invalid.push({ task: task, invalidId: depId });

            }

        });

    });

    return invalid;

}



/*=========================================================

        Scheduling Order (Internal Only)

=========================================================*/
// IMPORTANT: this is intentionally separate from Project.tasks.
// Project.tasks is the single master collection every page
// (Dashboard, Tasks, Gantt, Analytics, Reports) reads from, and its
// order must always stay exactly as read from the Excel "ID" column
// (Requirement: display order is ALWAYS 1, 2, 3 ... 21). Scheduling,
// however, needs a dependency-respecting (topological) order so a
// task's start date is only computed after every task it depends on
// already has its own end date. This used to be done by literally
// reassigning Project.tasks = sorted, which is exactly why the UI
// used to show tasks in an order like 1, 6, 2, 7, 3, 8 ... instead
// of 1-21: the scheduler's internal processing order and the
// display order were the same array. They're now fully independent.
//
// Any task id passed in `excludeIds` (invalid-dependency or
// circular-dependency tasks) is left out of the returned order
// entirely - scheduleTasks() skips them anyway, and Kahn's algorithm
// can't produce a valid position for a cyclic task in the first
// place.

function computeSchedulingOrder(excludeIds) {

    excludeIds = excludeIds || new Set();

    const taskMap = {};
    const inDegree = {};
    const adjList = {};

    const schedulable = Project.tasks.filter(
        task => !excludeIds.has(task.id)
    );

    schedulable.forEach(task => {
        taskMap[task.id] = task;
        inDegree[task.id] = 0;
        adjList[task.id] = [];
    });

    schedulable.forEach(task => {
        task.dependencies.forEach(depId => {
            if (taskMap[depId]) {
                adjList[depId].push(task.id);
                inDegree[task.id]++;
            }
        });
    });

    // Deterministic tie-breaking: among tasks that are equally ready
    // to schedule, process them in ID order rather than whatever
    // order Object.keys()/array iteration happens to produce.
    let queue = schedulable
        .filter(task => inDegree[task.id] === 0)
        .map(task => task.id)
        .sort((a, b) => a - b);

    const sorted = [];

    while (queue.length > 0) {

        const id = queue.shift();
        sorted.push(taskMap[id]);

        const ready = [];

        adjList[id].forEach(succId => {
            inDegree[succId]--;
            if (inDegree[succId] === 0) ready.push(succId);
        });

        if (ready.length > 0) {
            queue = queue.concat(ready).sort((a, b) => a - b);
        }

    }

    // Any task that never reached inDegree 0 is part of (or
    // downstream of) a cycle that slipped through the exclude list -
    // leave it out rather than let it corrupt the schedule order.
    if (sorted.length !== schedulable.length) {

        console.warn(
            "[Scheduler] " + (schedulable.length - sorted.length) +
            " task(s) could not be placed in a dependency order and " +
            "were left unscheduled."
        );

    }

    return sorted;

}



/*=========================================================

            Main Scheduling Logic

=========================================================*/

function scheduleTasks(schedulingOrder, blockedTaskIds) {

    schedulingOrder = schedulingOrder || Project.tasks;
    blockedTaskIds = blockedTaskIds || new Set();

    schedulingOrder.forEach(task => {

        if (blockedTaskIds.has(task.id))
            return;

        // ----------------------------
        // Independent Task
        // ----------------------------
        // An independent task has nothing else to derive its timing
        // from, so IT needs a valid Start Date of its own - that's
        // the only case where the validity check makes sense.

        if (task.dependencies.length === 0) {

            const manualStart = task.manualStartDate instanceof Date &&
                !isNaN(task.manualStartDate.getTime())
                ? task.manualStartDate
                : task.startDate;

            if (!(manualStart instanceof Date) ||
                isNaN(manualStart.getTime())) {

                console.error(
                    `[Scheduler] Task ${task.id} has an invalid start date.`,
                    manualStart
                );

                return;
            }

            task.startDate = new Date(manualStart);
            task.endDate = addDays(
                new Date(task.startDate),
                task.duration * 7
            );

            console.log("[Scheduler] Task",task.id,"scheduled:",task.startDate,"->",task.endDate);

            return;

        }

        // ----------------------------
        // Dependent Task
        // ----------------------------
        // BUGFIX: this used to fall through the same "does this task
        // have a valid Start Date" check as an independent task -
        // but a dependent task's start date is ALWAYS overwritten
        // below from its dependency's end date, so whatever (possibly
        // blank or unparseable) value happened to be sitting in its
        // own Start Date cell is irrelevant. Gating on it here meant
        // one task with a messy raw date - e.g. a date typed as text
        // that failed to parse - would abort before ever reaching the
        // dependency calculation, leaving it (and, since later tasks
        // depend on this one, everything downstream of it) permanently
        // unscheduled. A whole dependency chain could silently break
        // at whichever row happened to have the messiest source data.

        const latest = latestDependencyEnd(task);

        if (!(latest instanceof Date) ||
            isNaN(latest.getTime())) {

            console.warn(
                `[Scheduler] Task ${task.id}: waiting for dependency.`
            );

            return;
        }

        const dependencyStart = addDays(
            new Date(latest),
            1
        );

        const manualStart = task.manualStartDate instanceof Date &&
            !isNaN(task.manualStartDate.getTime())
            ? new Date(task.manualStartDate)
            : null;

        // Dependency is the earliest legal start. The user's editable
        // Start Date remains a valid later/manual start, so changing the
        // date in the Gantt page still has an effect without allowing a
        // task to start before its dependency is complete.
        task.startDate = manualStart && manualStart > dependencyStart
            ? manualStart
            : dependencyStart;

        task.endDate = addDays(
            new Date(task.startDate),
            task.duration * 7
        );
        console.log("[Scheduler] Task:",
    task.id,
    "Start:",
    task.startDate,
    "End:",
    task.endDate);

    });

    console.table(Project.tasks);

}

/*=========================================================

            Project Start

=========================================================*/

function projectStartDate() {

    let earliest = null;

    Project.tasks.forEach(task => {

        // BUGFIX: skip tasks that never got a valid startDate.
        // Comparing `someDate < null` coerces null to 0 (epoch),
        // which is why an unscheduled task used to silently overrule
        // every real date and could pin "earliest" to null forever.
        if (!(task.startDate instanceof Date) || isNaN(task.startDate))
            return;

        if (

            earliest == null ||

            task.startDate < earliest

        ) {

            earliest = task.startDate;

        }

    });

    return earliest;

}



/*=========================================================

            Project End

=========================================================*/

function projectEndDate() {

    let latest = null;

    Project.tasks.forEach(task => {

        // BUGFIX: skip tasks that never got a valid endDate (e.g.
        // blocked by an invalid "Depends On" reference). Without
        // this, task.endDate could be null, and while a lone null
        // wouldn't win a `>` comparison, it's the same class of bug
        // as projectStartDate() above - be explicit rather than rely
        // on comparison coercion.
        if (!(task.endDate instanceof Date) || isNaN(task.endDate))
            return;

        if (

            latest == null ||

            task.endDate > latest

        ) {

            latest = task.endDate;

        }

    });

    return latest;

}



/*=========================================================

            Project Duration

=========================================================*/

function calculateProjectDates() {

    Project.projectStart =

        projectStartDate();

    Project.projectEnd =

        projectEndDate();

}



/*=========================================================

        Circular Dependency Detection

=========================================================*/

function findCircularDependencies() {

    const visited = {};

    const stack = [];        // current DFS path, as task IDs
    const onStack = {};

    const map = {};

    Project.tasks.forEach(task => {

        map[task.id] = task;

    });

    const cycles = [];

    function dfs(id) {

        if (onStack[id]) {

            // Back-edge into the current path: the slice from where
            // `id` first appears in the path, through to now, plus
            // `id` again to close the loop, IS the cycle - e.g.
            // stack = [6, 7], id = 6 -> cyclePath = [6, 7, 6].
            const startIndex = stack.indexOf(id);
            const cyclePath = stack.slice(startIndex).concat(id);

            cycles.push(cyclePath);

            return;
        }

        if (visited[id])
            return;

        visited[id] = true;
        onStack[id] = true;
        stack.push(id);

        const task = map[id];

        if (task) {

            task.dependencies.forEach(depId => {

                // Only recurse into IDs that actually exist - an
                // invalid dependency reference is a separate problem
                // (handled by validateDependencies()), not a cycle.
                if (map[depId]) {
                    dfs(depId);
                }

            });

        }

        stack.pop();
        onStack[id] = false;

    }

    Project.tasks.forEach(task => dfs(task.id));

    return cycles;

}
/*=========================================================
    Excel Sync Change Tracking
=========================================================*/
// Only persist tasks whose relevant Excel-backed fields changed during a
// scheduler operation. This is important for shared editing: a browser may
// have an older snapshot than the workbook currently open in Excel.

function captureExcelSyncState() {
    return new Map(Project.tasks.map(task => [Number(task.id), JSON.stringify({
        activity: task.activity,
        startDate: task.startDate instanceof Date ? task.startDate.getTime() : task.startDate,
        endDate: task.endDate instanceof Date ? task.endDate.getTime() : task.endDate,
        duration: task.duration,
        dependencies: task.dependencies || [],
        status: task.status || "Not Started",
        isDowntime: !!task.isDowntime,
        reason: task.reason || "",
        parentActivityId: task.parentActivityId ?? null
    })]));
}

function changedExcelTaskIds(beforeState) {
    return Project.tasks
        .filter(task => {
            const current = JSON.stringify({
                activity: task.activity,
                startDate: task.startDate instanceof Date ? task.startDate.getTime() : task.startDate,
                endDate: task.endDate instanceof Date ? task.endDate.getTime() : task.endDate,
                duration: task.duration,
                dependencies: task.dependencies || [],
                status: task.status || "Not Started",
                isDowntime: !!task.isDowntime,
                reason: task.reason || "",
                parentActivityId: task.parentActivityId ?? null
            });
            return beforeState.get(Number(task.id)) !== current;
        })
        .map(task => task.id);
}

/*=========================================================

        Add Downtime

=========================================================*/
// Downtime is a first-class scheduled task. It is inserted immediately
// after the selected activity (or after that activity's existing
// downtime chain), receives its own task ID, and becomes part of the
// dependency graph. The next row is made dependent on the new downtime
// so the scheduler naturally pushes it and anything downstream.

function addDowntimeAfterActivity(parentTaskID, name, reason, startDate, duration) {

    const beforeState = captureExcelSyncState();
    const parent = findTask(parentTaskID);

    if (!parent) {
        console.error(`[Scheduler] Cannot add downtime: Activity ${parentTaskID} was not found.`);
        return null;
    }

    const cleanName = String(name || "").trim();
    const cleanReason = String(reason || "").trim();

    if (!cleanName || !cleanReason) {
        console.error("[Scheduler] Downtime name and reason are required.");
        return null;
    }

    const parsedDuration = Math.max(1, Math.round(Number(duration) || 1));
    const parsedStart = startDate instanceof Date && !isNaN(startDate.getTime())
        ? new Date(startDate)
        : (parent.endDate instanceof Date && !isNaN(parent.endDate.getTime())
            ? addDays(parent.endDate, 1)
            : null);

    if (!parsedStart) {
        console.error(`[Scheduler] Cannot add downtime after Activity ${parentTaskID}: no valid start date.`);
        return null;
    }

    const parentIndex = Project.tasks.indexOf(parent);
    let insertIndex = parentIndex + 1;
    let dependencyAnchor = parent;

    // If the selected activity already has downtime rows directly below it,
    // place the new downtime after the existing downtime chain.
    while (
        insertIndex < Project.tasks.length &&
        Project.tasks[insertIndex].isDowntime &&
        Number(Project.tasks[insertIndex].parentActivityId) === Number(parent.id)
    ) {
        dependencyAnchor = Project.tasks[insertIndex];
        insertIndex++;
    }

    const nextTask = Project.tasks[insertIndex] || null;

    // Keep the physical Excel row reference separate from the visible ID.
    // This is what allows a downtime to be deleted even after a server restart.
    const anchorOriginalId =
        dependencyAnchor._excelOriginalId ?? dependencyAnchor.id;

    // Give the new object a temporary unique numeric ID BEFORE inserting it.
    // The previous implementation used null here, which caused the following
    // activity to receive dependency 0 and left the downtime without an ID.
    const maxExistingId = Project.tasks.reduce(
        (max, task) => Math.max(max, Number(task.id) || 0),
        0
    );

    const downtime = {
        id: maxExistingId + 1,
        activity: cleanName,
        isDowntime: true,
        reason: cleanReason,
        parentActivityId: parent.id,
        duration: parsedDuration,
        startDate: new Date(parsedStart),
        manualStartDate: new Date(parsedStart),
        endDate: null,
        dependencies: [dependencyAnchor.id],
        owner: "",
        status: "Not Started",
        progress: 0,
        priority: "Medium",
        insertAfterId: dependencyAnchor.id,
        _excelOriginalId: null,
        _insertAfterOriginalId: anchorOriginalId,
        _afterOriginalId: parent._excelOriginalId ?? parent.id,
        _downtimeUid: `dt-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`,
        _downtimeSequence: Project.tasks.filter(task =>
            task.isDowntime && Number(task._afterOriginalId) === Number(parent._excelOriginalId ?? parent.id)
        ).length
    };

    Project.tasks.splice(insertIndex, 0, downtime);

    // Renumber immediately. This remaps ALL existing dependency references
    // and guarantees that the new downtime gets the ID represented by its
    // actual row position:
    //
    // 12 Activity
    // 13 Downtime
    // 14 Activity
    //
    renumberTasksContinuously();

    const downtimeIndex = Project.tasks.indexOf(downtime);
    const previousTask = Project.tasks[downtimeIndex - 1] || null;
    const followingTask = Project.tasks[downtimeIndex + 1] || null;

    // The downtime always depends on the row immediately above it.
    downtime.dependencies = previousTask ? [previousTask.id] : [];

    // The activity immediately below the downtime MUST depend on the downtime.
    // Replace the direct dependency on the previous activity, but preserve
    // any unrelated dependencies that may already exist.
    if (followingTask) {
        const existing = Array.isArray(followingTask.dependencies)
            ? followingTask.dependencies.map(Number)
            : [];

        const withoutPrevious = previousTask
            ? existing.filter(id => Number(id) !== Number(previousTask.id))
            : existing;

        followingTask.dependencies = [
            ...new Set([...withoutPrevious, Number(downtime.id)])
        ];
        if (!followingTask.isDowntime) {
            followingTask._virtualDependencyUid = downtime._downtimeUid;
        }
    }

    runScheduler();

    updateDashboard();

    if (typeof drawGantt === "function") {
        drawGantt();
    }

    if (typeof saveProjectToExcel === "function") {
        saveProjectToExcel(changedExcelTaskIds(beforeState));
    }
    if (typeof saveDowntimeState === "function") {
        saveDowntimeState();
    }

    return downtime;

}


/*=========================================================
    Remove Downtime
=========================================================*/
// Downtime deletion is a full ordered-project operation.  It deliberately
// does NOT use the small /api/save patch endpoint because removing a row
// requires the backend to delete the corresponding physical Excel row.
//
// The stable _excelOriginalId survives a server restart, so this works for
// downtimes loaded from Excel as well as ones created during the current
// browser session.

function removeDowntime(taskID) {

    const beforeState = captureExcelSyncState();
    const downtime = findTask(taskID);

    if (!downtime || !downtime.isDowntime) {
        console.error(`[Scheduler] Cannot remove downtime ${taskID}: downtime not found.`);
        return false;
    }

    const index = Project.tasks.indexOf(downtime);
    if (index < 0) return false;

    const previousTask = Project.tasks[index - 1] || null;
    const followingTask = Project.tasks[index + 1] || null;

    if (!confirm(`Remove downtime "${downtime.activity}" (ID ${downtime.id})?`)) {
        return false;
    }

    // Remove the downtime from the in-memory project first.
    Project.tasks.splice(index, 1);

    // Restore the dependency chain.  If the task below the downtime was
    // dependent on the downtime, it now depends on the activity immediately
    // above the removed downtime. Other independent dependencies are kept.
    if (followingTask) {
        const existing = Array.isArray(followingTask.dependencies)
            ? followingTask.dependencies.map(Number)
            : [];

        const withoutDowntime = existing.filter(
            id => Number(id) !== Number(downtime.id)
        );

        if (previousTask) {
            followingTask.dependencies = [
                ...new Set([
                    ...withoutDowntime,
                    Number(previousTask.id)
                ])
            ];
            if (previousTask.isDowntime) {
                followingTask._virtualDependencyUid = previousTask._downtimeUid;
            } else {
                delete followingTask._virtualDependencyUid;
                followingTask._excelDependencies = [Number(previousTask._excelOriginalId)];
            }
        } else {
            followingTask.dependencies = withoutDowntime;
            delete followingTask._virtualDependencyUid;
        }
    }

    // Rebuild continuous IDs and remap every dependency reference.
    renumberTasksContinuously();

    runScheduler();

    updateDashboard();

    if (typeof drawGantt === "function") {
        drawGantt();
    }

    if (typeof saveProjectToExcel === "function") {
        saveProjectToExcel(changedExcelTaskIds(beforeState));
    }
    if (typeof saveDowntimeState === "function") {
        saveDowntimeState();
    }

    return true;
}


/*=========================================================
    Continuous Task IDs
==========================================================*/
// IDs are presentation/order IDs as well as the dependency IDs used by the
// scheduler. Whenever a downtime is inserted, every row is renumbered in
// its visible order: 1, 2, 3, ... N. Dependencies are remapped at the same
// time, so inserting a downtime after Activity 12 gives the downtime ID 13
// and the old Activity 13 becomes 14.

function renumberTasksContinuously() {

    const oldIdByTask = new Map();

    Project.tasks.forEach(task => {
        oldIdByTask.set(task, Number(task.id));
    });

    const oldToNew = new Map();

    Project.tasks.forEach((task, index) => {
        const oldId = oldIdByTask.get(task);
        const newId = index + 1;
        if (Number.isFinite(oldId)) {
            oldToNew.set(oldId, newId);
        }
    });

    Project.tasks.forEach((task, index) => {

        task.id = index + 1;

        if (Array.isArray(task.dependencies)) {
            task.dependencies = task.dependencies
                .map(depId => oldToNew.get(Number(depId)))
                .filter(id => Number.isFinite(id));
        }

        if (task.parentActivityId != null) {
            const mappedParent = oldToNew.get(Number(task.parentActivityId));
            task.parentActivityId =
                Number.isFinite(mappedParent) ? mappedParent : task.parentActivityId;
        }

        if (task.insertAfterId != null) {
            const mappedAnchor = oldToNew.get(Number(task.insertAfterId));
            task.insertAfterId =
                Number.isFinite(mappedAnchor) ? mappedAnchor : task.insertAfterId;
        }

        // Remove temporary insertion bookkeeping after the Excel save
        // function has consumed it; the original Excel ID remains until save.
    });

    // The newly-created downtime's dependency is its preceding task.
    Project.tasks.forEach((task, index) => {
        if (!task.isDowntime) return;

        const previous = Project.tasks[index - 1];

        if (previous) {
            task.dependencies = [previous.id];
            task.parentActivityId = task.parentActivityId || previous.id;
            task.insertAfterId = previous.id;
        }
    });

    return oldToNew;
}



/*=========================================================

        Update Duration

=========================================================*/

function updateDuration(taskID, duration) {

    const beforeState = captureExcelSyncState();
    const task = findTask(taskID);

    if (!task)

        return;



    task.duration = duration;



    runScheduler();



    updateDashboard();



    if (

        typeof drawGantt ===

        "function"

    ) {

        drawGantt();

    }

    if (typeof saveProjectToExcel === "function") {
        saveProjectToExcel(changedExcelTaskIds(beforeState));
    }

}



/*=========================================================

        Update Start Date (from the Gantt page's editable
        Start Date field)

=========================================================*/
// Sets an explicit new Start Date on a task and lets the normal
// scheduler run take it from there. Mirrors updateDuration() above:
// this only makes sense for an independent task (no dependencies) -
// a dependent task's Start Date is always derived from its
// dependency's End Date inside scheduleTasks(), so runScheduler()
// will simply overwrite it again on the very next pass. That is
// expected: dependency wins over a manually typed date, exactly
// like requirement 6/8 describe.

function updateStartDate(taskID, newStartDate) {

    const beforeState = captureExcelSyncState();
    const task = findTask(taskID);

    if (!task)
        return;

    if (!(newStartDate instanceof Date) || isNaN(newStartDate.getTime())) {
        console.error(`[Scheduler] Task ${taskID}: invalid Start Date value.`, newStartDate);
        return;
    }

    task.manualStartDate = new Date(newStartDate);
    task.startDate = new Date(newStartDate);

    runScheduler();

    updateDashboard();

    if (typeof drawGantt === "function") {
        drawGantt();
    }

    if (typeof saveProjectToExcel === "function") {
        saveProjectToExcel(changedExcelTaskIds(beforeState));
    }

}



/*=========================================================

        Update Dependency (from the Gantt page's editable
        Dependency dropdown)

=========================================================*/
// Replaces a task's entire dependency list with either a single
// dependency ID or none. The underlying data model supports
// multiple dependencies per task (see parseDependencies() in
// excel.js), but the editable dropdown in the UI only ever offers
// one dependency at a time - matching the mock in the spec - so this
// intentionally collapses task.dependencies down to a single-item
// (or empty) array rather than trying to merge/append.

function updateDependency(taskID, newDependencyID) {

    const beforeState = captureExcelSyncState();
    const task = findTask(taskID);

    if (!task)
        return;

    if (newDependencyID === null || newDependencyID === "" || typeof newDependencyID === "undefined") {
        task.dependencies = [];
        delete task._virtualDependencyUid;
        task._excelDependencies = [];
    } else {
        const dependency = findTask(Number(newDependencyID));
        task.dependencies = [Number(newDependencyID)];

        if (dependency && dependency.isDowntime) {
            task._virtualDependencyUid = dependency._downtimeUid;
            // Keep the last real Excel dependency untouched. The downtime is
            // virtual and must never be written into the workbook.
        } else if (dependency) {
            delete task._virtualDependencyUid;
            task._excelDependencies = [Number(dependency._excelOriginalId)];
        } else {
            delete task._virtualDependencyUid;
        }
    }

    runScheduler();

    updateDashboard();

    if (typeof drawGantt === "function") {
        drawGantt();
    }

    if (typeof saveProjectToExcel === "function") {
        saveProjectToExcel(changedExcelTaskIds(beforeState));
    }

}



/*=========================================================

        Update Status / Progress

=========================================================*/

function updateStatus(taskID, newStatus) {

    const beforeState = captureExcelSyncState();
    const task = findTask(taskID);

    if (!task)
        return;

    const allowed = ["Completed", "In Progress", "Not Started", "Delayed"];
    if (!allowed.includes(newStatus))
        newStatus = "Not Started";

    task.status = newStatus;

    if (newStatus === "Completed") {
        task.progress = 100;
    } else if (newStatus === "Not Started") {
        task.progress = 0;
    } else if (newStatus === "In Progress") {
        const current = Number(task.progress);
        task.progress = (isNaN(current) || current <= 0 || current >= 100) ? 50 : current;
    } else if (newStatus === "Delayed") {
        const current = Number(task.progress);
        task.progress = (isNaN(current) || current >= 100) ? 50 : Math.max(0, current);
    }

    updateDashboard();

    if (typeof drawGantt === "function") {
        drawGantt();
    }

    if (typeof saveProjectToExcel === "function") {
        saveProjectToExcel(changedExcelTaskIds(beforeState));
    }

}



