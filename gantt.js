function drawCustomGantt() {

    const container = document.getElementById("ganttContainer") || document.getElementById("gantt");

    if (!container) {
        console.warn("Gantt container not found (#ganttContainer or #gantt)");
        return;
    }

    // ----------------------------------------------------
    // Mouse-wheel horizontal scrolling
    // ----------------------------------------------------
    // The Gantt timeline is intentionally wider than the viewport.
    // Keep the existing scrollbar, but also let the mouse wheel
    // move the timeline horizontally when the pointer is over the
    // Gantt scroll container. The handler is installed only once
    // because drawCustomGantt() can be called repeatedly after Excel
    // uploads or schedule changes.
    if (!container.__ganttWheelScrollAttached) {
        container.addEventListener("wheel", function (event) {
            // Do not interfere with browser/page zoom gestures.
            if (event.ctrlKey) return;

            const maxScrollLeft = container.scrollWidth - container.clientWidth;

            // If there is no horizontal overflow, leave normal wheel
            // behavior untouched.
            if (maxScrollLeft <= 0) return;

            // Normal wheel movement becomes horizontal movement.
            // Shift+wheel is also supported naturally by using deltaX
            // when the browser provides it.
            let delta = event.deltaY;
            if (event.shiftKey && Math.abs(event.deltaX) > Math.abs(event.deltaY)) {
                delta = event.deltaX;
            }

            if (delta === 0) return;

            const previousScrollLeft = container.scrollLeft;
            container.scrollLeft += delta;

            // Prevent the page from moving vertically while the user
            // is deliberately navigating the wide Gantt timeline.
            // This also keeps the page from taking over when the timeline
            // reaches its left or right edge.
            if (container.scrollLeft !== previousScrollLeft || maxScrollLeft > 0) {
                event.preventDefault();
            }
        }, { passive: false });

        container.__ganttWheelScrollAttached = true;
    }

    container.innerHTML = "";

    if (Project.tasks.length === 0) return;

    //----------------------------------------------------
    // Find Project Start and End
    //----------------------------------------------------
    // BUGFIX: new Date(null) does NOT produce an Invalid Date - it
    // silently resolves to the Unix epoch (1 Jan 1970). Any task
    // still sitting on a null startDate/endDate (unparsed Excel
    // date, or a task the scheduler had to leave unscheduled
    // because of a bad "Depends On" reference) used to sail straight
    // through the isNaN() check as a "valid" 1970 date. That dragged
    // minDate back 50+ years, so totalWeeks below exploded into the
    // thousands and every real bar ended up scrolled off-screen -
    // this was the actual cause of "no bars appear". We now treat a
    // task as having no usable date unless BOTH startDate and
    // endDate are actual Date objects.

    function hasUsableDates(task) {
        return task.startDate instanceof Date && !isNaN(task.startDate) &&
               task.endDate instanceof Date && !isNaN(task.endDate);
    }

    const datedTasks = Project.tasks.filter(hasUsableDates);

    if (datedTasks.length === 0) {
        console.warn("No tasks with valid start/end dates - nothing to draw.");
        return;
    }

    let minDate = new Date(datedTasks[0].startDate);
    let maxDate = new Date(datedTasks[0].endDate);

    datedTasks.forEach(task => {

        const startDate = new Date(task.startDate);
        const endDate = new Date(task.endDate);

        if (startDate < minDate)
            minDate = startDate;

        if (endDate > maxDate)
            maxDate = endDate;

    });

    //----------------------------------------------------
    // Weekly Timeline Anchor
    //----------------------------------------------------
    // The timeline is drawn in whole-week columns (no daily
    // sub-columns). Every week column still starts on the Monday of
    // the week that contains it, so the very first column starts on
    // the Monday on/before the earliest task start date (the
    // "project week start"). dayWidth is kept as the underlying unit
    // for all date-to-pixel math (bar position/width, today marker) -
    // this does NOT change any scheduled date, it only decides where
    // the week columns are drawn from and how a bar is positioned
    // within its starting week.

    const dayWidth = 3.5;
    const weekWidth = dayWidth * 7;
    const weekStart = mondayOnOrBefore(minDate);

    // How many week columns the grid needs, from weekStart up to the
    // latest task end date. maxDate is exclusive (see the note by
    // weekSpan below), so this is rounded UP to whole weeks - the
    // same rule used for each task's own bar width - to avoid an
    // extra, empty trailing week column.
    const totalWeeks = Math.max(
        1,
        Math.ceil(differenceInDays(weekStart, maxDate) / 7)
    );

    // ----------------------------------------------------
    // Safety Guard - bad/garbage dates can still slip through
    // as "valid" Date objects (e.g. a mis-parsed Excel serial
    // landing decades away) and blow totalWeeks up into the
    // thousands. Rendering that many columns would hang the
    // browser and look like a "crash" once a project has more
    // than a handful of tasks. Instead of freezing, bail out
    // with a clear on-screen message so the real problem (a bad
    // date somewhere in the sheet) is easy to spot.
    // ----------------------------------------------------
    const MAX_RENDERABLE_WEEKS = 260; // ~5 years - generous for any real project

    if (totalWeeks > MAX_RENDERABLE_WEEKS) {
        container.innerHTML = `
            <div class="ganttError">
                <i class="fa-solid fa-triangle-exclamation"></i>
                This project's timeline spans ${totalWeeks} weeks, which looks like
                a data problem (check for a task with a bad Start Date or Duration)
                rather than a real schedule. Fix the dates in the sheet and
                refresh the connected workbook to see the chart.
            </div>
        `;
        console.error(
            `Gantt aborted: totalWeeks (${totalWeeks}) exceeds the sane limit ` +
            `of ${MAX_RENDERABLE_WEEKS}. Check task start/end dates.`
        );
        return;
    }

    //----------------------------------------------------
    // Main Grid
    //----------------------------------------------------

    const table = document.createElement("table");
    table.className = "ganttTable";

    // IMPORTANT: Keep every timeline week at a real, fixed pixel width.
    // The previous table used only min-width on <th>/<td>, which lets the
    // browser shrink columns when the table is constrained by its
    // container. The task bars are calculated in 40px/day, so a shrinking
    // grid makes the bars appear to stop before their actual End Date.
    // Lock the table to the same pixel geometry used by the bar renderer.
    //
    // Left of the timeline there are now THREE logical sections living
    // in the same table (so their rows can never drift out of vertical
    // alignment with each other): Activity, Task Details (Start /
    // Duration / Dependency / Progress), then the weekly Gantt grid.
    const taskIdColumnWidth = 42;
    const taskColumnWidth = 330;
    const detailColumnWidths = { start: 82, duration: 72, dependency: 78, status: 82 };
    const detailsTotalWidth = detailColumnWidths.start + detailColumnWidths.duration +
        detailColumnWidths.dependency + detailColumnWidths.status;
    // Cumulative left offset of each frozen (sticky) column - used both
    // for `position: sticky; left:` on every cell in that column and for
    // placing the today-marker after the frozen section.
    const stickyLeft = {
        id: 0,
        activity: taskIdColumnWidth,
        start: taskIdColumnWidth + taskColumnWidth,
        duration: taskIdColumnWidth + taskColumnWidth + detailColumnWidths.start,
        dependency: taskIdColumnWidth + taskColumnWidth + detailColumnWidths.start + detailColumnWidths.duration,
        status: taskIdColumnWidth + taskColumnWidth + detailColumnWidths.start + detailColumnWidths.duration + detailColumnWidths.dependency
    };
    const ganttOffsetPx = taskIdColumnWidth + taskColumnWidth + detailsTotalWidth;

    const totalTableWidth = ganttOffsetPx + (totalWeeks * weekWidth);
    table.style.width = `${totalTableWidth}px`;
    table.style.minWidth = `${totalTableWidth}px`;
    table.style.tableLayout = "fixed";

    // Explicit column definitions prevent the browser from redistributing
    // width between cells. This keeps Start/End dates and bar pixels on
    // exactly the same coordinate system.
    const colgroup = document.createElement("colgroup");
    const idCol = document.createElement("col");
    idCol.style.width = `${taskIdColumnWidth}px`;
    colgroup.appendChild(idCol);

    const taskCol = document.createElement("col");
    taskCol.style.width = `${taskColumnWidth}px`;
    colgroup.appendChild(taskCol);
    ["start", "duration", "dependency", "status"].forEach(key => {
        const col = document.createElement("col");
        col.style.width = `${detailColumnWidths[key]}px`;
        colgroup.appendChild(col);
    });
    for (let i = 0; i < totalWeeks; i++) {
        const weekCol = document.createElement("col");
        weekCol.style.width = `${weekWidth}px`;
        colgroup.appendChild(weekCol);
    }
    table.appendChild(colgroup);

    //----------------------------------------------------
    // Header - Three rows:
    //   Row 1: section labels (Activity / Task Details / Gantt Chart)
    //   Row 2: Start / Duration / Dependency / Progress column names,
    //          plus the Month headers for the timeline
    //   Row 3: Weekly start-date headers (e.g. 05 Jan, 12 Jan, ...)
    // Daily columns (weekday name + individual dates) have been
    // removed entirely per the compact weekly view - the timeline
    // now shows only whole-week columns. Bar positioning still uses
    // exact day-level math (see the bar-placement block below), so
    // this is purely a visual simplification, not a scheduling change.
    //----------------------------------------------------

    const thead = document.createElement("thead");

    function makeStickyHeaderCell(text, leftPx) {
        const th = document.createElement("th");
        th.innerHTML = text;
        th.style.position = "sticky";
        th.style.left = `${leftPx}px`;
        th.style.zIndex = "12";
        return th;
    }

    // Row 1: section labels
    const sectionRow = document.createElement("tr");
    sectionRow.className = "ganttHeaderSection";

    const idSectionTh = makeStickyHeaderCell("ID", stickyLeft.id);
    idSectionTh.rowSpan = 3;
    idSectionTh.style.width = `${taskIdColumnWidth}px`;
    idSectionTh.style.textAlign = "center";
    sectionRow.appendChild(idSectionTh);

    const activitySectionTh = makeStickyHeaderCell("Activity", stickyLeft.activity);
    activitySectionTh.rowSpan = 3;
    activitySectionTh.style.width = `${taskColumnWidth}px`;
    sectionRow.appendChild(activitySectionTh);

    const detailsSectionTh = makeStickyHeaderCell("Task Details", stickyLeft.start);
    detailsSectionTh.colSpan = 4;
    sectionRow.appendChild(detailsSectionTh);

    const ganttSectionTh = document.createElement("th");
    ganttSectionTh.innerHTML = "Gantt Chart";
    ganttSectionTh.colSpan = totalWeeks;
    sectionRow.appendChild(ganttSectionTh);

    // Row 2: Start / Duration / Dependency / Progress column names
    // (rowSpan 2, since they have no month/week sub-division of their
    // own) plus the Month headers for the timeline.
    const columnRow = document.createElement("tr");
    columnRow.className = "ganttHeaderColumns";

    const startTh = makeStickyHeaderCell("Start Date", stickyLeft.start);
    startTh.rowSpan = 2;
    startTh.style.width = `${detailColumnWidths.start}px`;
    columnRow.appendChild(startTh);

    const durationTh = makeStickyHeaderCell("Duration", stickyLeft.duration);
    durationTh.rowSpan = 2;
    durationTh.style.width = `${detailColumnWidths.duration}px`;
    columnRow.appendChild(durationTh);

    const dependencyTh = makeStickyHeaderCell("Dependency", stickyLeft.dependency);
    dependencyTh.rowSpan = 2;
    dependencyTh.style.width = `${detailColumnWidths.dependency}px`;
    columnRow.appendChild(dependencyTh);

    const statusTh = makeStickyHeaderCell("Progress", stickyLeft.status);
    statusTh.rowSpan = 2;
    statusTh.style.width = `${detailColumnWidths.status}px`;
    statusTh.style.borderRight = "2px solid #e2e8f0";
    columnRow.appendChild(statusTh);

    // Row 3: Weekly start-date headers (e.g. 05 Jan, 12 Jan, ...)
    const weekRow = document.createElement("tr");
    weekRow.className = "ganttHeaderWeek";

    // Track the month header cell currently being built. Rather than
    // computing colSpan retroactively by indexing back into the row
    // (fragile once a project spans 3+ months), each month's <th> is
    // created once and its colSpan is simply incremented by 1 for
    // every subsequent week that still falls in that same month.
    let currentMonth = null;
    let currentMonthCell = null;

    for (let w = 0; w < totalWeeks; w++) {
        const weekDate = addDays(weekStart, w * 7);

        // Week date cell
        // Show the date at the start of each week instead of "Week 1",
        // "Week 2", etc. The timeline is anchored to Monday, so each
        // label represents the Monday/start date of that weekly column.
        const weekTh = document.createElement("th");
        weekTh.style.width = weekWidth + "px";
        weekTh.className = "ganttWeekHeaderCell";

        // Show the actual start date of every week (for example, 5-Jan)
        // instead of a generic Week 1 / Week 2 label. The text is wrapped
        // in a span so CSS can rotate only the label while keeping the
        // table cell itself at the normal weekly width.
        const weekLabel = document.createElement("span");
        const day = weekDate.getDate();
        const month = weekDate.toLocaleDateString("en-GB", { month: "short" });
        weekLabel.textContent = `${day}-${month}`;
        weekTh.appendChild(weekLabel);

        weekTh.title = `Week starting ${weekDate.toLocaleDateString("en-GB", {
            day: "2-digit",
            month: "long",
            year: "numeric"
        })}`;
        weekRow.appendChild(weekTh);

        // Month tracking for row 2 - a week is attributed to the
        // month its Monday (week start) falls in.
        const monthKey = weekDate.getFullYear() + "-" + weekDate.getMonth();
        if (monthKey !== currentMonth) {
            currentMonth = monthKey;

            currentMonthCell = document.createElement("th");
            currentMonthCell.innerHTML = weekDate.toLocaleDateString("en-GB", {
                month: "long",
                year: "numeric"
            });
            currentMonthCell.colSpan = 1;
            columnRow.appendChild(currentMonthCell);
        } else {
            currentMonthCell.colSpan += 1;
        }
    }

    thead.appendChild(sectionRow);
    thead.appendChild(columnRow);
    thead.appendChild(weekRow);
    table.appendChild(thead);

    //----------------------------------------------------
    // Body
    //----------------------------------------------------

    const tbody = document.createElement("tbody");

    Project.tasks.forEach(task => {

        // ---------------------------------------------
        // Validate task data (see hasUsableDates note above -
        // this is what actually skips unschedulable tasks now,
        // instead of the old isNaN(new Date(null)) check that
        // never caught them)
        // ---------------------------------------------

        if (!hasUsableDates(task)) {
            console.warn("Skipping task with no valid dates yet:", task);
            return;
        }

        const startDate = new Date(task.startDate);
        const endDate = new Date(task.endDate);

        const activity = task.activity || "Untitled Task";

        const row = document.createElement("tr");

        // ---------------------------------------------
        // Task Name (ACTIVITY section)
        // ---------------------------------------------

        const idCell = document.createElement("td");
        idCell.className = "taskIdCell";
        idCell.textContent = task.id;
        idCell.title = `Activity ID: ${task.id}`;
        idCell.style.left = `${stickyLeft.id}px`;
        idCell.style.width = `${taskIdColumnWidth}px`;
        idCell.style.minWidth = `${taskIdColumnWidth}px`;
        idCell.style.maxWidth = `${taskIdColumnWidth}px`;
        row.appendChild(idCell);

        const taskCell = document.createElement("td");
        taskCell.className = "taskName";
        taskCell.title = task.isDowntime
            ? `Downtime ID: ${task.id} | ${activity}`
            : `ID: ${task.id} | ${activity}`;
        taskCell.style.left = `${stickyLeft.activity}px`;
        taskCell.style.width = `${taskColumnWidth}px`;
        taskCell.style.minWidth = `${taskColumnWidth}px`;
        taskCell.style.maxWidth = `${taskColumnWidth}px`;

        const taskLabel = document.createElement("span");
        taskLabel.className = task.isDowntime ? "ganttTaskLabel downtimeTaskLabel" : "ganttTaskLabel";
        taskLabel.textContent = activity;
        taskCell.appendChild(taskLabel);

        // A downtime is a first-class scheduled task, but only normal
        // activities can be used as the insertion point for another
        // downtime. The small action button stays hidden until the row
        // is hovered so the compact Gantt remains uncluttered.
        if (!task.isDowntime) {
            const addDowntimeButton = document.createElement("button");
            addDowntimeButton.type = "button";
            addDowntimeButton.className = "downtimeAddBtn";
            addDowntimeButton.dataset.taskId = task.id;
            addDowntimeButton.innerHTML = '<i class="fa-solid fa-plus"></i>';
            addDowntimeButton.title = `Add downtime after Activity ${task.id}`;
            addDowntimeButton.setAttribute("aria-label", `Add downtime after Activity ${task.id}`);
            taskCell.appendChild(addDowntimeButton);
        } else {
            // A persisted downtime must remain removable after a server
            // restart. The button is bound to the current numeric task ID,
            // which is restored from Excel on load.
            const removeDowntimeButton = document.createElement("button");
            removeDowntimeButton.type = "button";
            removeDowntimeButton.className = "downtimeRemoveBtn";
            removeDowntimeButton.dataset.taskId = task.id;
            removeDowntimeButton.innerHTML = '<i class="fa-solid fa-trash-can"></i>';
            removeDowntimeButton.title = `Remove downtime ${task.id}`;
            removeDowntimeButton.setAttribute("aria-label", `Remove downtime ${task.id}`);
            taskCell.appendChild(removeDowntimeButton);
        }

        row.appendChild(taskCell);

        // ---------------------------------------------
        // TASK DETAILS section - Start / Duration / Dependency /
        // Progress. These stay in the same <tr> as the Activity
        // name and the Gantt bar, so the three sections can never
        // drift out of vertical alignment with each other; they are
        // simply styled (via CSS position:sticky) to stay visible
        // while the Gantt section on the right scrolls horizontally.
        // ---------------------------------------------

        row.appendChild(buildStartDateCell(task, stickyLeft.start, detailColumnWidths.start));
        row.appendChild(buildDurationCell(task, stickyLeft.duration, detailColumnWidths.duration));
        row.appendChild(buildDependencyCell(task, stickyLeft.dependency, detailColumnWidths.dependency));
        row.appendChild(buildStatusCell(task, stickyLeft.status, detailColumnWidths.status));

        // ---------------------------------------------
        // GANTT section - one grid cell per WEEK (not per day).
        // This is purely a visual simplification: the bar itself is
        // still positioned/sized from the exact Start Date -> End
        // Date using day-level math below, so a task that starts or
        // ends mid-week still renders at the correct sub-week pixel
        // position. The weekly columns only change the grid lines
        // drawn behind the bar, never the underlying date calculation.
        // ---------------------------------------------

        for (let w = 0; w < totalWeeks; w++) {

            const weekDate = addDays(weekStart, w * 7);

            const cell = document.createElement("td");
            cell.className = "gridCell weekCell";
            cell.style.width = `${weekWidth}px`;
            cell.style.minWidth = `${weekWidth}px`;
            cell.style.maxWidth = `${weekWidth}px`;
            row.appendChild(cell);

        }

        // Calculate bar position in DAYS, not whole week columns.
        //
        // BUGFIX: the previous version snapped every bar to "left: 0"
        // of its start week column and sized it in whole-column
        // multiples (weekSpan * weekWidth). That's only correct when a
        // task happens to start on the Monday of its week - for any
        // task starting mid-week (the normal case), the bar was both
        // shifted left of its real start and cut short of its real
        // end, e.g. a task from Tue 06 Jan to Tue 13 Jan rendered as a
        // single Mon-Sun column instead of reaching 13 Jan.
        //
        // Fix: work in day-width units. The bar's left offset is how
        // many days the start date sits into the timeline.
        //
        // BUGFIX (off-by-one): differenceInDays(start, end) measures
        // the gap between the START of the Start Date and the START
        // of the End Date - it does not include the End Date's own
        // day. A task running 06 Jan -> 13 Jan therefore measured as
        // 7 days (06-12 Jan) and the bar's right edge landed at the
        // very beginning of 13 Jan instead of covering it, so every
        // bar stopped one day short of its actual End Date. Adding 1
        // makes the width span the End Date inclusively (06-13 Jan =
        // 8 days), matching what Start/End actually represent.
        //
        // The weekly grid now has one <td> per WEEK, not per day, so
        // the bar is placed into the exact WEEK cell its Start Date
        // falls in (startWeekIndex below), offset from the LEFT EDGE
        // of that week cell by however many days into the week the
        // Start Date sits (startOffsetDays % 7). It then overflows
        // across neighbouring week cells by its real pixel width -
        // exactly the same "overflow past the starting cell" approach
        // used before, just anchored to a week-wide cell instead of a
        // day-wide one.
        const startOffsetDays = differenceInDays(weekStart, task.startDate);
        const durationDays = Math.max(1, differenceInDays(task.startDate, task.endDate) + 1);

        const startWeekIndex = Math.floor(startOffsetDays / 7);
        const offsetWithinWeekDays = startOffsetDays - (startWeekIndex * 7);

        const barLeftPx = offsetWithinWeekDays * dayWidth;
        const barWidthPx = durationDays * dayWidth;

        // Create bar with progress
        const barContainer = document.createElement("div");
        barContainer.className = "ganttBarContainer";
        barContainer.style.cssText = `
            position: absolute;
            left: ${barLeftPx}px;
            width: ${barWidthPx}px;
            height: 22px;
            top: 10px;
            border-radius: 4px;
            overflow: hidden;
        `;

        // Downtime is a real scheduled task, but its visual language is
        // deliberately different from normal task-status colors.
        const status = task.status || "Not Started";
        barContainer.style.backgroundColor = task.isDowntime
            ? "#ef4444"
            : statusColor(status);
        if (task.isDowntime) {
            barContainer.classList.add("downtimeBar");
        }

        // ----------------------------------------------------
        // Progress Validation
        // ----------------------------------------------------

        let progress = Number(task.progress);

        if (isNaN(progress))
            progress = 0;

        progress = Math.max(0, Math.min(progress, 100));

        // ----------------------------------------------------
        // Progress Overlay
        // ----------------------------------------------------

        if (progress > 0) {

            const progressBar = document.createElement("div");
            progressBar.className = "ganttProgress";

            progressBar.style.cssText = `
                height: 100%;
                width: ${progress}%;
                background: rgba(255,255,255,0.3);
                border-radius: 4px 0 0 4px;
                transition: width 0.3s ease;
            `;

            if (progress === 100) {
                progressBar.style.borderRadius = "4px";
            }

            barContainer.appendChild(progressBar);
        }

        // ----------------------------------------------------
        // Tooltip
        // ----------------------------------------------------

        if (task.isDowntime) {
            barContainer.title =
`Downtime: ${activity}
Reason: ${task.reason || "No reason provided"}
Start: ${formatDate(startDate)}
End: ${formatDate(endDate)}
Duration: ${task.duration} week${task.duration === 1 ? "" : "s"}`;
        } else {
            barContainer.title =
`Task: ${activity}
Start: ${formatDate(startDate)}
End: ${formatDate(endDate)}
Progress: ${progress}%
Status: ${status}`;
        }

        // ----------------------------------------------------
        // Add Bar
        // ----------------------------------------------------

        // The 6 leading <td> cells in this row are ID + Activity + the 4
        // Task Details cells; the weekly grid cells start right after
        // them, so the bar's home cell is offset by 5, not 1.
        const firstCell = row.children[6 + startWeekIndex];

        if (firstCell) {

            // The bar belongs to the exact Start Date cell, but it must
            // visually continue through every day up to and including
            // End Date. Some table layouts clip absolutely positioned
            // children at cell boundaries unless overflow is explicitly
            // enabled, so make that contract explicit here as well as in CSS.
            firstCell.style.position = "relative";
            firstCell.style.overflow = "visible";
            barContainer.style.width = `${barWidthPx}px`;
            barContainer.style.maxWidth = "none";
            firstCell.appendChild(barContainer);

        }

        tbody.appendChild(row);
    });

    table.appendChild(tbody);
    container.appendChild(table);

    //----------------------------------------------------
    // Today Marker
    //----------------------------------------------------

    addTodayMarker(container, weekStart, weekWidth, totalWeeks, ganttOffsetPx);

    //----------------------------------------------------
    // Wire up the editable Task Details inputs (Start Date,
    // Duration, Dependency) just rendered above.
    //----------------------------------------------------

    attachGanttRowEvents(container);
}

// Backward-compatible alias expected by callers
function drawGantt() {
    drawCustomGantt();
}

/*=========================================================
            Monday-of-week helper
=========================================================*/
// Returns the Monday on/before the given date (the "project week
// start" for the week that date falls in). Does not mutate the
// scheduler's dates - purely a display/layout helper for the
// weekly Gantt grid.
function mondayOnOrBefore(date) {
    const d = new Date(date);
    const day = d.getDay(); // 0 = Sunday, 1 = Monday, ... 6 = Saturday
    const diffToMonday = (day === 0) ? 6 : day - 1;
    d.setDate(d.getDate() - diffToMonday);
    d.setHours(0, 0, 0, 0);
    return d;
}

/*=========================================================
            Today Marker
=========================================================*/

function addTodayMarker(container, weekStart, weekWidth, totalWeeks, ganttOffsetPx) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const start = new Date(weekStart);
    start.setHours(0, 0, 0, 0);

    // Fractional week offset, so the marker lands on the correct
    // day within its week column rather than snapping to the
    // start of the column.
    const diffDays = Math.floor((today - start) / (1000 * 60 * 60 * 24));
    const weekOffset = diffDays / 7;

    if (diffDays >= 0 && weekOffset < totalWeeks) {
        const table = container.querySelector(".ganttTable");
        if (!table) return;

        // ganttOffsetPx is the fixed pixel width of the frozen
        // Activity + Task Details columns to the left of the weekly
        // timeline (see drawCustomGantt) - using the same constant the
        // grid itself was built with keeps the red line accurate
        // regardless of how many detail columns precede the timeline.
        const marker = document.createElement("div");
        marker.className = "todayLine";
        marker.style.cssText = `
            position: absolute;
            left: ${ganttOffsetPx + weekOffset * weekWidth}px;
            top: 0;
            bottom: 0;
            width: 2px;
            background: #ef4444;
            z-index: 100;
            pointer-events: none;
        `;

        const tooltip = document.createElement("div");
        tooltip.style.cssText = `
            position: absolute;
            top: 4px;
            left: 50%;
            transform: translateX(-50%);
            background: #ef4444;
            color: white;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 11px;
            white-space: nowrap;
            pointer-events: none;
        `;
        tooltip.textContent = "Today: " + today.toLocaleDateString("en-GB", {
            day: "numeric",
            month: "short",
            year: "numeric"
        });
        marker.appendChild(tooltip);

        // Insert into the table header area
        table.style.position = "relative";
        table.appendChild(marker);
    }
}

/*=========================================================
        TASK DETAILS section - editable cell builders
=========================================================*/
// One builder per column (Start Date / Duration / Dependency /
// Progress). Each returns a single <td> that is appended into the
// same <tr> as the Activity name and the Gantt bar, so alignment is
// guaranteed by construction - there is only ever one row per task,
// never a separate row for a separate "panel".

function buildStartDateCell(task, leftPx, widthPx) {

    const cell = document.createElement("td");
    cell.className = "detailCell";
    cell.style.left = `${leftPx}px`;
    cell.style.width = `${widthPx}px`;
    cell.style.minWidth = `${widthPx}px`;
    cell.style.maxWidth = `${widthPx}px`;

    const input = document.createElement("input");
    input.type = "date";
    input.className = "ganttStartInput";
    input.dataset.taskId = task.id;

    if (task.startDate instanceof Date && !isNaN(task.startDate)) {
        input.value = toDateInputValue(task.startDate);
    }

    // Start Date is editable for every activity. When a task has a
    // dependency, the scheduler treats this value as the task's
    // preferred/manual start and the dependency end as the earliest
    // allowed start. This means the user can move a task later while
    // the dependency still prevents it from starting too early.
    input.title = (task.dependencies && task.dependencies.length > 0)
        ? "Editable start date. The dependency still sets the earliest allowed start."
        : "Editable start date.";

    cell.appendChild(input);
    return cell;

}

function buildDurationCell(task, leftPx, widthPx) {

    const cell = document.createElement("td");
    cell.className = "detailCell";
    cell.style.left = `${leftPx}px`;
    cell.style.width = `${widthPx}px`;
    cell.style.minWidth = `${widthPx}px`;
    cell.style.maxWidth = `${widthPx}px`;

    const wrap = document.createElement("div");
    wrap.className = "ganttDurationWrap";

    const input = document.createElement("input");
    input.type = "number";
    input.min = "1";
    input.step = "1";
    input.className = "ganttDurationInput";
    input.dataset.taskId = task.id;
    input.value = task.duration;

    const unit = document.createElement("span");
    unit.className = "ganttDurationUnit";
    unit.textContent = task.duration === 1 ? "Week" : "Weeks";

    wrap.appendChild(input);
    wrap.appendChild(unit);
    cell.appendChild(wrap);
    return cell;

}

function buildDependencyCell(task, leftPx, widthPx) {

    const cell = document.createElement("td");
    cell.className = "detailCell";
    cell.style.left = `${leftPx}px`;
    cell.style.width = `${widthPx}px`;
    cell.style.minWidth = `${widthPx}px`;
    cell.style.maxWidth = `${widthPx}px`;

    const select = document.createElement("select");
    select.className = "ganttDependencySelect";
    select.dataset.taskId = task.id;

    const noneOption = document.createElement("option");
    noneOption.value = "";
    noneOption.textContent = "None";
    select.appendChild(noneOption);

    const currentDependency = (task.dependencies && task.dependencies.length > 0)
        ? task.dependencies[0]
        : null;

    Project.tasks.forEach(other => {

        if (other.id === task.id) return; // no self-dependency

        const option = document.createElement("option");
        option.value = other.id;
        option.textContent = String(other.id);

        if (currentDependency === other.id) {
            option.selected = true;
        }

        select.appendChild(option);

    });

    if (currentDependency === null) {
        noneOption.selected = true;
    }

    cell.appendChild(select);
    return cell;

}

function buildStatusCell(task, leftPx, widthPx) {

    const cell = document.createElement("td");
    cell.className = "detailCell";
    cell.style.left = `${leftPx}px`;
    cell.style.width = `${widthPx}px`;
    cell.style.minWidth = `${widthPx}px`;
    cell.style.maxWidth = `${widthPx}px`;
    cell.style.borderRight = "2px solid #e2e8f0";

    const select = document.createElement("select");
    select.className = "ganttStatusSelect";
    select.dataset.taskId = task.id;

    ["Completed", "In Progress", "Not Started", "Delayed"].forEach(status => {
        const option = document.createElement("option");
        option.value = status;
        option.textContent = status;
        option.selected = (task.status || "Not Started") === status;
        select.appendChild(option);
    });

    cell.appendChild(select);

    return cell;

}

// Formats a Date as "YYYY-MM-DD" for an <input type="date"> value,
// using local date parts (not toISOString(), which shifts by the
// browser's UTC offset and can silently roll the date back/forward
// a day).
function toDateInputValue(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

/*=========================================================
        TASK DETAILS section - event wiring
=========================================================*/
// Attached fresh after every drawCustomGantt() call, since the whole
// table (including these inputs) is rebuilt from scratch each time.
// Each handler updates the underlying Project.tasks data through the
// scheduler (never just the displayed text), which is what makes
// Requirement 8 (live rescheduling) work: runScheduler() inside each
// scheduler.js function recalculates every affected task and
// drawGantt() redraws the chart with the new dates.

function attachGanttRowEvents(container) {

    container.querySelectorAll(".downtimeAddBtn").forEach(button => {
        button.addEventListener("click", function (event) {
            event.preventDefault();
            event.stopPropagation();
            openDowntimeModal(Number(this.dataset.taskId));
        });
    });

    container.querySelectorAll(".downtimeRemoveBtn").forEach(button => {
        button.addEventListener("click", function (event) {
            event.preventDefault();
            event.stopPropagation();

            const taskId = Number(this.dataset.taskId);

            if (typeof removeDowntime === "function") {
                removeDowntime(taskId);
            }
        });
    });

    container.querySelectorAll(".ganttStartInput").forEach(input => {

        input.addEventListener("change", function () {

            const taskId = Number(this.dataset.taskId);

            if (!this.value) return;

            // Parse the "YYYY-MM-DD" input value as a local date
            // (not UTC midnight) so it lines up with every other date
            // in the app, all of which are local-time Date objects.
            const [year, month, day] = this.value.split("-").map(Number);
            const newDate = new Date(year, month - 1, day);

            updateStartDate(taskId, newDate);

        });

    });

    container.querySelectorAll(".ganttDurationInput").forEach(input => {

        input.addEventListener("change", function () {

            const taskId = Number(this.dataset.taskId);
            const value = Math.max(1, Math.round(Number(this.value) || 1));

            updateDuration(taskId, value);

        });

    });

    container.querySelectorAll(".ganttDependencySelect").forEach(select => {

        select.addEventListener("change", function () {

            const taskId = Number(this.dataset.taskId);
            const value = this.value === "" ? null : this.value;

            updateDependency(taskId, value);

        });

    });

    container.querySelectorAll(".ganttStatusSelect").forEach(select => {

        select.addEventListener("change", function () {

            const taskId = Number(this.dataset.taskId);
            updateStatus(taskId, this.value);

        });

    });

}

/*=========================================================
        Downtime Creation Modal
=========================================================*/

function openDowntimeModal(parentTaskID) {

    const parent = findTask(parentTaskID);
    if (!parent) return;

    const existing = document.getElementById("downtimeModalOverlay");
    if (existing) existing.remove();

    const defaultStart = parent.endDate instanceof Date && !isNaN(parent.endDate)
        ? addDays(parent.endDate, 1)
        : new Date();

    const overlay = document.createElement("div");
    overlay.id = "downtimeModalOverlay";
    overlay.className = "downtimeModalOverlay";

    overlay.innerHTML = `
        <div class="downtimeModal" role="dialog" aria-modal="true" aria-labelledby="downtimeModalTitle">
            <div class="downtimeModalHeader">
                <div>
                    <h3 id="downtimeModalTitle">Add Downtime</h3>
                    <p>Insert a downtime event after Activity ${parent.id}.</p>
                </div>
                <button type="button" class="downtimeModalClose" aria-label="Close">&times;</button>
            </div>

            <form id="downtimeForm">
                <div class="downtimeField">
                    <label for="downtimeName">Downtime Name</label>
                    <input id="downtimeName" name="name" type="text" maxlength="120"
                           placeholder="e.g. Vehicle Breakdown" required>
                </div>

                <div class="downtimeField">
                    <label for="downtimeReason">Reason</label>
                    <textarea id="downtimeReason" name="reason" rows="3" maxlength="500"
                              placeholder="Explain why the project cannot progress..." required></textarea>
                </div>

                <div class="downtimeFormGrid">
                    <div class="downtimeField">
                        <label for="downtimeStart">Start Date</label>
                        <input id="downtimeStart" name="startDate" type="date"
                               value="${toDateInputValue(defaultStart)}" required>
                    </div>

                    <div class="downtimeField">
                        <label for="downtimeDuration">Duration</label>
                        <div class="downtimeDurationWrap">
                            <input id="downtimeDuration" name="duration" type="number"
                                   min="1" step="1" value="1" required>
                            <span>Week(s)</span>
                        </div>
                    </div>
                </div>

                <div id="downtimeFormError" class="downtimeFormError" hidden></div>

                <div class="downtimeModalActions">
                    <button type="button" class="downtimeCancelBtn">Cancel</button>
                    <button type="submit" class="downtimeSubmitBtn">
                        <i class="fa-solid fa-plus"></i>
                        Add Downtime
                    </button>
                </div>
            </form>
        </div>
    `;

    document.body.appendChild(overlay);

    const close = () => overlay.remove();

    overlay.querySelector(".downtimeModalClose").addEventListener("click", close);
    overlay.querySelector(".downtimeCancelBtn").addEventListener("click", close);

    overlay.addEventListener("click", event => {
        if (event.target === overlay) close();
    });

    overlay.querySelector("#downtimeForm").addEventListener("submit", event => {
        event.preventDefault();

        const name = overlay.querySelector("#downtimeName").value.trim();
        const reason = overlay.querySelector("#downtimeReason").value.trim();
        const startValue = overlay.querySelector("#downtimeStart").value;
        const duration = Math.max(1, Math.round(Number(
            overlay.querySelector("#downtimeDuration").value
        ) || 1));
        const error = overlay.querySelector("#downtimeFormError");

        if (!name || !reason || !startValue) {
            error.textContent = "Downtime name, reason and start date are required.";
            error.hidden = false;
            return;
        }

        const [year, month, day] = startValue.split("-").map(Number);
        const startDate = new Date(year, month - 1, day);

        if (isNaN(startDate.getTime())) {
            error.textContent = "Please enter a valid start date.";
            error.hidden = false;
            return;
        }

        const created = addDowntimeAfterActivity(
            parentTaskID,
            name,
            reason,
            startDate,
            duration
        );

        if (!created) {
            error.textContent = "The downtime could not be added. Check the browser console for details.";
            error.hidden = false;
            return;
        }

        close();
    });

    const nameInput = overlay.querySelector("#downtimeName");
    if (nameInput) nameInput.focus();
}
