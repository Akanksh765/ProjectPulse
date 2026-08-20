/*=========================================================
                ProjectPulse V2
                dashboard.js
=========================================================*/


/*=========================================================
                Dashboard Refresh
=========================================================*/

function updateDashboard() {

    updateKPIs();
    updateDashboardWidgets();

}


/*=========================================================
                KPI Cards
=========================================================*/

function updateKPIs() {

    // Guard: these cards only exist while the Dashboard page is the
    // active page (see app.js). If something calls updateDashboard()
    // while the user is looking at the Gantt page (or any other page),
    // silently skip instead of throwing on a null element.
    if (!document.getElementById("totalTasks"))
        return;

    const totalTasks = Project.tasks.length;

    const completedTasks = Project.tasks.filter(
        t => t.status === "Completed"
    ).length;

    const delayedTasks = Project.tasks.filter(
        t => t.status === "Delayed"
    ).length;

    const progress = totalTasks === 0
        ? 0
        : Math.round(

            Project.tasks.reduce(

                (sum, task) =>

                sum + task.progress,

                0

            ) / totalTasks

        );

    const health = calculateHealth();


    document.getElementById("totalTasks").textContent =
        totalTasks;

    document.getElementById("delayedTasks").textContent =
        delayedTasks;

    document.getElementById("projectHealth").textContent =
        health + "%";


    if (Project.projectEnd) {

        document.getElementById("projectFinish").textContent =
            formatDate(Project.projectEnd);

    }

}


/*=========================================================
        Dashboard Overview Widgets
=========================================================*/

let projectProgressChart = null;

function updateDashboardWidgets() {

    const stats = getProjectStatistics();

    updateProjectProgressChart(stats);
    updateUpcomingActivities();
    updateCurrentProjectStatus();

}

function updateProjectProgressChart(stats) {

    const canvas = document.getElementById("projectProgressChart");
    const center = document.getElementById("projectProgressCenter");

    if (!canvas || !center || typeof Chart === "undefined")
        return;

    const total = stats.totalTasks;
    const completed = stats.completed;
    const inProgress = stats.inProgress;
    const notStarted = stats.notStarted;

    const overallProgress = total === 0
        ? 0
        : Math.round(
            Project.tasks.reduce((sum, task) => sum + Number(task.progress || 0), 0) / total
        );

    center.textContent = overallProgress + "%";

    if (projectProgressChart) {
        projectProgressChart.destroy();
        projectProgressChart = null;
    }

    projectProgressChart = new Chart(canvas.getContext("2d"), {
        type: "doughnut",
        data: {
            labels: ["Completed", "In Progress", "Not Started"],
            datasets: [{
                data: [completed, inProgress, notStarted],
                backgroundColor: ["#22c55e", "#2563eb", "#f59e0b"],
                borderWidth: 0,
                hoverOffset: 5
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: "70%",
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return ` ${context.label}: ${context.raw}`;
                        }
                    }
                }
            }
        }
    });

    const legend = document.getElementById("projectProgressLegend");
    if (!legend) return;

    legend.innerHTML = `
        <div class="progressLegendItem">
            <span class="legendDot completed"></span>
            <span>Completed</span>
            <strong>${completed}</strong>
        </div>
        <div class="progressLegendItem">
            <span class="legendDot inProgress"></span>
            <span>In Progress</span>
            <strong>${inProgress}</strong>
        </div>
        <div class="progressLegendItem">
            <span class="legendDot notStarted"></span>
            <span>Not Started</span>
            <strong>${notStarted}</strong>
        </div>
    `;

}

function updateUpcomingActivities() {

    const container = document.getElementById("upcomingActivities");
    if (!container) return;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let upcoming = Project.tasks
        .filter(task => task.startDate instanceof Date && !isNaN(task.startDate))
        .filter(task => task.startDate >= today && task.status !== "Completed")
        .sort((a, b) => a.startDate - b.startDate)
        .slice(0, 4);

    // If the project data is historical, still show the next unfinished
    // activities instead of leaving the dashboard empty.
    if (upcoming.length === 0) {
        upcoming = Project.tasks
            .filter(task => task.status !== "Completed")
            .sort((a, b) => {
                const aDate = a.startDate instanceof Date ? a.startDate.getTime() : Infinity;
                const bDate = b.startDate instanceof Date ? b.startDate.getTime() : Infinity;
                return aDate - bDate || a.id - b.id;
            })
            .slice(0, 4);
    }

    if (upcoming.length === 0) {
        container.innerHTML = `<div class="dashboardEmptyState">No upcoming activities.</div>`;
        return;
    }

    container.innerHTML = upcoming.map(task => `
        <div class="upcomingItem">
            <div class="upcomingIcon"><i class="fa-solid fa-calendar-day"></i></div>
            <div class="upcomingDetails">
                <div class="upcomingActivity">${escapeDashboardHTML(task.activity)}</div>
                <div class="upcomingMeta">
                    ${task.startDate ? formatDate(task.startDate) : "Date not set"}
                    ${task.endDate ? ` – ${formatDate(task.endDate)}` : ""}
                </div>
            </div>
            <span class="upcomingStatus ${statusClass(task.status)}">${escapeDashboardHTML(task.status)}</span>
        </div>
    `).join("");

}

function updateCurrentProjectStatus() {

    const activity = document.getElementById("currentActivity");
    const progress = document.getElementById("currentActivityProgress");
    const finish = document.getElementById("currentActivityFinish");
    const next = document.getElementById("nextActivity");

    if (!activity || !progress || !finish || !next) return;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let current = Project.tasks.find(task => task.status === "In Progress");

    if (!current) {
        current = Project.tasks.find(task =>
            task.startDate instanceof Date &&
            task.endDate instanceof Date &&
            task.startDate <= today &&
            task.endDate >= today &&
            task.status !== "Completed"
        );
    }

    if (!current) {
        current = Project.tasks
            .filter(task => task.status !== "Completed")
            .sort((a, b) => {
                const aDate = a.startDate instanceof Date ? a.startDate.getTime() : Infinity;
                const bDate = b.startDate instanceof Date ? b.startDate.getTime() : Infinity;
                return aDate - bDate || a.id - b.id;
            })[0];
    }

    if (!current) {
        activity.textContent = "No active activity";
        progress.textContent = "0%";
        finish.textContent = "--";
        next.textContent = "No pending activities";
        return;
    }

    activity.textContent = current.activity;
    progress.textContent = `${Number(current.progress || 0)}%`;
    finish.textContent = current.endDate ? formatDate(current.endDate) : "--";

    const currentIndex = Project.tasks.indexOf(current);
    const nextTask = Project.tasks
        .filter((task, index) => index > currentIndex && task.status !== "Completed")
        .sort((a, b) => {
            const aDate = a.startDate instanceof Date ? a.startDate.getTime() : Infinity;
            const bDate = b.startDate instanceof Date ? b.startDate.getTime() : Infinity;
            return aDate - bDate || a.id - b.id;
        })[0];

    next.textContent = nextTask ? nextTask.activity : "No pending activities";

}

function escapeDashboardHTML(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

/*=========================================================
                Health Calculation
=========================================================*/

function calculateHealth() {

    if (Project.tasks.length === 0)
        return 100;

    let score = 100;

    Project.tasks.forEach(task => {

        if (task.status === "Delayed")
            score -= 5;

        if (task.priority === "High" &&
            task.status === "Delayed")
            score -= 3;

    });

    if (score < 0)
        score = 0;

    return score;

}


/*=========================================================
                Tasks Page Table
=========================================================*/
// Same row shape/order as the Dashboard's table (via
// buildTaskRowsHTML below), just targeting the Tasks page's own
// <tbody>. Always reflects the CURRENT Project.tasks - called on
// every visit to the Tasks page and after every Excel refresh.

function updateTasksPageTable() {

    const tbody =
        document.getElementById("tasksPageTable");

    if (!tbody)
        return;

    tbody.innerHTML = buildTaskRowsHTML(Project.tasks);

}


/*=========================================================
                Task Row Markup (shared)
=========================================================*/
// Single source of truth for how a task becomes a <tr> - used by
// both the Dashboard table and the Tasks page table so they can
// never drift out of sync with each other. Tasks are rendered in
// the exact order of the array passed in - callers are expected to
// pass Project.tasks itself, which is always kept in Excel ID order
// (1, 2, 3 ... 21) by the scheduler; this function does no sorting
// of its own.

function buildTaskRowsHTML(tasks) {

    return tasks.map(task => `

        <tr>

            <td>${task.id}</td>

            <td>${task.activity}</td>

            <td>${task.owner || ""}</td>

            <td>

                <span class="${statusClass(task.status)}">

                    ${task.status}

                </span>

            </td>

            <td>

                ${task.progress}%

            </td>

            <td>

                ${formatDate(task.startDate)}

            </td>

            <td>

                ${formatDate(task.endDate)}

            </td>

            <td>

                ${task.duration} Week(s)

            </td>

        </tr>

    `).join("");

}


/*=========================================================
            Status CSS Classes
=========================================================*/

function statusClass(status) {

    switch (status) {

        case "Completed":

            return "statusCompleted";

        case "In Progress":

            return "statusProgress";

        case "Delayed":

            return "statusDelayed";

        default:

            return "statusNotStarted";

    }

}


/*=========================================================
            Project Statistics
=========================================================*/

function getProjectStatistics() {

    return {

        totalTasks:

            Project.tasks.length,

        completed:

            Project.tasks.filter(

                t => t.status === "Completed"

            ).length,

        delayed:

            Project.tasks.filter(

                t => t.status === "Delayed"

            ).length,

        inProgress:

            Project.tasks.filter(

                t => t.status === "In Progress"

            ).length,

        notStarted:

            Project.tasks.filter(

                t => t.status === "Not Started"

            ).length,

        projectStart:

            Project.projectStart,

        projectEnd:

            Project.projectEnd

    };

}
