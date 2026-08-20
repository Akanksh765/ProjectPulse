/*=========================================================
                ProjectPulse V2
                  app.js
=========================================================*/
document.addEventListener("DOMContentLoaded", initializeApp);


/*=========================================================
            Page Router State
=========================================================*/
// Which sidebar page is currently shown in the <main id="dashboard">
// content area. Only "dashboard" and "gantt" are fully wired up
// (that's the whole app so far); the rest render a simple
// placeholder instead of doing nothing when clicked.
let currentPage = "dashboard";

const PAGE_TITLES = {
    dashboard: "Dashboard",
    tasks: "Tasks",
    gantt: "Gantt",
    analytics: "Analytics",
    risks: "Risks",
    reports: "Reports"
};


function initializeApp() {

    renderSidebar();

    attachSidebarNavigation();

    renderPage();

    attachEvents();

    // Automatically load the configured local Excel workbook when the
    // app is started through server.py and the configured workbook is
    // loaded automatically.
    if (typeof loadWorkbookFromServer === "function") {
        loadWorkbookFromServer();
    }

}


/*=========================================================
                Sidebar
=========================================================*/

function renderSidebar() {

    document.getElementById("sidebar").innerHTML = `

    <div class="sidebar">

        <div class="logo">

            <i class="fa-solid fa-diagram-project"></i>

            <span>ProjectPulse</span>

        </div>

        <div class="menu">

            <div class="menuItem active" data-page="dashboard">
                <i class="fa-solid fa-chart-line"></i>
                Dashboard
            </div>

            <div class="menuItem" data-page="tasks">
                <i class="fa-solid fa-list-check"></i>
                Tasks
            </div>

            <div class="menuItem" data-page="gantt">
                <i class="fa-solid fa-calendar-days"></i>
                Gantt
            </div>

            <div class="menuItem" data-page="analytics">
                <i class="fa-solid fa-chart-pie"></i>
                Analytics
            </div>

            <div class="menuItem" data-page="risks">
                <i class="fa-solid fa-triangle-exclamation"></i>
                Risks
            </div>

            <div class="menuItem" data-page="reports">
                <i class="fa-solid fa-file-export"></i>
                Reports
            </div>

        </div>

        <div class="profile">

            <div class="avatar">

                PP

            </div>

            <div class="userInfo">

                <span>Project Manager</span>

                <span>ProjectPulse V2</span>

            </div>

        </div>

    </div>

    `;

}


/*=========================================================
            Sidebar Navigation
=========================================================*/
// The sidebar itself is only rendered once (renderSidebar() is not
// called again on every page switch), so its click listeners are
// attached once here rather than inside attachEvents(), which is
// re-run every time the page content is re-rendered.

function attachSidebarNavigation() {

    const items = document.querySelectorAll(".menuItem");

    items.forEach(item => {

        item.addEventListener("click", function () {
            navigateTo(this.dataset.page);
        });

    });

}


function navigateTo(page) {

    if (!page || page === currentPage) {

        // Still re-sync the active highlight even if it's a no-op
        // navigation, just in case something got out of sync.
        highlightActiveMenuItem();
        return;

    }

    currentPage = page;

    highlightActiveMenuItem();

    renderPage();

    attachEvents();

    if (page === "dashboard") {

        updateDashboard();

    }

    if (page === "gantt" && typeof drawGantt === "function") {

        drawGantt();

    }

}


function highlightActiveMenuItem() {

    document.querySelectorAll(".menuItem").forEach(item => {

        item.classList.toggle(
            "active",
            item.dataset.page === currentPage
        );

    });

}


/*=========================================================
            Page Dispatch
=========================================================*/

function renderPage() {

    switch (currentPage) {

        case "dashboard":
            renderDashboardPage();
            break;

        case "tasks":
            renderTasksPage();
            break;

        case "gantt":
            renderGanttPage();
            break;

        default:
            renderPlaceholderPage(currentPage);

    }

}


/*=========================================================
            Refresh Current Page
=========================================================*/
// Re-renders whichever page is currently on screen using the latest
// Project.tasks, WITHOUT changing currentPage or re-running
// navigation. Used after the local Excel bridge refreshes the project
// so the visible page reflects the latest data immediately.

function refreshCurrentPage() {

    renderPage();

    attachEvents();

    if (currentPage === "gantt" && typeof drawGantt === "function") {
        drawGantt();
    }

}


/*=========================================================
                Dashboard Page
=========================================================*/
// KPI cards + searchable task table. The Gantt chart used to live
// at the bottom of this same page, which is why it only ever showed
// a handful of tasks comfortably - everything was squeezed under
// the cards and the table in one long scrolling page. It now has
// its own page (see renderGanttPage below), reached from the
// sidebar's "Gantt" item.

function renderDashboardPage() {

    document.getElementById("dashboard").innerHTML = `

<div class="dashboard dashboardHome">

    <div class="header">

        <h1>Project Dashboard</h1>

        <div class="actions">

            <span id="excelSyncStatus" class="excelSyncStatus" title="Local Excel connection status">
                Connecting to Excel...
            </span>

        </div>

    </div>

    <div class="cards dashboardKpis">

        <div class="card">
            <div class="cardTitle">Total Tasks</div>
            <div class="cardValue" id="totalTasks">0</div>
        </div>

        <div class="card">
            <div class="cardTitle">Delayed</div>
            <div class="cardValue" id="delayedTasks">0</div>
        </div>

        <div class="card">
            <div class="cardTitle">Project Health</div>
            <div class="cardValue" id="projectHealth">100%</div>
        </div>

        <div class="card">
            <div class="cardTitle">Project Finish</div>
            <div class="cardValue" id="projectFinish">--</div>
        </div>

    </div>

    <div class="dashboardOverviewGrid">

        <section class="dashboardPanel progressPanel">
            <div class="panelHeading">
                <div>
                    <h3>Project Progress</h3>
                    <p>Current task distribution</p>
                </div>
            </div>

            <div class="progressContent">
                <div class="donutWrap">
                    <canvas id="projectProgressChart"></canvas>
                    <div class="donutCenter">
                        <strong id="projectProgressCenter">0%</strong>
                        <span>Completed</span>
                    </div>
                </div>

                <div id="projectProgressLegend" class="progressLegend"></div>
            </div>
        </section>

        <section class="dashboardPanel upcomingPanel">
            <div class="panelHeading">
                <div>
                    <h3>Upcoming Activities</h3>
                    <p>What is coming next</p>
                </div>
                <button class="viewAllBtn" id="viewAllTasksBtn" type="button">
                    View all <i class="fa-solid fa-arrow-right"></i>
                </button>
            </div>

            <div id="upcomingActivities" class="upcomingList"></div>
        </section>

    </div>

    <section class="dashboardPanel currentStatusPanel">
        <div class="panelHeading">
            <div>
                <h3>Current Project Status</h3>
                <p>The activity currently driving the project</p>
            </div>
            <span class="statusPulse"><span></span> Live overview</span>
        </div>

        <div class="currentStatusGrid">
            <div class="currentStatusMain">
                <span class="statusLabel">Current Activity</span>
                <strong id="currentActivity">No active activity</strong>
            </div>

            <div class="currentMetric">
                <span class="statusLabel">Progress</span>
                <strong id="currentActivityProgress">0%</strong>
            </div>

            <div class="currentMetric">
                <span class="statusLabel">Finish</span>
                <strong id="currentActivityFinish">--</strong>
            </div>

            <div class="currentNext">
                <span class="statusLabel">Next Activity</span>
                <strong id="nextActivity">No pending activities</strong>
            </div>
        </div>
    </section>

</div>

`;

    updateDashboard();

    const viewAll = document.getElementById("viewAllTasksBtn");
    if (viewAll) {
        viewAll.addEventListener("click", () => navigateTo("tasks"));
    }

}


/*=========================================================
                Tasks Page
=========================================================*/
// Shows every task from Project.tasks - the single normalized
// collection populated by the Excel import/scheduler pipeline - in
// the required column set, in the collection's own order (which is
// always Excel ID order; see scheduler.js for why that's now
// guaranteed). Reuses the existing .tableContainer / .projectTable
// styling so it matches the Dashboard's task table.

function renderTasksPage() {

    document.getElementById("dashboard").innerHTML = `

<div class="dashboard">

    <div class="header">

        <h1>Tasks</h1>

        <div class="actions">

            <input
                type="text"
                class="searchBox"
                id="searchBox"
                placeholder="Search Activity...">

        </div>

    </div>

    <div class="tableContainer">

        <h3 class="sectionTitle">

            All Tasks

        </h3>

        <table class="projectTable">

            <thead>

                <tr>

                    <th>ID</th>

                    <th>Activity</th>

                    <th>Owner</th>

                    <th>Status</th>

                    <th>Progress</th>

                    <th>Start</th>

                    <th>Finish</th>

                    <th>Duration</th>

                </tr>

            </thead>

            <tbody id="tasksPageTable">

            </tbody>

        </table>

    </div>

    ${Project.tasks.length === 0
        ? '<p style="color:#94a3b8;margin-top:15px;">No tasks are currently available from the connected Excel workbook.</p>'
        : ""}

</div>

`;

    updateTasksPageTable();

}


/*=========================================================
                Gantt Page
=========================================================*/
// Its own dedicated page (instead of a small block wedged under the
// task table) so the chart gets the full content area to work with -
// this is what actually makes a 20+ task project usable: the grid
// gets real vertical room and scrolls on its own instead of being
// squeezed at the bottom of an already-long dashboard page.

function renderGanttPage() {

    document.getElementById("dashboard").innerHTML = `

<div class="dashboard ganttPageWrapper">

    <div class="header">

        <h1>Project Timeline</h1>

        <div class="actions"></div>

    </div>

    <div class="ganttContainer ganttPageContainer">

        <div id="ganttContainer"></div>

    </div>

</div>

`;

    if (Project.tasks.length === 0) {

        document.getElementById("ganttContainer").innerHTML = `
            <div class="ganttEmpty">
                <i class="fa-solid fa-calendar-days"></i>
                <p>No tasks are currently available from the connected Excel workbook.</p>
            </div>
        `;

    }

}


/*=========================================================
                Placeholder Pages
=========================================================*/
// Tasks / Analytics / Risks / Reports aren't built out yet. Rather
// than leaving the sidebar item looking broken (clicking it and
// nothing happening), show a plain "coming soon" state.

function renderPlaceholderPage(page) {

    const title = PAGE_TITLES[page] || "Page";

    document.getElementById("dashboard").innerHTML = `

<div class="dashboard">

    <div class="header">

        <h1>${title}</h1>

    </div>

    <div class="comingSoon">

        <i class="fa-solid fa-hammer"></i>

        <p>${title} isn't built yet - check back soon.</p>

    </div>

</div>

`;

}


/*=========================================================
            Event Listeners
=========================================================*/
// Re-run after every page render, since renderDashboardPage() /
// renderGanttPage() replace the #dashboard content wholesale
// (innerHTML =), which destroys the old elements and any listeners
// on them. Every lookup here is null-guarded, so it's safe to call
// regardless of which page is currently on screen - a page that
// doesn't have a given control just skips wiring it up.

function attachEvents() {

    const search = document.getElementById("searchBox");

    if (search) {

        search.addEventListener("keyup", function () {

            filterTasks(this.value);

        });

    }

}


/*=========================================================
                Search
=========================================================*/

function filterTasks(keyword) {

    keyword = keyword.toLowerCase();

    const rows = document.querySelectorAll("#taskTable tr, #tasksPageTable tr");

    rows.forEach(row => {

        row.style.display =
            row.innerText
                .toLowerCase()
                .includes(keyword)

                ? ""

                : "none";

    });

}
