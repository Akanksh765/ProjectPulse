/*==================================================
 Utility Functions
===================================================*/


/*------------------------------------------
 Excel Date Parsing
-------------------------------------------*/
// The single, authoritative date parser (parseExcelDate) lives in
// excel.js. It used to be duplicated here under the same name as a
// simpler, less correct version - having two functions with the
// same name in different files meant the app's date parsing
// correctness depended entirely on <script> load order (whichever
// file loaded last silently won). Do not re-add a date parser here;
// use parseExcelDate() instead.


/*------------------------------------------
 Add Days
-------------------------------------------*/

function addDays(date, days) {

    const d = new Date(date);

    d.setDate(d.getDate() + days);

    return d;

}


/*------------------------------------------
 Format Date
-------------------------------------------*/

function formatDate(date) {

    if (!(date instanceof Date))

        return "";

    if (isNaN(date))

        return "";

    return date.toLocaleDateString(

        "en-GB",

        {

            day: "2-digit",

            month: "short",

            year: "numeric"

        }

    );

}


/*------------------------------------------
 Find Task
-------------------------------------------*/

function findTask(id) {

    return Project.tasks.find(

        task => task.id === id

    );

}


/*------------------------------------------
 Get Latest Dependency End Date
-------------------------------------------*/

function latestDependencyEnd(task) {

    let latest = null;

    task.dependencies.forEach(depID => {

        const dependency = findTask(depID);

        if (!dependency)

            return;

        if (

            latest === null ||

            dependency.endDate > latest

        ) {

            latest = dependency.endDate;

        }

    });

    return latest;

}


/*------------------------------------------
 Status Color
-------------------------------------------*/

function statusColor(status) {

    switch (status) {

        case "Completed":

            return "#22c55e";

        case "Delayed":

            return "#ef4444";

        case "In Progress":

            return "#2563eb";

        case "Critical":

            return "#f97316";

        default:

            return "#94a3b8";

    }

}


/*------------------------------------------
 Console Banner
-------------------------------------------*/

function banner() {

    console.log(

        "%cProjectPulse V2 Loaded",

        "color:white;background:#2563eb;padding:8px;font-size:14px;border-radius:5px"

    );

}

banner();
/*=========================================================
            Gantt Date Utilities
=========================================================*/

function startOfDay(date) {

    const d = new Date(date);

    d.setHours(0, 0, 0, 0);

    return d;

}

function differenceInDays(start, end) {

    const oneDay = 1000 * 60 * 60 * 24;

    return Math.round(

        (startOfDay(end) - startOfDay(start)) / oneDay

    );

}
