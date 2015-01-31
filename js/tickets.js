//
// Main page loading and data processing script
//----------------------------------------------
"use strict";

var JSON_DATA = [];   // Use global variables to allow Google and YUI to coexist
var KPI_nTasks, KPI_nWaiting, KPI_nRunning, KPI_nCompleted, KPI_nDelayed, KPI_nStopped;

var set_Status_Icon = function (status) {
    var sCodes = {
        "waiting": "fa-pause",
        "running": "fa-cog fa-spin",
        "complete": "fa-check-circle",
        "delayed": "fa-warning",
        "stopped": "fa-exclamation-circle"
    };
    return ( "<i class='fa " + sCodes[status.toLowerCase()] + " fa-fw'></i>");
}

var set_Status_Color = function (status) {
    var sCodes = {
        "waiting": "#a2a6eb",
        "running": "rgb(66, 184, 221)",
        "complete": "rgb(28, 184, 65)",
        "delayed": "rgb(223, 117, 20)",
        "stopped": "rgb(202, 60, 60)"
    };
    return ( "color:" + sCodes[status.toLowerCase()]);
}

var dCount = function (data, ixPhase, ixStat) {
    var sPhases = ["01 - Business preparation", "02 - Technical cutover", "03 - Master data migration", "04 - Stock migration", "05 - Sales orders migration", "06 - Technical setup"];
    var sStatus = ["Waiting", "Running", "Complete", "Delayed", "Stopped"];
    return( data.getFilteredRows([{column:9, value:sPhases[ixPhase]}, {column:10, value:sStatus[ixStat]}]).length );
}

var data_display = function() {
    // Process JSON_DATA into a Google DataTable
    var data = new google.visualization.DataTable();
    data.addColumn("string", "✔");
    data.addColumn("number", "#");
    data.addColumn("string", "Task");
    data.addColumn("string", "Owner");
    data.addColumn("datetime", "Starts");
    data.addColumn("datetime", "Ends");
    data.addColumn("number", "Dep.On");
    data.addColumn("string", "Remarks");
    data.addColumn("boolean", "C.Path");
    data.addColumn("string", "Phase");
    data.addColumn("string", "Status");
    data.addRows(JSON_DATA.length);
    for( var j in JSON_DATA ) {
        // console.log("#:" + JSON_DATA[j].id + " - " + JSON_DATA[j].subject);
        data.setValue(parseInt(j), 0, set_Status_Icon(JSON_DATA[j].status));
        data.setValue(parseInt(j), 1, JSON_DATA[j].id);
        data.setValue(parseInt(j), 2, JSON_DATA[j].task);
        data.setValue(parseInt(j), 3, JSON_DATA[j].owner);
        data.setValue(parseInt(j), 4, new Date(Date.parse(JSON_DATA[j].starts)));
        data.setValue(parseInt(j), 5, ( JSON_DATA[j].ends.length > 0 ? new Date(Date.parse(JSON_DATA[j].ends)) : null ));
        data.setValue(parseInt(j), 6, ( JSON_DATA[j].dependsOn > 0 ? JSON_DATA[j].dependsOn : null ));
        data.setValue(parseInt(j), 7, JSON_DATA[j].remarks);
        data.setValue(parseInt(j), 8, JSON_DATA[j].cpath);
        data.setValue(parseInt(j), 9, JSON_DATA[j].phase);
        data.setValue(parseInt(j), 10, JSON_DATA[j].status);
        data.setProperty(parseInt(j), 0, "style", set_Status_Color(JSON_DATA[j].status));
        data.setProperty(parseInt(j), 7, "className", "remark-text");
    }
    // Apply formats
    var dFmt1 = new google.visualization.DateFormat({pattern: "EEE dd HH:mm"});
    var dFmt2 = new google.visualization.DateFormat({pattern: "HH:mm"});
    dFmt1.format(data, 4);
    dFmt2.format(data, 5);
    // Create views with data ready to display
    var tableView = new google.visualization.DataView(data);
    tableView.hideColumns([6, 8, 9, 10]);
    tableView.hideRows(data.getFilteredRows([{column:10, value:"Complete"}]));
    var statusView = new google.visualization.data.group(data, [10], [{"column": 1, "aggregation": google.visualization.data.count, "type": "number"}]);
    var phaseView = new google.visualization.DataTable();
    phaseView.addColumn("string", "Phase");
    phaseView.addColumn("number", "Waiting");
    phaseView.addColumn("number", "Running");
    phaseView.addColumn("number", "Complete");
    phaseView.addColumn("number", "Delayed");
    phaseView.addColumn("number", "Stopped");
    phaseView.addRow(["1.Business Prep",   dCount(data,0,0), dCount(data,0,1), dCount(data,0,2), dCount(data,0,3), dCount(data,0,4)]);
    phaseView.addRow(["2.Technical Cutover",dCount(data,1,0), dCount(data,1,1), dCount(data,1,2), dCount(data,1,3), dCount(data,1,4)]);
    phaseView.addRow(["3.Master Data",   dCount(data,2,0), dCount(data,2,1), dCount(data,2,2), dCount(data,2,3), dCount(data,2,4)]);
    phaseView.addRow(["4.Stock Migration",   dCount(data,3,0), dCount(data,3,1), dCount(data,3,2), dCount(data,3,3), dCount(data,3,4)]);
    phaseView.addRow(["5.Sales Orders",  dCount(data,4,0), dCount(data,4,1), dCount(data,4,2), dCount(data,4,3), dCount(data,4,4)]);
    phaseView.addRow(["6.Technical Setup",  dCount(data,5,0), dCount(data,5,1), dCount(data,5,2), dCount(data,5,3), dCount(data,5,4)]);
    // Build a Dashboard out of that
    var tasksTable = new google.visualization.Table( document.getElementById("tasks_table") );
    tasksTable.draw(tableView, {allowHtml: true, sortColumn: 4});
    var statusPie = new google.visualization.PieChart( document.getElementById("status_pie") );
    statusPie.draw(statusView, {tooltip: {showColorCode:true}, chartArea: {left:4, top:4, width:"90%", height:"90%"}, legend:{alignment:"center"}, colors:["rgb(28, 184, 65)","rgb(223, 117, 20)","rgb(66, 184, 221)","rgb(202, 60, 60)","#a2a6eb"]});
    var phaseColumn = new google.visualization.ColumnChart( document.getElementById("phase_column") );
    phaseColumn.draw(phaseView, {isStacked:true, legend:"bottom", colors:["#a2a6eb","rgb(66, 184, 221)","rgb(28, 184, 65)","rgb(223, 117, 20)","rgb(202, 60, 60)"]});
    // Update KPIs
    KPI_nTasks = data.getNumberOfRows();
    KPI_nWaiting = data.getFilteredRows([{column:10, value:"Waiting"}]).length;
    KPI_nRunning = data.getFilteredRows([{column:10, value:"Running"}]).length;
    KPI_nCompleted = data.getFilteredRows([{column:10, value:"Complete"}]).length;
    KPI_nDelayed = data.getFilteredRows([{column:10, value:"Delayed"}]).length;
    KPI_nStopped = data.getFilteredRows([{column:10, value:"Stopped"}]).length;
}

YUI().use("node-base", "node-event-delegate", function (Y) {
    // This just makes sure that the href="#" attached to the <a> elements
    // don't scroll you back up the page.
    Y.one("body").delegate("click", function (e) {
        e.preventDefault();
    }, 'a[href="#"]');
});

YUI().use("transition", "node", function (Y) {
    // Toggle chart visibility
    Y.one("#chart_toggle").on("click", function(e) {
        var chevron = Y.one("#chevron");
        if( chevron.hasClass("fa-chevron-up") ) {
            Y.one("#charts").transition({opacity: 0, height: "0px", easing: "ease-out", duration: 0.35});
            chevron.removeClass("fa-chevron-up").addClass("fa-chevron-down");   //hide
        } else if( chevron.hasClass("fa-chevron-down") ) {
            Y.one("#charts").transition({opacity: 1.0, height: "441px", easing: "ease-in", duration: 0.35});
            chevron.removeClass("fa-chevron-down").addClass("fa-chevron-up");   //show
        }
    });
});

YUI().use("io", "json-parse", "node", function (Y) {
    // Load ticket data as JSON
    var callback = {
        timeout: 3000,
        on: {
            success: function(x, o) {
                Y.log("Loaded: " + o.responseText.length + " characters", "info");
                try {
                    JSON_DATA = Y.JSON.parse(o.responseText);
                    Y.log("Parsed: " + JSON_DATA.length + " tasks", "info");
                    data_display();
                    // Update KPIs
                    Y.one("#today").set("text", new Date().toDateString());
                    Y.one("#open_tasks").set("text", KPI_nTasks - KPI_nCompleted);
                    Y.one("#wtg_count").set("text", KPI_nWaiting);
                    Y.one("#run_count").set("text", KPI_nRunning);
                    Y.one("#dne_count").set("text", KPI_nCompleted);
                    Y.one("#dly_count").set("text", KPI_nDelayed);
                    Y.one("#err_count").set("text", KPI_nStopped);
                    Y.one("#visible").set("text", KPI_nTasks - KPI_nCompleted);
                    // Reload the page in 5 minutes
                    var handle = Y.later( 1000 * 60 * 5, window, function(){
                        window.location.reload();
                    }, [], false);
                } catch(e) {
                    Y.log("Data parsing error", "error", "callback-success-catch");
                    return;
                }
            },
            failure: function(x, o) {
                Y.log("Data loading failure: " + o.statusText, "error", "callback-failure");
            }
        }
    };
//    Y.io("../cgi-bin/cutover/cutover.cgi", callback);
    Y.io("./data/cutover.json", callback);
});
