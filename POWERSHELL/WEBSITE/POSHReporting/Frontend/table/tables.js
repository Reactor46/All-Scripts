function toggle_DataTables() {

    // Add table-header and table id to header
    $("h2").each(
        function () {
            var element = $(this);
            var tableID = element.next("table").attr("id")
            element.addClass("table-header").attr("id", tableID + "_header");
        }
    );

    $(".table-header").click(
        function()
        {
            var header = $(this);
            header.toggleClass("tablecollapsed");

            var ID = get_ID(header)
            toggle_Collapse($("#" + ID + "_wrapper"));
        }
    );

    conditional_formating()

    var tables = $("table")
    tables.addClass("ms-Table").attr("width", "100%");

    tables.each(
        function () 
        {
            var table = $(this);
            var headers = table.find("thead").find("th");
            var containsStatus;

            headers.each(
                function () {
                    if ($(this).text() == "Status") {
                        containsStatus = true;
                    }
                }
            );

            if(containsStatus == true)
            {
                table.DataTable(
                    {
                        "aaSorting": [],
                        "responsive": true,
                        "lengthMenu": [[10, 25, 50, -1], [10, 25, 50, "All"]],
                        "dom": 'tip',
                        buttons: [
                            {
                                extend: 'csv',
                                text: 'CSV',
                                fieldSeparator: ';',
                                filename: function (){ return ExportFileName },
                                exportOptions: {
                                    modifier: {
                                        search: 'none'
                                    }
                                }
                            },
                            {
                                extend: 'excel',
                                text: 'Excel',
                                title: '',
                                filename: function (){ return ExportFileName },
                                exportOptions: {
                                    modifier: {
                                        search: 'none'
                                    }
                                }

                            }

                        ],
                        columnDefs: [
                            { responsivePriority: 1, targets: 0 },
                            { responsivePriority: 1, targets: -1 }
                        ],
                    }
                );
            }

            else
            {
                table.DataTable(
                    {
                        "aaSorting": [],
                        "responsive": true,
                        "lengthMenu": [[10, 25, 50, -1], [10, 25, 50, "All"]],
                        "dom": 'tip',
                        buttons: [
                            {
                                extend: 'csv',
                                text: 'CSV',
                                fieldSeparator: ';',
                                filename: function (){ return ExportFileName },
                                exportOptions: {
                                    modifier: {
                                        search: 'none'
                                    }
                                }
                            },
                            {
                                extend: 'excel',
                                text: 'Excel',
                                title: '',
                                filename:  function (){ return ExportFileName },
                                exportOptions: {
                                    modifier: {
                                        search: 'none'
                                    }
                                }

                            }

                        ],
                    }
                );
            }

        }

    );
    var TableElements = document.querySelectorAll(".ms-Table");
    for (var i = 0; i < TableElements.length; i++) {
        new fabric['Table'](TableElements[i]);
    }
}


/* utility function */
function get_TableData(table, ColumnName) {
    var headers = table.find("thead").find("th");
    var index;

    headers.each(
        function () {
            if ($(this).text() == ColumnName) {
                index = $(this).index();
            }
        }
    );

    var allData = table.find("tbody").find("td");

    var selectedData = [];

    allData.each(
        function () {
            if ($(this).index() == index) {
                selectedData.push($(this)[0]);
            }
        }
    )

    return selectedData;
}

/* colors statuses in tables */
function conditional_formating() {
    let tables = $("table");

    tables.each(
        function () {
            let tableData = get_TableData($(this), "Status")

            $(tableData).each(
                function () {
                    if ($(this).text() == "OK") {
                        $(this).addClass("status-ok");
                    }

                    if ($(this).text() == "Warning") {
                        $(this).addClass("status-warning");
                    }

                    if ($(this).text() == "Error") {
                        $(this).addClass("status-error");
                    }
                }
            );

        }
    );
}

/* creates select row effect on table */
$("tbody").find("tr").click(
    function () {
        $(this).siblings().removeClass("rowselected");
        $(this).toggleClass("rowselected");
    }
);

// renders informational icon on table headers
function render_HeadersInfo() {

    $("table").each(
        function () {

            var table = $(this);
            var tableId = $(this).attr("id");
            var header = table.prev("h2");
            var tableData_Status = get_TableData(table, "Status");
            var errorCount = 0;
            var warningCount = 0;

            // Create a elements
            var Icon = $("<span></span>")

            var tableDataArray = tableData_Status.map(data => $(data).text());
            
            // Count events
            tableDataArray.map(
                function (element) {
                    if (element == "Error") {
                        errorCount++
                    }
                    if (element == "Warning") {
                        warningCount++
                    }
                }
            );

            var events = $("<span></span>").addClass("event-container");  

            var TextSpan = $("<span></span>").text(header.text()).addClass("header-text").attr('title',header.text());
            var textElement = $("<span></span>").append(TextSpan)

            if(errorCount !== 0)
            {
                var errors = $("<span></span>").text(errorCount).addClass("header-event error");
                events.append(errors);
            }
            
            if(warningCount !== 0)
            {
               var warnings = $("<span></span>").text(warningCount).addClass("header-event warning");
               events.append(warnings);
            }
            if(table_hasHeader(table, "Status") && errorCount == 0 && warningCount == 0)
            {
                var ok = $("<span></span>").addClass("header-event ok");
                events.append(ok);
            }
            
            textElement.append(events)
            header.text("").append(textElement)
        }
    );

}

function table_hasHeader(table, header)
{
    var headers = $(table).find("th");
    var hasHeader = false;

    headers.each(
        function()
        {
            if($(this).text() == header)
            {
                hasHeader = true;
            }
        }
    );

    return hasHeader
}