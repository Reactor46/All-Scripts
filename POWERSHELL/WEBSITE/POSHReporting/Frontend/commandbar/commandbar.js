function create_CommandBar() {

    build_CommandBar()

    create_ColumnDropDown();
    
    initialize_commandBar();

    handle_Pagination();

    handle_searchOptions()

    handle_SeachBox()

    Handle_ExportButtons()
}

function build_CommandBar() {

    $("table").each(
        function () {
            var table = $(this)
            var tableID = table.attr("id")
            var commandBar = (` <div id="tablename_commandBar">
                                <div class="ms-CommandBar">
                                    <div class="ms-CommandBar-sideCommands">
                                        <div class="ms-CommandButton">
                                            <button class="ms-CommandButton-button js-Paginate-Select" title="Pick how many entries in table to show">
                                                <span class="ms-CommandButton-icon ms-fontColor-themePrimary">
                                                    <i class="ms-Icon ms-Icon--ViewList"></i>
                                                </span>
                                                <span class="ms-CommandButton-label">View</span>
                                                <span class="ms-CommandButton-dropdownIcon">
                                                    <i class="ms-Icon ms-Icon--ChevronDown"></i>
                                                </span>
                                            </button>
                                            <ul class="ms-ContextualMenu is-opened" id="tablename_paginatecontrol" style="min-width: 100px;">
                                                <li class="ms-ContextualMenu-item">
                                                    <a class="ms-ContextualMenu-link checked js-Paginate-Select" tabindex="1" value="10">10 entries</a>
                                                </li>
                                                <li class="ms-ContextualMenu-item">
                                                    <a class="ms-ContextualMenu-link js-Paginate-Select" tabindex="1" value="25">25 entries</a>
                                                </li>
                                                <li class="ms-ContextualMenu-item">
                                                    <a class="ms-ContextualMenu-link js-Paginate-Select" tabindex="1" value="50">50 entries</a>
                                                </li>
                                                <li class="ms-ContextualMenu-item">
                                                    <a class="ms-ContextualMenu-link js-Paginate-Select" tabindex="1" value="100">100 entries</a>
                                                </li>
                                                <li class="ms-ContextualMenu-item">
                                                    <a class="ms-ContextualMenu-link js-Paginate-Select" tabindex="1" value="-1">All entries</a>
                                                </li>
                                            </ul>
                                        </div>
                                    </div>
                                    <div class="ms-CommandBar-mainArea">
                        
                                        <div id="tablename_search" class="ms-SearchBox ms-SearchBox--commandBar">
                                            <input class="ms-SearchBox-field" type="text" value="" placeholder="Search" data-type="normal">
                                            <label class="ms-SearchBox-label">
                                                <i class="ms-SearchBox-icon ms-Icon ms-Icon--Search"></i>
                                                <span class="ms-SearchBox-text" style="display: none;">Search</span>
                                            </label>
                                            <div class="ms-CommandButton ms-SearchBox-clear ms-CommandButton--noLabel">
                                                <button class="ms-CommandButton-button">
                                                    <span class="ms-CommandButton-icon">
                                                        <i class="ms-Icon ms-Icon--Cancel"></i>
                                                    </span>
                                                    <span class="ms-CommandButton-label"></span>
                                                </button>
                                            </div>
                                            <div class="ms-CommandButton ms-SearchBox-exit ms-CommandButton--noLabel">
                                                <button class="ms-CommandButton-button">
                                                    <span class="ms-CommandButton-icon">
                                                        <i class="ms-Icon ms-Icon--ChromeBack"></i>
                                                    </span>
                                                    <span class="ms-CommandButton-label"></span>
                                                </button>
                                            </div>
                                            <div class="ms-CommandButton ms-SearchBox-filter ms-CommandButton--noLabel">
                                                <button class="ms-CommandButton-button">
                                                    <span class="ms-CommandButton-icon">
                                                        <i class="ms-Icon ms-Icon--Search"></i>
                                                    </span>
                                                    <span class="ms-CommandButton-label"></span>
                                                </button>
                                            </div>
                                        </div>
                        
                                        <div class="ms-CommandButton ms-CommandButton--noLabel column-picker">
                                            <button class="ms-CommandButton-button" title="Pick a table column to search in">
                                                <span class="ms-CommandButton-icon ms-fontColor-themePrimary">
                                                    <i class="ms-Icon ms-Icon--CaretSolidDown"></i>
                                                </span>
                                                <span class="ms-CommandButton-label"></span>
                                            </button>
                                            <ul class="ms-ContextualMenu is-opened" id="tablename_ColumnPicker">
                                                <li class="ms-ContextualMenu-item">
                                                    <a class="ms-ContextualMenu-link column-picker checked" tabindex="1">Any column</a>
                                                </li>
                                            </ul>
                                        </div>
                                        <div class="ms-CommandButton">
                                            <button class="ms-CommandButton-button js-SearchOptions-select">
                                                <span class="ms-CommandButton-icon ms-fontColor-themePrimary">
                                                    <i class="ms-Icon ms-Icon--DocumentSearch"></i>
                                                </span>
                                                <span class="ms-CommandButton-label">Options</span>
                                                <span class="ms-CommandButton-dropdownIcon">
                                                    <i class="ms-Icon ms-Icon--ChevronDown"></i>
                                                </span>
                                            </button>
                                            <ul id="tablename_searchOptions" class="ms-ContextualMenu" style="min-width: 116px;">
                                                <li class="ms-ContextualMenu-item">
                                                    <a class="ms-ContextualMenu-link search-options regex-toggle" tabindex="1">RexEx</a>
                                                </li>
                                            </ul>
                                        </div>
                                        <div class="ms-CommandButton">
                                            <button class="ms-CommandButton-button">
                                                <span class="ms-CommandButton-icon ms-fontColor-themePrimary">
                                                    <i class="ms-Icon ms-Icon--Generate"></i>
                                                </span>
                                                <span class="ms-CommandButton-label">Export</span>
                                                <span class="ms-CommandButton-dropdownIcon">
                                                    <i class="ms-Icon ms-Icon--ChevronDown"></i>
                                                </span>
                                            </button>
                                            <ul id="tablename_Export" class="ms-ContextualMenu" style="min-width: 116px;">
                                                <li class="ms-ContextualMenu-item">   
                                                    <a class="ms-ContextualMenu-link export-excel" tabindex="1">Excel</a>
                                                    <a class="ms-ContextualMenu-link export-csv" tabindex="1">Csv</a>
                                                </li>
                                            </ul>
                                        </div>
                                        <div class="ms-CommandButton ms-CommandBar-overflowButton ms-CommandButton--noLabel">
                                            <button class="ms-CommandButton-button">
                                                <span class="ms-CommandButton-icon">
                                                    <i class="ms-Icon ms-Icon--More"></i>
                                                </span>
                                                <span class="ms-CommandButton-label"></span>
                                            </button>
                                            <ul class="ms-ContextualMenu is-opened ms-ContextualMenu--hasIcons">
                                                <li class="ms-ContextualMenu-item">
                                                    <a class="ms-ContextualMenu-link search-options" tabindex="1">Debug</a>
                                                    <i class="ms-Icon ms-Icon--CodeEdit"></i>
                                                </li>
                                            </ul>
                                        </div>
                                    </div>
                                </div>
                            </div>`).replace(/tablename/g, tableID)

            $(commandBar).insertBefore(table)
        }
    );
}

function DebugCommandBar() {
    var commandBar = $("#tablename_commandBar");
    var wrapperDiv = $("#tablename_wrapper");
    wrapperDiv.prepend(commandBar);
}

var CommandBar;

function initialize_commandBar() {

    overwrite_handleBlur();

    var CommandBarElements = document.querySelectorAll(".ms-CommandBar");
    for (var i = 0; i < CommandBarElements.length; i++) {
        CommandBar = new fabric['CommandBar'](CommandBarElements[i]);
        // Disable drawing overflow
        CommandBar._runOverflow = function() {}
    }

    // This ensures that clear button in search bar clears results
    $(".ms-SearchBox-clear").mousedown(
        function () {
            $(this).parent().find(".ms-SearchBox-field").keyup();
        }
    );
}

// Handle pagination switch
function handle_Pagination() {
    $(".ms-ContextualMenu-link.js-Paginate-Select").click(
        function () {
           
            let paginateButton = $(this);
            let tableID = get_ID(paginateButton)

            $("#" + tableID + "_paginatecontrol").find(".ms-ContextualMenu-link.js-Paginate-Select.checked").removeClass("checked");

            paginateButton.parent()
            paginateButton.toggleClass("checked");
            let value = paginateButton.attr("value")
            

            $("#" + tableID).DataTable().page.len(value).draw();
        }
    );
}


// Send queries to DataTables API via Office UI JS search box
function handle_SeachBox() {
    $(".ms-SearchBox-field").keyup(
        function () {

            var element = $(this);
            let tableID = get_ID(element);
            //let tableID = element.parent(".ms-SearchBox").attr("id").replace("_search", "");
            let table = $("#" + tableID);
            let query = element.val();

            let searchMode = element.next(".ms-SearchBox-label").find(".ms-SearchBox-text");

            var searchType = element.attr("data-type");

            if (searchType == "regex") {
                var regex = true;
                var smart = false
            }
            else {
                var regex = false;
                var smart = true;
            }

            var columnIndex = searchMode.attr("value");

            if (searchMode.text() == "Search") {
                table.DataTable().search(query, regex, smart).draw();
            }
            else {
                table.DataTable().column(columnIndex).search(query, regex, smart).draw();
            }

        }
    );
}


// Overwrites the _handleBlur prototype form search to prevent unfocusing 
function overwrite_handleBlur() {
    SearchBlur = function (e) {
    }

    fabric.SearchBox.prototype._handleBlur = SearchBlur;
}

// Attach event handler to do check mark on options
function handle_searchOptions() {

    //handle regex
    $(".ms-ContextualMenu-link.search-options").click(
        function () {

            var element = $(this);
            element.toggleClass("checked");

            var ID = get_ID(element);

            if (element.hasClass("checked") == true) {
                var searchBox = $("#" + ID + "_search")
                searchBox.find(".ms-SearchBox-field").attr("data-type", "regex");
            }
            else {
                var searchBox = $("#" + ID + "_search")
                searchBox.find(".ms-SearchBox-field").attr("data-type", "normal")
            }
        }
    );
}


// Create collumn select based on Headers
function create_ColumnDropDown() {

    var tables = $("table");

    tables.each(

        function () {
            var table = $(this)
            var tableID = $(this).attr("id")
            var headers = table.find("thead").find("th");

            headers.each(
                function () {

                    var header = $(this);
                    var li = $("<li></li>").addClass("ms-ContextualMenu-item")
                    var a = $("<a></a>").addClass("ms-ContextualMenu-link column-picker")
                        .attr("tabindex", '1')
                        .attr("value", header.index())
                        .text(header.text());

                    li.append(a)

                    var ul = $("#" + tableID + "_ColumnPicker");

                    ul.append(li);
                }
            );
        }
    );

    handle_ColumnPicker();
}


function handle_ColumnPicker() {

    // Change column filter 
    $(".ms-ContextualMenu-link.column-picker").click(
        function () {

            let element = $(this)
            let value = element.text()
            let valueAttr = element.attr("value")

            var ID = get_ID(element);
            var searchBox = $("#" + ID + "_search")

            // Using helper function to trigger cleaning up query
            triggerEvent((searchBox.find(".ms-SearchBox-clear"))[0], 'mousedown')

            if (value == "Any column") {
                value = "Search";
                valueAttr = "";
            }

            searchBox.find(".ms-SearchBox-field").attr("placeholder", value);
            searchBox.find(".ms-SearchBox-text").text(value).attr("value", valueAttr);
        }
    );

    //Attach event handler to do check mark on column picker
    $(".ms-ContextualMenu-link.column-picker").click(
        function () {
            let element = $(this);
            element.parent().parent().find(".ms-ContextualMenu-link").removeClass("checked");
            element.toggleClass("checked");
        }
    );
}


/* Helper functions */



function get_ID(jQueryElement) {
    var id = jQueryElement.attr("id")

    if (id == undefined) {
        id = get_ID(jQueryElement.parent())
    }
    else {
        var n = id.lastIndexOf("_")
        id = id.substring(0, n)
    }

    return id
}

/* Export Button functions */

// Declaring variable for storing name of exported files
var ExportFileName; 
function Handle_ExportButtons()
{

    /* CSV Export trigger */
    $(".ms-ContextualMenu-link.export-csv").click(
        function () {

            var element = $(this);
            var ID = get_ID(element);

            ExportFileName = ID.replace(/-/g,"").replace(/_/g,"");
            $("#" + ID).DataTable().buttons(".dt-button.buttons-csv").trigger();
        }
    );

    /* Excel Export trigger */
    $(".ms-ContextualMenu-link.export-excel").click(
        function () {

            var element = $(this);
            var ID = get_ID(element);
            
            ExportFileName = ID.replace(/-/g,"").replace(/_/g,"");

            $("#" + ID).DataTable().buttons(".dt-button.buttons-excel").trigger();
        }
    );
    
}

