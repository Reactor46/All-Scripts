$(document).ready(
    function () {

        toggle_DataTables();

        create_CommandBar()
    }
);

function create_CommandBar() {

    //add when switching to prod
    //build_CommandBar()

    create_ColumnDropDown();

    //delete when switching to prod
    DebugCommandBar();
    
    initialize_commandBar();

    handle_Pagination();

    handle_searchOptions()

    handle_SeachBox()
}


function DebugCommandBar() {
    var commandBar = $("#tablename_commandBar");
    var wrapperDiv = $("#tablename_wrapper");
    wrapperDiv.prepend(commandBar);
}


function initialize_commandBar() {

    overwrite_handleBlur();

    var CommandBarElements = document.querySelectorAll(".ms-CommandBar");
    for (var i = 0; i < CommandBarElements.length; i++) {
        new fabric['CommandBar'](CommandBarElements[i]);
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
            $(".ms-ContextualMenu-link.js-Paginate-Select").removeClass("checked");
            let paginateButton = $(this);
            paginateButton.toggleClass("checked");
            let value = paginateButton.attr("value")
            let tableID = get_ID(paginateButton)

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
function triggerEvent(el, type) {
    if ('createEvent' in document) {
        // modern browsers, IE9+
        var e = document.createEvent('HTMLEvents');
        e.initEvent(type, false, true);
        el.dispatchEvent(e);
    } else {
        // IE 8
        var e = document.createEventObject();
        e.eventType = type;
        el.fireEvent('on' + e.eventType, e);
    }
}


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