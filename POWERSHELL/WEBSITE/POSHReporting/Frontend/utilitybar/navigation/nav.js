function render_HeadersMenu() {

    var navigationDiv = $("#Navigation").find(".utilitybar-category-content");

    $("table").each(
        function () {

            var table = $(this);
            var tableId = $(this).attr("id");
            var headerText = table.prev("h2").text();
            var tableData_Status = get_TableData(table, "Status");

            // Create a elements
            var a = $("<a></a>").addClass("ms-bgColor-neutralSecondary--hover").attr("data-target", "#" + tableId + "_header");
            var spanIcon = $("<span></span>").addClass("nav-icon");
            var spanText = $("<span></span>").addClass("nav-text").html(headerText);
            var Icon = $("<span></span>")

            var tableDataArray = tableData_Status.map(data => $(data).text());

            //  Check what icon shoud be assigned based on the properties in Status column
            if (tableDataArray.length == 0) {
                Icon.addClass("icon-info");
            }
            else if (tableDataArray.includes("Error")) {
                Icon.addClass("icon-error");
            }
            else if (tableDataArray.includes("Warning")) {
                Icon.addClass("icon-warning");
            }
            else if (tableDataArray.includes("OK")) {
                Icon.addClass("icon-ok");
            }

            // Glue elements
            spanIcon.html(Icon);
            a.append(spanIcon).append(spanText);

            // Append it to navigationDiv to create navigation
            navigationDiv.append(a);
        }
    );

    /* event handler to scroll table into view */
    $(".utilitybar-category-content").find("a").click(
        function () {

            var targetID = $(this).attr("data-target")

            // if view mode is tablet or smaller, close utilitybar nad open table
            if ($(window).innerWidth() <= 1024) {
                $(".topbar-button").click();
            }

            // if table is collapsed, uncollapse it
            if ($(targetID).hasClass("tablecollapsed") == true) {
                $(targetID).click();

                // scroll into view after anmiation
                setTimeout(function(){scroll_IntoView(targetID)}, 301)
            }
            else
            {
                // scroll into view
                scroll_IntoView(targetID)
            }

        }
    );
}

function scroll_IntoView(id)
{
    $(id)[0].scrollIntoView(true);

    // fix position due to fixed top bar
    var topbarHeight = $("#top-bar").height();
    var scrolledY = window.scrollY;
    if (scrolledY) 
    {
        window.scroll(0, scrolledY - topbarHeight);
    }
}