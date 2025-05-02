// event handler on clicking filter buttons
handle_Filters2()


function handle_filters() {
    $(".filter-button").click(
        function () {
            // get values
            buttons = $(".filter-button")
            var filters = [];

            buttons.each(
                function () {
                    button = $(this);

                    if (button.hasClass("active") == true) {
                        filters.push(button.find(".ms-Button-label").text());
                    }
                }

            );

            // if query not empty enable regex else disable regex
            var regexbtn = $("#toggle-regex")

            if (filters.length <= 1) {
                if (regexbtn.hasClass("active") == true) {
                    regexbtn.click();
                }
            }
            else if (regexbtn.hasClass("active") == false) {
                regexbtn.click();
            }

            //build regex syntax
            var regexQuery = filters.join("|");

            //add query to search and execute
            $(".ms-SearchBox-field").val(regexQuery).keyup()
        }
    );
}


function handle_Filters2() {
    $(".filter-button").click(
        function () 
        {
            buttons = $(".filter-button")
            var filters = [];

            buttons.each(
                function () {
                    button = $(this);

                    if (button.hasClass("active") == true) {
                        filters.push(button.find(".ms-Button-label").text());
                    }
                }

            );

            var regexQuery = filters.join("|");

            $("table").each(
                function()
                {
                    var table = $(this);
                    var hasStatus;
                    var statusIndex;

                    table.find("th").each(
                        function()
                        {
                            header = $(this);
                            if (header.text() == "Status")
                            {
                                hasStatus = true;
                                statusIndex = header.index();
                            }
                        }
                    );

                    if(hasStatus == true)
                    {
                        $(this).DataTable().column(statusIndex).search(regexQuery, true, false).draw();
                    }
                    
                }
            );
        }
    );
}