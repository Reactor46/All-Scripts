handle_RegEx();
handle_Expand();
handle_Collapse();
handle_View();

function handle_RegEx() {
    $("#toggle-regex").click(
        function () {

            var active = $(this).hasClass("active");


            $(".ms-CommandButton-button.js-SearchOptions-select").click();

            $(".search-options.regex-toggle").each(
                function () {
                    var regexBtn = $(this);

                    if (active == true && regexBtn.hasClass("checked") == false) {
                        regexBtn.click()
                    }
                    else if (active == false && regexBtn.hasClass("checked") == true) {
                        regexBtn.click()
                    }
                }
            );
        }
    );

}

function handle_Expand() {

    // Workaround for default button behavior! Needs to be refactor!
    $("#expand-tables").unbind("click");

    $("#expand-tables").click(
        function () {
            
            $(".table-header").each(
                function()
                {
                    var header = $(this);
                    if(header.hasClass("tablecollapsed") == true)
                    {
                        this.click();
                    }
                }
            );
        }
    );
}

function handle_Collapse() {

    // Workaround for default button behavior! Needs to be refactor!
    $("#collapse-tables").unbind("click");

    $("#collapse-tables").click(
        function () {
            
            $(".table-header").each(
                function()
                {
                    var header = $(this);
                    if(header.hasClass("tablecollapsed") == false)
                    {
                        this.click();
                    }
                }
            );
        }
    );
}

function handle_View() {

    // Workaround for default button behavior! Needs to be refactor!
    $("#view-10").unbind("click");
    $("#view-all").unbind("click");

    $("#view-10").click(
        function () 
        {
            var button = $(this);

            $(".ms-CommandButton-button.js-Paginate-Select").click();

            $(".ms-ContextualMenu-link.js-Paginate-Select").each(
                function () {
                    var paginate = $(this);

                    if(paginate.attr("value") == 10)
                    {
                        paginate.click();
                    }
                   
                }
            );
            
        }
    );

    $("#view-all").click(
        function () 
        {
            var button = $(this);

            $(".ms-CommandButton-button.js-Paginate-Select").click();

            $(".ms-ContextualMenu-link.js-Paginate-Select").each(
                function () {

                    var paginate = $(this);
                    
                    if(paginate.attr("value") == -1)
                    {
                        paginate.click();
                    }
                }
            );
            
        }
    );
}
