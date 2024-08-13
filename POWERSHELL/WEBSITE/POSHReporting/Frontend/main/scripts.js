$(document).ready(() => {


    render_HeadersMenu();
    render_HeadersInfo();

    toggle_DataTables();

    create_CommandBar();

    collapse_tables();

    release_overlay();

    handle_windowReajust();
});

function release_overlay() {
    if ($(window).innerWidth() <= 500) {
        setTimeout(
            function () {
                $("#main-content").removeClass("loading");
            }, 300
        );
    }
    else {
        $("#main-content").removeClass("loading");
    }
}

$('.topbar-button').click(
    function () {
        $("body").toggleClass("inactive");
        $('#wrapper').toggleClass('menuToggled');
    }
);

$('.utilitybar-category-header').click(
    function () {
        $(this).toggleClass("open");
        let element = $(this).parent().find(".utilitybar-category-content");
        toggle_Collapse(element);
    }
);

$(".ms-Button--menu").click(
    function () {
        $(this).toggleClass("active ms-Button--primary");
    }
);


/* Creates animation for navigation */
function toggle_Collapse(element) {

    var elementHeigth = element.height();

    if (element.css("display") == "none") {
        element.toggleClass("animate uncollapsing collapsed");
        element.toggle();
        element.toggleClass("collapsed")
        element.css("height", elementHeigth);

        setTimeout(
            function () {
                element.css("height", "");
                element.toggleClass("animate uncollapsing");
            }
            , 300);
    }

    else {
        element.toggleClass("animate collapsing");

        element.height(elementHeigth);
        element.height("0");

        setTimeout(function () {
            element.toggle();
            element.height("");
            element.toggleClass("animate collapsing");
        }, 300);
    }
}

var ToggleElements = document.querySelectorAll(".ms-Toggle");
for (var i = 0; i < ToggleElements.length; i++) {
    new fabric['Toggle'](ToggleElements[i]);
}


/* Helper function */
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

function collapse_tables() {
    if ($(window).innerWidth() <= 500) {
        $(".table-header").click();
    }
}

// reajust tables if orientation is changed

function handle_windowReajust()
{
    function adjust_Tables()
    {
        $('table').DataTable()
            .columns.adjust()
            .responsive.recalc();
    }

    $(window).resize(adjust_Tables())
             .on( "orientationchange", adjust_Tables());
}