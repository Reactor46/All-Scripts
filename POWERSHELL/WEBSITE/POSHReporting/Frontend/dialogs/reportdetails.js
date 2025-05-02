(function() {
    var aboutDiv = $("#details_Dialog");
    var button = $("#details_Button")[0];
    var dialog = aboutDiv.find(".ms-Dialog--lgHeader")[0];
    var actionButtonElements = aboutDiv[0].querySelectorAll(".ms-Dialog-action");
    var actionButtonComponents = [];
    // Wire up the dialog
    var dialogComponent = new fabric['Dialog'](dialog);
    // Wire up the buttons
    for (var i = 0; i < actionButtonElements.length; i++) {
      actionButtonComponents[i] = new fabric['Button'](actionButtonElements[i], actionHandler);
    }
    //unbind event handler for click
    $(button).unbind("click")

    // When clicking the button, open the dialog
    button.onclick = function() {
      openDialog(dialog);
    };
    function actionHandler(event) {
      
    }
    function openDialog(dialog) {
      // Open the dialog
      dialogComponent.open();
    }
  }());