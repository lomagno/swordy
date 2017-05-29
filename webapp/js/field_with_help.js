function FieldWithHelp(elementId) {
    // Container
    var container = $('#' + elementId);
    
    // Help message bar
    var helpMessageBar = new StaticMessageBar({
        element: container.find('.helpMessageBar'),
        onClose: function() {
            helpIcon.show();
        }
    });    

    // Help icon
    var helpIcon = container.find('.helpIcon');
    helpIcon.click(function() {
        helpIcon.hide();
        helpMessageBar.show();
    });     
}