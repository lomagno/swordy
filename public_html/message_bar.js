function MessageBar(elementId) {
    var
        m_self = this,
        m_container = $('#' + elementId),
        m_text = m_container.find('.mb-text')
    ;
    
    // Close link
    var m_closeLink = m_container.find('.mb-close-link');    
    m_closeLink.click(function(event) {
        event.preventDefault();
        m_self.close();
    });
    
    // Close the message bar if the user click outside of the message bar
    $(document).mouseup(function (e) {
        if (!m_container.is(e.target) && m_container.has(e.target).length === 0)
            m_self.close();
    });    
    
    this.showMessage = function(text) {       
        m_container.show();
        m_text.text(text);
    };
    
    this.close = function() {
        m_text.text('');
        m_container.hide();
    };
}