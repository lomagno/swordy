function MessageBar(elementId) {
    var
        m_self = this,
        m_container = $('#' + elementId),
        m_content = m_container.find('.mb-content'),
        m_lastList = null
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
    
    this.show = function() {
        m_container.show();
    };        
    
    this.close = function() {
        m_self.reset();
        m_container.hide();
    };
    
    this.appendParagraph = function(text) {
        var paragraph = $('<p></p>');
        paragraph.text(text);
        m_content.append(paragraph);
    };
    
    this.appendList = function() {
        m_lastList = $('<ul></ul>');
        m_content.append(m_lastList);
    };
    
    this.appendListItem = function(text) {
        if (m_lastList === null)
            return;
        
        var listItem = $('<li></li>');
        listItem.text(text);
        m_lastList.append(listItem);
    };
    
    this.showMessage = function(text) {
        m_self.reset();
        m_self.appendParagraph(text);
        m_self.show();
    };
    
    this.reset = function() {
        m_content.empty();
        m_lastList = null;        
    };
}