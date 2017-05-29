/*
 * args:
 * - element (can be a element ID or a JQuery object)
 * - onClose
 */
function StaticMessageBar(pars) {
    var
        m_self = this,
        m_container = typeof pars.element === 'string' ? $('#' + pars.element) : pars.element
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
        m_container.hide();
        pars.onClose();
    };    
}