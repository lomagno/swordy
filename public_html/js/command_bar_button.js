/*
 * pars:
 * - elementId
 * - onClick
 */
function CommandBarButton(pars) {
    this.setEnabled = function(status) {
        if (status) {
            m_iconContainer.removeClass(M_DISABLED_ICON_CLASS);
            m_iconContainer.addClass(M_ENABLED_ICON_CLASS);            
        }
        else {
            m_iconContainer.removeClass(M_ENABLED_ICON_CLASS);
            m_iconContainer.addClass(M_DISABLED_ICON_CLASS);            
        }
        
        m_button.prop('disabled', !status);
    };
    
    var
        M_ENABLED_ICON_CLASS = 'ms-fontColor-themePrimary',
        M_DISABLED_ICON_CLASS = 'ms-fontColor-neutralSecondary',
        m_container,
        m_button,
        m_iconContainer
    ;
    
    m_container = $('#' + pars.elementId);
    m_button = m_container.find('.ms-CommandButton-button');    
    m_button.click(pars.onClick);
    m_iconContainer = m_container.find('.ms-CommandButton-icon');
}
