/* global fabric */

/*
 * pars:
 * - elementId
 * - validators
 */
function TextEdit(pars) {
    this.show = function() {
        m_container.show();
    };
    
    this.hide = function() {
        m_container.hide();
    };
    
    function onTextInputChanged() {
        var text = $(this).val().trim();
        for (var i in m_validators) {
            var validatorReport = m_validators[i](text);
            if (validatorReport.isValid) {
                m_currentErrorId = -1;
                m_errorMessage.text('');
                m_errorMessage.hide();
            }
            else {
                var oldCurrentErrorId = m_currentErrorId;
                if (i !== oldCurrentErrorId) {
                    m_currentErrorId = i;
                    m_errorMessage.text(validatorReport.errorMessage);
                    m_errorMessage.show();
                }
            }
        }
    }
    
    var
        m_container,
        m_textInput,
        m_errorMessage,
        m_validators,
        m_currentErrorId;
    ;
    
    // Container
    m_container = $('#' + pars.elementId);       
    new fabric['TextField'](m_container[0]);
    
    // Input element
    m_textInput = m_container.find('.ms-TextField-field');
    m_textInput.on('input', onTextInputChanged);
    
    // Error message
    m_errorMessage = m_container.find('.textInputErrorMessage');
    
    // Validation
    m_validators = pars.validators;
    m_currentErrorId = -1;
}