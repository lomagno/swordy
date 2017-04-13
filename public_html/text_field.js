/* global fabric */

/*
 * pars:
 * - elementId
 * - value
 * - validators
 * - errorId
 * - onErrorStatusChanged
 */
function TextField(pars) {
    this.getValue = function() {
        return m_textInput.val();
    };
    
    this.setValue = function(text) {
        m_textInput.val(text);
    };
    
    this.show = function() {
        m_container.show();
    };
    
    this.hide = function() {
        m_container.hide();
    };
    
    this.setLabel = function(text) {
        m_label.text(text);
    };
    
    function onTextInputChanged() {
        var text = $(this).val().trim();
        var validatorReport;
        var errorId = null;
        for (var i=0; i<m_validators.length; ++i) {
            validatorReport = m_validators[i](text);
            if (!validatorReport.isValid) {
                errorId = i;
                break;
            }
        }
        
        if (errorId !== m_errorId ) {
            m_errorId = errorId;
            if (m_errorId === null) {
                m_errorMessage.text('');
                m_errorMessage.hide();
            }
            else {
                m_errorMessage.text(validatorReport.errorMessage);
                m_errorMessage.show();                
            }
            pars.onErrorStatusChanged(m_errorId);
        }    
    }
    
    var
        m_container,
        m_label,
        m_textInput,
        m_errorMessage,
        m_validators,
        m_errorId;
    ;
    
    // Container
    m_container = $('#' + pars.elementId);       
    new fabric['TextField'](m_container[0]);   
    
    // Label
    m_label = m_container.find('.ms-Label');
    
    // Input element
    m_textInput = m_container.find('.ms-TextField-field');
    m_textInput.val(pars.value);
    m_textInput.on('input', onTextInputChanged);    
    
    // Error message
    m_errorMessage = m_container.find('.textInputErrorMessage');
    
    // Validation
    m_validators = pars.validators;
    if (pars.errorId === undefined)
        m_errorId = null;
    else
        m_errorId = pars.errorId;
}