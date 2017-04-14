/* global fabric */

/*
 * pars:
 * - elementId
 * - validators
 * - onErrorStatusChanged
 */
function TextField(pars) {
    this.getValue = function() {
        return m_textInput.val();
    };
    
    this.setValue = function(text) {
        m_textInput.val(text);
        validate();
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
        console.log('onTextInputChanged()');
        validate();
    }
    
    function validate() {       
        var text = m_textInput.val();
        
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
    m_textInput.on('input', onTextInputChanged);
    m_textInput.val(pars.value);
    
    // Error message
    m_errorMessage = m_container.find('.textInputErrorMessage');
    
    // Validation
    m_validators = pars.validators;
    m_errorId = null;
}