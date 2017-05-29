/* global Office, swire, ButtonElements, fabric */

'use strict';

(function () {   
    var m_scalarNameTextField,
        m_decimalsTextField,
        m_successMessageBar,
        m_errorMessageBar,
        m_insertScalarButton,
        m_stataNameRx = new RegExp(/^[a-zA-Z_][a-zA-Z_0-9]{0,31}$/);
    
    Office.initialize = function (/* reason */) {
        $(document).ready(function () { 
            // Insert scalar button
            m_insertScalarButton = $('#insertScalarButton');
            new fabric['Button'](m_insertScalarButton[0], onInsertScalarButtonClicked);
            
            // Scalar name text field
            m_scalarNameTextField = new TextField({
                elementId: 'scalarNameTextField',
                validators: [
                    function (text) {
                        if (text === '')
                            return {
                                isValid: false,
                                errorMessage: 'A Stata scalar name is required'
                            };
                        else
                            return {isValid: true};
                    },
                    function (text) {
                        if (!m_stataNameRx.test(text))
                            return {
                                isValid: false,
                                errorMessage: 'Not valid Scalar name'
                            };
                        else
                            return {isValid: true};
                    }
                ],
                onErrorStatusChanged: updateInsertScalarButtonStatus
            });
            m_scalarNameTextField.setValue('', false);
            new FieldWithHelp('scalarNameTextField');           

            // Decimals text field
            m_decimalsTextField = new TextField({
                elementId: 'decimalsTextField',
                validators: [
                    function (text) {
                        if (text === '')
                            return {
                                isValid: false,
                                errorMessage: 'Decimals must be set'
                            };
                        else
                            return {isValid: true};
                    },
                    function (text) { // TODO: should this validator be common?
                        if (!($.isNumeric(text) && isInteger(text)))
                            return {
                                isValid: false,
                                errorMessage: 'An integer number must be entered'
                            };
                        else
                            return {isValid: true};
                    },
                    function (text) {
                        if (+text < 0 || +text > 20) {
                            return {
                                isValid: false,
                                errorMessage: 'An integer value between 0 and 20 is required'
                            };
                        } else
                            return {isValid: true};
                    }
                ],
                onErrorStatusChanged: updateInsertScalarButtonStatus
            });
            m_decimalsTextField.setValue('3');
            new FieldWithHelp('decimalsTextField');
            
            // Success message bar
            m_successMessageBar = new MessageBar('successMessageBar');
            
            // Error message bar
            m_errorMessageBar = new MessageBar('errorMessageBar');                       
        });
    };   
    
    function onInsertScalarButtonClicked() {
        var scalarName = m_scalarNameTextField.getValue().trim();
        var decimals = m_decimalsTextField.getValue().trim();        
        
        var request = {
            job: [
                {
                    method: 'com.stata.sfi.Scalar.getValue',
                    args: [scalarName]
                }
            ]
        };
        
        $.ajax({
            url: 'https://localhost:50000',
            data: swire.encode(request),
            method: "POST",
            success: function (data) {
                // Decode response
                var response = swire.decode(data);
                
                // Check errors
                if (response.status !== 'ok') {
                    m_errorMessageBar.showMessage('SWire error. Please check that you are using SWire verson 0.2 or later.');
                    return;
                }                
                if (response.output[0].status !== 'ok') {
                    m_errorMessageBar.showMessage('SWire error. Please check that you are using SWire verson 0.2 or later.');
                    return;                    
                }                
                if (response.output[0].output === null) {
                    m_errorMessageBar.showMessage('Not existing scalar');
                    return;
                }
                
                // Scalar value
                var scalarValue = response.output[0].output;
                
                // Insert scalar value in Word
                var text = scalarValue.toFixed(decimals);
                Office.context.document.setSelectedDataAsync(
                    text,
                    {coercionType: 'text'},
                    function (asyncResult) {            
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
                            m_successMessageBar.showMessage('The scalar was correctly inserted.');
                        else
                            m_errorMessageBar.showMessage('Cannot insert the scalar.');
                }); 
            },
            error: function (jqXHR, textStatus, errorThrown) {
                m_errorMessageBar.showMessage('Cannot connect to SWire.');
                foo1 = jqXHR;
                console.log(jqXHR);
                console.log(textStatus);
                console.log(errorThrown);
            }
        });        
    }          
    
    function updateInsertScalarButtonStatus() {
        if (m_scalarNameTextField.isValid() && m_decimalsTextField.isValid())
            m_insertScalarButton.prop('disabled', false);
        else
            m_insertScalarButton.prop('disabled', true);
    }  
})();