/* global Office, swire */

'use strict';

(function () {   
    var scalarNameTextEdit,
        decimalsTextEdit,
        decimalsErrorMsg,
        scalarNameErrorMsg,
        errorMsg,
        errorMsgText,
        insertScalarButton,
        stataNameRx = new RegExp(/^[a-zA-Z_][a-zA-Z_0-9]{0,31}$/),
        isScalarNameValid = false,
        isDecimalsValid = true;
    
    Office.initialize = function (/* reason */) {
        $(document).ready(function () {                        
            // Scalar name text edit
            scalarNameTextEdit = $('#scalarNameTextEdit');
            scalarNameTextEdit.on('input', onScalarNameTextEditChanged);
            scalarNameErrorMsg = $('#scalarNameErrorMsg');
            
            // Decimals text edit
            decimalsTextEdit = $('#decimalsTextEdit');
            decimalsTextEdit.on('input', onDecimalsTextEditChanged);
            decimalsErrorMsg = $('#decimalsErrorMsg');
            
            // Error message
            errorMsg = $('#error-msg');            
            $('#error-msg-close-link').click(function(event) {
                event.preventDefault();
                errorMsg.hide();
            });
            errorMsgText = $('#error-msg-text');
            
            // Insert scalar button
            insertScalarButton = $('#insertScalarButton');
            $('#insertScalarButton').click(onInsertScalarButtonClicked);
            
            // Stata name regular expression
            stataNameRx = new RegExp(/^[a-zA-Z_][a-zA-Z_0-9]{0,31}$/);            
        });
    };   
    
    function onInsertScalarButtonClicked() {
        errorMsg.hide();
        var scalarName = scalarNameTextEdit.val().trim();
        var decimals = decimalsTextEdit.val().trim();
        
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
                    showErrorMsg('SWire returned an error');
                    return;
                }                
                if (response.output[0].status !== 'ok') {
                    showErrorMsg('SWire returned an error');
                    return;                    
                }                
                if (response.output[0].output === null) {
                    showErrorMsg('Not existing scalar');
                    return;
                }
                
                // Scalar value
                var scalarValue = response.output[0].output;
                
                // Insert scalar value in Word
                insertNumber(scalarValue, decimals);
            },
            error: function (/* jqXHR, textStatus, errorThrown */) {
                showErrorMsg('Cannot communicate with Stata');
            }
        });        
    }
    
    function insertNumber(number, decimals) {
        var text = number.toFixed(decimals);
        Office.context.document.setSelectedDataAsync(text, {coercionType: 'text'}, function (asyncResult) {            
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                var error = asyncResult.error;
                console.log(error.name + ": " + error.message);                 
            }
        });        
    }
    
    function onScalarNameTextEditChanged() {
        errorMsg.hide();
        
        // Validate scalar name
        var text = $(this).val().trim();
        if (text === '') {
            isScalarNameValid = false;
            scalarNameErrorMsg.text('A Stata scalar name is required');
            scalarNameErrorMsg.show();            
        }
        else if (!stataNameRx.test(text)) {
            isScalarNameValid = false;
            scalarNameErrorMsg.text('Not valid Stata scalar name');
            scalarNameErrorMsg.show();
        }
        else {
            isScalarNameValid= true;
            scalarNameErrorMsg.hide();
        }
        
        updateInsertScalarButtonStatus();
    }
    
    function onDecimalsTextEditChanged() {
        errorMsg.hide();
        
        // Validate decimals
        var text = $(this).val().trim();
        if (text === '') {
            isDecimalsValid = false;
            decimalsErrorMsg.text('Decimals must be set');
            decimalsErrorMsg.show();
        }
        else if (!($.isNumeric(text) && Number.isInteger(+text))) {
            isDecimalsValid = false;
            decimalsErrorMsg.text('An integer number must be entered');
            decimalsErrorMsg.show();
        }
        else if (+text < 0 || +text > 20) {
            isDecimalsValid = false;
            decimalsErrorMsg.text('A integer value between 0 and 20 is required');
            decimalsErrorMsg.show();
        }
        else {
            isDecimalsValid = true;
            decimalsErrorMsg.hide();
        }
        
        updateInsertScalarButtonStatus();
    }

    function updateInsertScalarButtonStatus() {
        if (isScalarNameValid && isDecimalsValid)
            insertScalarButton.prop('disabled', false);
        else
            insertScalarButton.prop('disabled', true);
    }
    
    function showErrorMsg(msg) {
        errorMsgText.text(msg);
        errorMsg.show();
    }
})();