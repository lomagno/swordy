/* global Office, swire */

'use strict';

(function () {   
    var matrixNameTextEdit,
        decimalsTextEdit,
        decimalsErrorMsg,
        matrixNameErrorMsg,
        errorMsg,
        errorMsgText,
        insertMatrixButton,
        stataNameRx = new RegExp(/^[a-zA-Z_][a-zA-Z_0-9]{0,31}$/),
        isMatrixNameValid = false,
        isDecimalsValid = true;
    
    Office.initialize = function (/* reason */) {
        $(document).ready(function () {
            
            // Matrix name text edit
            matrixNameTextEdit = $('#matrixNameTextEdit');
            matrixNameTextEdit.on('input', onMatrixNameTextEditChanged);
            matrixNameErrorMsg = $('#matrixNameErrorMsg');
            
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
            
            // Insert matrix button
            insertMatrixButton = $('#insertMatrixButton');
            insertMatrixButton.click(onInsertMatrixButtonClicked);
            
            // Stata name regular expression
            stataNameRx = new RegExp(/^[a-zA-Z_][a-zA-Z_0-9]{0,31}$/);            
        });
    };   
    
    function onInsertMatrixButtonClicked() {
        errorMsg.hide();
        var matrixName = matrixNameTextEdit.val().trim();
        var decimals = decimalsTextEdit.val().trim();
        
        var request = {
            job: [
                {
                    method: '$getMatrix',
                    args: {
                        name: matrixName
                    }
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
                    showErrorMsg('Not existing matrix');
                    return;
                }
                
                // Matrix data
                var matrixData = response.output[0].output;
                
                // Insert scalar value in Word
                insertMatrix(matrixData, decimals);
            },
            error: function (/* jqXHR, textStatus, errorThrown */) {
                showErrorMsg('Cannot communicate with Stata');
            }
        });        
    }
    
    function insertMatrix(matrixData, decimals) {
        
        var rows = matrixData.rows;
        var cols = matrixData.cols;
        var data = matrixData.data;
        
        // Prepare table
        var table = new Office.TableData();
        table.headers = [];
        table.rows = [];
        var k=0;
        for (var i=0; i<rows; ++i) {
            var row = [];
            for (var j=0; j<cols; ++j) {
                row.push(data[k].toFixed(decimals));
                k++;
            }
            table.rows.push(row);
        }    
        
        Office.context.document.setSelectedDataAsync(table, {coercionType: 'table'}, function (asyncResult) {            
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                var error = asyncResult.error;
                showErrorMsg(error.name + ": " + error.message);                 
            }
        });        
    }
    
    function onMatrixNameTextEditChanged() {
        errorMsg.hide();
        
        // Validate scalar name
        var text = $(this).val().trim();
        if (text === '') {
            isMatrixNameValid = false;
            matrixNameErrorMsg.text('A Stata scalar name is required');
            matrixNameErrorMsg.show();            
        }
        else if (!stataNameRx.test(text)) {
            isMatrixNameValid = false;
            matrixNameErrorMsg.text('Not valid Stata scalar name');
            matrixNameErrorMsg.show();
        }
        else {
            isMatrixNameValid= true;
            matrixNameErrorMsg.hide();
        }
        
        updateInsertMatrixButtonStatus();
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
        else if (!($.isNumeric(text) && isInteger(text))) {
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
        
        updateInsertMatrixButtonStatus();
    }

    function updateInsertMatrixButtonStatus() {
        if (isMatrixNameValid && isDecimalsValid)
            insertMatrixButton.prop('disabled', false);
        else
            insertMatrixButton.prop('disabled', true);
    }
    
    function showErrorMsg(msg) {
        errorMsgText.text(msg);
        errorMsg.show();
    }
    
    // This function is required for recent version of IE, because
    // the Number.isInteger function is not supported
    function isInteger(num){
        var numCopy = parseFloat(num);
        return !isNaN(numCopy) && numCopy == numCopy.toFixed();
    }    
})();