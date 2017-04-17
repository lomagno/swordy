/* global Office, swire, fabric */

'use strict';

(function () {   
    var m_matrixNameTextField,
        m_decimalsTextField,
        m_successMessageBar,
        m_errorMessageBar,
        m_insertMatrixButton,
        m_stataNameRx = new RegExp(/^[a-zA-Z_][a-zA-Z_0-9]{0,31}$/);
    
    Office.initialize = function (/* reason */) {
        $(document).ready(function () {
            // Insert matrix button
            m_insertMatrixButton = $('#insertMatrixButton');
            new fabric['Button'](m_insertMatrixButton[0], onInsertMatrixButtonClicked);            
            
            // Validators
            var integerNumberValidator = function (text) {
                if (!($.isNumeric(text) && isInteger(text)))
                    return {
                        isValid: false,
                        errorMessage: 'An integer number must be entered'
                    };
                else
                    return {isValid: true};
            };

            // Data name text field
            m_matrixNameTextField = new TextField({
                elementId: 'matrixNameTextField',
                validators: [
                    function (text) {
                        if (text === '')
                            return {
                                isValid: false,
                                errorMessage: 'A Stata matrix name is required'
                            };
                        else
                            return {isValid: true};
                    },
                    function (text) {
                        if (!m_stataNameRx.test(text))
                            return {
                                isValid: false,
                                errorMessage: 'Not valid matrix Stata data name'
                            };
                        else
                            return {isValid: true};
                    }
                ],
                onErrorStatusChanged: updateInsertMatrixButtonStatus
            });
            m_matrixNameTextField.setValue('', false);

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
                    integerNumberValidator,
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
                onErrorStatusChanged: updateInsertMatrixButtonStatus
            });
            m_decimalsTextField.setValue('3');
            
            // Success message bar
            m_successMessageBar = new MessageBar('successMessageBar');
            
            // Error message bar
            m_errorMessageBar = new MessageBar('errorMessageBar');          
        });
    };   
    
    function onInsertMatrixButtonClicked() {        
        var matrixName = m_matrixNameTextField.getValue().trim();
        var decimals = m_decimalsTextField.getValue().trim();        
        
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
                    m_errorMessageBar.showMessage('SWire error. Please check that you are using SWire verson 0.2 or later.');
                    return;
                }                
                if (response.output[0].status !== 'ok') {
                    m_errorMessageBar.showMessage('SWire error. Please check that you are using SWire verson 0.2 or later.');
                    return;                    
                }                
                if (response.output[0].output === null) {
                    m_errorMessageBar.showMessage('Not existing matrix');
                    return;
                }
                
                // Matrix data
                var matrixData = response.output[0].output;
                
                // Insert matrix in Word
                insertMatrix({
                    matrixData: matrixData,
                    decimals: decimals,
                    onComplete: function(asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
                            m_successMessageBar.showMessage('The matrix was correctly inserted.');
                        else
                            m_errorMessageBar.showMessage('Cannot insert the matrix.');                        
                    }           
                });
            },
            error: function (/* jqXHR, textStatus, errorThrown */) {
                m_errorMessageBar.showMessage('Cannot connect to SWire.');
            }
        });        
    }
    
    /*
     * args:
     * - matrixData
     * - decimals
     * - onComplete
     */
    function insertMatrix(args) {
        // Arguments
        var
            matrixData = args.matrixData,
            decimals = args.decimals,
            onComplete = args.onComplete;
        
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
                
        Office.context.document.setSelectedDataAsync(
            table,
            {
                coercionType: 'table'
            },
            function (asyncResult) {            
                onComplete(asyncResult);
            }
        );        
    }    
    
    function updateInsertMatrixButtonStatus(errorId) {
        if (errorId === null)
            m_insertMatrixButton.prop('disabled', false);
        else
            m_insertMatrixButton.prop('disabled', true);
    }   
})();