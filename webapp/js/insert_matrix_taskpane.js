/* global Office, swire, fabric */

'use strict';

(function () {   
    var m_matrixNameTextField,
        m_decimalsTextField,
        m_successMessageBar,
        m_errorMessageBar,
        m_insertMatrixButton,
        m_stataNameRx = new RegExp(/^[a-zA-Z_][a-zA-Z_0-9]{0,31}$/),
        m_decimalsListRx = new RegExp(/^\s*([0-9]|1[0-9]|20)\s*(,\s*([0-9]|1[0-9]|20))*\s*$/);
    
    Office.initialize = function (/* reason */) {
        $(document).ready(function () {
            // Insert matrix button
            m_insertMatrixButton = $('#insertMatrixButton');
            new fabric['Button'](m_insertMatrixButton[0], onInsertMatrixButtonClicked);

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
                    function (text) {
                        if (!m_decimalsListRx.test(text))
                            return {
                                isValid: false,
                                errorMessage:
                                    'Not valid decimals list: must be a list of integers between 0 and 20 separated by commas.'
                                    + ' Example: 4, 1, 5.'
                            };
                        else
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
        var decimals = splitCommaSeparatedValues(m_decimalsTextField.getValue());        
        
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
     * - decimals (numeric vector)
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
        
        // Fill missing decimals
        var missingDecimals = cols - decimals.length;
        for (var i=0; i<missingDecimals; ++i)
            decimals.push(decimals[decimals.length-1]);        
        
        // Prepare table
        var table = new Office.TableData();
        table.headers = [];
        table.rows = [];
        var k=0;
        for (var i=0; i<rows; ++i) {
            var row = [];
            for (var j=0; j<cols; ++j) {
                row.push(data[k].toFixed(decimals[j]));
                k++;
            }
            table.rows.push(row);
        }    
                
        // Insert table in the Word document
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
    
    function updateInsertMatrixButtonStatus() {
        if (m_matrixNameTextField.isValid() && m_decimalsTextField.isValid())
            m_insertMatrixButton.prop('disabled', false);
        else
            m_insertMatrixButton.prop('disabled', true);
    }   
})();