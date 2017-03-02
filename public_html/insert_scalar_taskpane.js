/* global Office, swire */

'use strict';

(function () {   
    var scalarNameTextEdit, decimalsTextEdit;
    
    Office.initialize = function (reason) {
        $(document).ready(function () {
            scalarNameTextEdit = $('#scalarNameTextEdit');
            decimalsTextEdit = $('#decimalsTextEdit');
            
            $('#insertScalarButton').click(retrieveScalar);
        });
    };       
    
    
    
    function retrieveScalar() {
        var scalarName = scalarNameTextEdit.val();
        var decimals = decimalsTextEdit.val();
        
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
                var response = swire.decode(data);
                var scalarValue = response.output[0].output;
                insertNumber(scalarValue, decimals);
            },
            error: function () {
                console.log('network error');
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
})();