/* global Office */

function getBindingProperties(bindingId) {
    var bindingProperties = {};
    
    // Numeric properties
    var numericProperties = ['id', 'decimals'];
    
    // Split the binding ID string
    var bindingIdArray = bindingId.split('.');
    
    var i = 0;
    while (i < bindingIdArray.length-1) {
        var key = bindingIdArray[i];
        var value = bindingIdArray[i+1];
        if ($.inArray(key, numericProperties) !== -1)
            value = +value;
        bindingProperties[key] = value;
        i = i + 2;
    } 
    
    return bindingProperties;
} 

// This function is required for recent version of IE, because
// the Number.isInteger function is not supported
function isInteger(num){
    var numCopy = parseFloat(num);
    return !isNaN(numCopy) && numCopy == numCopy.toFixed();
}

function syncScalarData(bindingId, scalarValue, decimals, onComplete) {
    var text = scalarValue.toFixed(decimals);
    Office.select('bindings#' + bindingId, function() {console.log('pippo errore');}).setDataAsync(text, {asyncContext: bindingId}, function(asyncResult) { // TODO: manage error
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
            onComplete();
        else
            console.error('Error: ' + asyncResult.error.message);
    });               
}

function syncMatrixData(bindingId, matrixData, decimals) {        
    // Create table
    var table = new Office.TableData();
    var rows = matrixData.rows;
    var cols = matrixData.cols;
    var data = matrixData.data;        
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

    // Set table data
    Office.context.document.bindings.getByIdAsync(bindingId, function (asyncResult) {
        console.log('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        asyncResult.value.setDataAsync(table, { coercionType: "table" }, function(asyncResult) {
            if (asyncResult.status === "failed")
                console.log('Error: ' + asyncResult.error.message);
            else
                console.log('Bound data: ' + asyncResult.value);               
        });         
    });
} 