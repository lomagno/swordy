/* global Office */

function getBindingProperties(bindingId) {
    var bindingProperties = {};
    
    // Numeric properties
    var numericProperties = ['id', 'startingRow' , 'startingColumn', 'decimals'];
    
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

/*
 * args:
 * - bindingId
 * - scalarValue
 * - report
 * - onComplete
 */
function syncScalarData(args) {
    var bindingId = args.bindingId;
    var scalarValue = args.scalarValue;
    var report = args.report;
    var onComplete = args.onComplete;  
    
    var bindingProperties = getBindingProperties(bindingId);    
    var text = scalarValue.toFixed(bindingProperties.decimals);
    
    Office.context.document.bindings.getByIdAsync(bindingId, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            var binding = asyncResult.value;
            binding.setDataAsync(
                text,
                {coercionType: "text"},
                function(asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded)      
                        report.syncOk.push(bindingId);               
                    else {
                        report.syncNotOk.push(bindingId);
                        report.syncNotOkErrorCodes.push(asyncResult.error.code);
                        console.log(asyncResult.error.message);
                    }

                    // Execute callback
                    if (report.syncOk.length + report.syncNotOk.length === report.count)
                        onComplete(report);                
                }
            );            
        }
        else {
            report.syncNotOk.push(bindingId);
            report.syncNotOkErrorCodes.push(asyncResult.error.code);
            console.log(asyncResult.error.message);
            
            // Execute callback
            if (report.syncOk.length + report.syncNotOk.length === report.count)
                onComplete(report);            
        }
    });              
}

/*
 * args:
 * - bindingId
 * - matrixData
 * - report
 * - onComplete
 */
function syncMatrixData(args) {    
    var bindingId = args.bindingId;
    var matrixData = args.matrixData;    
    var report = args.report;
    var onComplete = args.onComplete;  
    
    var bindingProperties = getBindingProperties(bindingId);
    var decimals = bindingProperties.decimals;
    
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
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            var binding = asyncResult.value;
            binding.setDataAsync(
                table,
                {
                    coercionType: "table",
                    startRow: bindingProperties.startingRow,
                    startColumn: bindingProperties.startingColumn
                },
                function(asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded)      
                        report.syncOk.push(bindingId);               
                    else {
                        report.syncNotOk.push(bindingId);
                        report.syncNotOkErrorCodes.push(asyncResult.error.code);
                        console.log(asyncResult.error.message);
                    }

                    // Execute callback
                    if (report.syncOk.length + report.syncNotOk.length === report.count)
                        onComplete(report);                
                }
            );
        }
        else {
            report.syncNotOk.push(bindingId);
            report.syncNotOkErrorCodes.push(asyncResult.error.code);
            console.log(asyncResult.error.message);
            
            // Execute callback
            if (report.syncOk.length + report.syncNotOk.length === report.count)
                onComplete(report);             
        }        
    });
} 

function syncBinding(bindingId) {
    
}
