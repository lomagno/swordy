/* global Office, swire */

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

/*
 * args:
 * - bindingIds
 * - onComplete
 */
function syncBindings(args) {    
    // Binding IDs
    var bindingIds = args.bindingIds;
    
    // Binding objects
    var bindingObjects = [];
    for (var i in bindingIds)
        bindingObjects.push(getBindingProperties(bindingIds[i]));
    
    // Init bindings report
    var bindingsReport = [];
    
    // Get bindings from the Word document
    getBindings(bindingIds, function(bindings) {
        // Requested Stata data
        var requestedStataData = [];
        for (var i in bindings)
            if (bindings[i] !== null) 
                requestedStataData.push({
                    name: bindingObjects[i].name,
                    type: bindingObjects[i].type
                });     
        
        // Check if no binding has been found
        if (requestedStataData.length === 0) {
            for (var i in bindingObjects)
                bindingsReport.push({
                    bindingObject: bindingObjects[i],
                    bindingFound: false,
                    dataFound: null,
                    syncOk: null
                });
            args.onComplete({
                connectionSuccess: null,
                bindingsReport: bindingsReport
            });
            return;
        }
        
        // SWire request
        var swireRequest = {                
            job: [
                {
                    method: '$getData',
                    args: {
                        data: requestedStataData
                    }                
                }
            ]
        };
       
        // SWire HTTP Ajax request
        $.ajax({
            url: 'https://localhost:50000',
            data: swire.encode(swireRequest),
            method: "POST", 
            success: function(swireEncodedResponse) {
                // Decode response
                var response = swire.decode(swireEncodedResponse);

                // Check errors
                if (response.status !== 'ok') {
                    // TODO: manage error
                    return;
                }                
                if (response.output[0].status !== 'ok') {
                    // TODO: manage error
                    return;                    
                }                

                // Stata data
                var retrievedStataData  = response.output[0].output.data;            
                
                // Count found Stata data
                var foundStataDataCount = 0;
                for (var i in retrievedStataData)
                    if (retrievedStataData[i] !== null)
                        foundStataDataCount++;
                
                // Check if no Stata data has been found
                if (foundStataDataCount === 0) {
                    for (var i in bindings)
                        bindingsReport.push({
                            bindingObject: bindingObjects[i],
                            bindingFound: bindings[i] === null ? false : true,
                            dataFound: false,
                            syncOk: null
                        });
                    args.onComplete({
                        connectionSuccess: true,
                        bindingsReport: bindingsReport
                    });
                    return;
                }
                
                var retrievedStataDataIndex = -1;
                var asyncDataSettingCompletedCount = 0;
                for (var i in bindings) {
                    var binding = bindings[i];
                    var bindingObject = bindingObjects[i];
                    if (binding !== null) {
                        retrievedStataDataIndex++;
                        var stataData = retrievedStataData[retrievedStataDataIndex];
                        if (stataData === null) {
                            bindingsReport.push({
                                bindingObject: bindingObject,
                                bindingFound: true,
                                dataFound: false,
                                syncOk: null
                            });                             
                            continue;
                        }
                        
                        var data, setDataAsyncOptions;
                        if (bindingObject.type === 'scalar') {
                            var scalarData = stataData;
                            
                            // Format scalar text
                            data = scalarData.toFixed(bindingObject.decimals);
                            
                            // Options
                            setDataAsyncOptions = {
                                coercionType: 'text',
                                asyncContext: bindingObject
                            };
                        }
                        else if (bindingObject.type === 'matrix') {
                            var matrixData = stataData;
                            
                            // Create table
                            data = new Office.TableData();
                            var rows = matrixData.rows;
                            var cols = matrixData.cols;
                            var values = matrixData.data;        
                            data.rows = [];
                            var k=0;
                            for (var i=0; i<rows; ++i) {
                                var row = [];
                                for (var j=0; j<cols; ++j) {
                                    row.push(values[k].toFixed(bindingObject.decimals));
                                    k++;
                                }
                                data.rows.push(row);
                            }   
                            
                            // Options
                            setDataAsyncOptions = {
                                coercionType: 'table',
                                startRow: bindingObject.startingRow,
                                startColumn: bindingObject.startingColumn,
                                asyncContext: bindingObject
                            };                           
                        }
                        
                        // Set data in the Word document
                        binding.setDataAsync(
                            data,
                            setDataAsyncOptions,
                            function(asyncResult) {
                                var bindingObject = asyncResult.asyncContext;
                                
                                if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
                                    bindingsReport.push({
                                        bindingObject: bindingObject,
                                        bindingFound: true,
                                        dataFound: true,
                                        syncOk: true
                                    });          
                                else
                                    bindingsReport.push({
                                        bindingObject: bindingObject,
                                        bindingFound: true,
                                        dataFound: true,
                                        syncOk: false,
                                        setDataErrorCode: asyncResult.error.code
                                    });                                        

                                asyncDataSettingCompletedCount++;

                                // Execute callback
                                if (asyncDataSettingCompletedCount === foundStataDataCount)
                                    args.onComplete({
                                        connectionSuccess: true,
                                        bindingsReport: bindingsReport
                                    });                
                            }
                        );                         
                    }
                    else
                        bindingsReport.push({
                            bindingObject: bindingObject,
                            bindingFound: false,
                            dataFound: null,
                            syncOk: null
                        });                         
                }
                        
            },
            error: function() {
                args.onComplete({
                    connectionSuccess: false
                });
            }
        });
    });            
}

function getBindings(bindingIds, onComplete) {
    var bindings = [];
    for (var i in bindingIds) {
        Office.context.document.bindings.getByIdAsync(
            bindingIds[i],
            {asyncContext: bindingIds.length},
            function (asyncResult) {    
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
                    bindings.push(asyncResult.value);
                else
                    bindings.push(null);

                if (bindings.length === asyncResult.asyncContext)
                    onComplete(bindings);
            }
        );          
    }
}
