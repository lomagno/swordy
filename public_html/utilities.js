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
    
    // Get bindings from the Word document
    getBindings(bindingIds, function(bindings) {
        // Init bindings report
        var bindingsReport = [];        
        
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
                    stataDataFound: null,
                    syncOk: null
                });
            args.onComplete({
                someBindingsFound: false,
                connectionSuccess: null,
                swireSuccess: null,
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
                    for (var i in bindings)
                        bindingsReport.push({
                            bindingObject: bindingObjects[i],
                            bindingFound: bindings[i] === null ? false : true,
                            stataDataFound: null,
                            syncOk: null
                        });
                    args.onComplete({
                        someBindingsFound: true,
                        connectionSuccess: true,
                        swireSuccess: false,
                        bindingsReport: bindingsReport
                    });
                    return;
                }                
                if (response.output[0].status !== 'ok') {
                    for (var i in bindings)
                        bindingsReport.push({
                            bindingObject: bindingObjects[i],
                            bindingFound: bindings[i] === null ? false : true,
                            stataDataFound: null,
                            syncOk: null
                        });
                    args.onComplete({
                        someBindingsFound: true,
                        connectionSuccess: true,
                        swireSuccess: false,
                        bindingsReport: bindingsReport
                    });
                    return;                   
                }                

                // Stata data
                var retrievedStataData  = response.output[0].output.data;
                
                var bindingIndex = -1;
                var retrievedStataDataIndex = -1;
                
                function manageNextBinding() {
                    bindingIndex++;
                    if (bindingIndex < bindings.length) {
                        var binding = bindings[bindingIndex];
                        var bindingObject = bindingObjects[bindingIndex];
                        if (binding !== null) {
                            retrievedStataDataIndex++;
                            var stataData = retrievedStataData[retrievedStataDataIndex];
                            if (stataData !== null) {
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
                                binding.setDataAsync(
                                    data,
                                    setDataAsyncOptions,
                                    function(asyncResult) {
                                        var bindingObject = asyncResult.asyncContext;

                                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
                                            bindingsReport.push({
                                                bindingObject: bindingObject,
                                                bindingFound: true,
                                                stataDataFound: true,
                                                syncOk: true
                                            });          
                                        else
                                            bindingsReport.push({
                                                bindingObject: bindingObject,
                                                bindingFound: true,
                                                stataDataFound: true,
                                                syncOk: false,
                                                setDataErrorCode: asyncResult.error.code
                                            });                                        
                                        manageNextBinding();
                                    }
                                );                                 
                            }
                            else {
                                bindingsReport.push({
                                    bindingObject: bindingObject,
                                    bindingFound: true,
                                    stataDataFound: false,
                                    syncOk: null
                                });                             
                                manageNextBinding();
                            }                            
                        }
                        else {
                            bindingsReport.push({
                                bindingObject: bindingObject,
                                bindingFound: false,
                                stataDataFound: null,
                                syncOk: null
                            });
                            manageNextBinding();
                        }
                    }
                    else
                        args.onComplete({
                            someBindingsFound: true,
                            connectionSuccess: true,
                            swireSuccess: true,
                            bindingsReport: bindingsReport
                        });                          
                }

                manageNextBinding();                        
            },
            error: function() {
                for (var i in bindings)
                    bindingsReport.push({
                        bindingObject: bindingObjects[i],
                        bindingFound: bindings[i] === null ? false : true,
                        stataDataFound: null,
                        syncOk: null
                    });                
                args.onComplete({
                    someBindingsFound: true,
                    connectionSuccess: false,
                    swireSuccess: null,
                    bindingsReport: bindingsReport
                });
            }
        });
    });            
}

function getBinding(bindingId, onComplete) {
    Office.context.document.bindings.getByIdAsync(
        bindingId,
        function (asyncResult) {    
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
                onComplete(asyncResult.value);
            else
                onComplete(null);
        }
    );        
}

function getBindings(bindingIds, onComplete) {    
    function onGetBindingComplete(binding) {
        m_bindings.push(binding);
        m_index++;
        if (m_index < bindingIds.length)
            getBinding(bindingIds[m_index], onGetBindingComplete);
        else
            onComplete(m_bindings);
    }
    
    var m_bindings = [];
    var m_index = 0;
    if (bindingIds.length > 0)
        getBinding(bindingIds[m_index], onGetBindingComplete);
    else
        onComplete(m_bindings);
}
