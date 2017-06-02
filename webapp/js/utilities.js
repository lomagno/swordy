/* global Office, swire */

function getBindingProperties(bindingId) {
    console.log('bindingId = ' + bindingId);
    
    var bindingProperties = {};
    
    // Split the binding ID string
    var bindingIdArray = bindingId.split('.');    
    
    // Read type
    var type;
    var i = 0;
    while (i < bindingIdArray.length-1) {
        var key = bindingIdArray[i];
        if (key === 'type') {
            type = bindingIdArray[i+1];
            break;
        }
        i = i + 2;
    }    
    
    // Read key and values
    var i = 0;
    while (i < bindingIdArray.length-1) {
        console.log('i = ' + i);
        var key = bindingIdArray[i];
        var value = bindingIdArray[i+1];
        
        console.log('key = ' + key);
        console.log('value = ' + value);
        
        if (
            key === 'id' ||
            key === 'startingRow' ||
            key === 'startingColumn'
        ) {
            console.log('t: numeric');
            value = +value;
        }
        else if (key === 'decimals') {
            console.log('t: decimals');
            if (type === 'scalar')
                value = +value;
            else {
                value = value.split('_');
                for (var j in value)
                    value[j] = +value[j];
            }
        }
        
        bindingProperties[key] = value;
        i = i + 2;
    } 
    
    return bindingProperties;
} 

function splitCommaSeparatedValues(string) {
    var vector = [];
    var splittedString = string.split(',');
    for (var i in splittedString)
        vector.push(parseInt(splittedString[i]));
    return vector;
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
                                    data = stataValueToText(scalarData, bindingObject.decimals, bindingObject.missings);
                                    
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
                                    var decimals =  bindingObject.decimals;
                                    var missingDecimalsCount = cols - decimals.length;
                                    for (var i=0; i<missingDecimalsCount; ++i)
                                        decimals.push(decimals[decimals.length-1]); 
                                    data.rows = [];
                                    var k=0;
                                    for (var i=0; i<rows; ++i) {
                                        var row = [];
                                        for (var j=0; j<cols; ++j) {
                                            // Text to be inserted in Word
                                            console.log(bindingObject);
                                            console.log('bindingObject.missings = ' + bindingObject.missings);
                                            var text = stataValueToText(values[k], bindingObject.decimals[j], bindingObject.missings);                                            
                                            console.log(text);
                                            
                                            // Add text to row
                                            row.push(text);
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

// TODO: Does this function must be used?
/*
 * args:
 * - binding
 * - data
 * - onComplete
 */
function writeScalarToBinding(args) {
    var
        binding = args.binding,
        data = args.data,
        onComplete = args.onComplete;
    
    // Binding ID
    var bindingId = binding.id;
    
    // Binding object
    var bindingObject = getBindingProperties(bindingId);
    
    // Format scalar text
    var text = data.toFixed(bindingObject.decimals);
    
    binding.setDataAsync(
        text,
        {
            coercionType: 'text',
            asyncContext: bindingObject
        },
        function(asyncResult) {
            onComplete(asyncResult.asyncContext);
        }
    );     
}

function getBindingListString(bindingObjectList) {
    var bindingListString = '';
    for (var i in bindingObjectList) {
        if (i > 0)
            bindingListString += ', ';
        bindingListString += bindingObjectList[i].name + ' (' + bindingObjectList[i].type + ')'; 
    }
    return bindingListString;
}

function getTextualReport(report) {
    // Parse the fields of the report
    var
        someBindingsFound = report.someBindingsFound,
        connectionSuccess = report.connectionSuccess,
        swireSuccess = report.swireSuccess,
        bindingsReport = report.bindingsReport;

    var status;
    var messages = [];

    if (!someBindingsFound) {
        status = 'error';
        messages.push('Cannot find any of the selected bindings in the Word document.');
    }
    else if (!connectionSuccess) {
        status = 'error';
        messages.push('Cannot connect to SWire.');
    }
    else if (!swireSuccess) {
        status = 'error';
        messages.push('SWire error. Please check that you are using SWire verson 0.2 or later.');        
    }
    else {
        var notFoundBindings = [];
        var bindingsWithNotFoundStataData = [];
        var bindingsWithSyncOk = [];
        var bindingsWithTableSizeError = [];
        var bindingsWithGenericError = [];
        for (var i in bindingsReport) {
            var bindingReport = bindingsReport[i];
            var bindingObject = bindingReport.bindingObject;
            if (!bindingReport.bindingFound)
                notFoundBindings.push(bindingObject);
            else if (!bindingReport.stataDataFound)
                bindingsWithNotFoundStataData.push(bindingObject);
            else if (!bindingReport.syncOk) {
                var setDataErrorCode = bindingReport.setDataErrorCode;
                switch (setDataErrorCode) {
                    case 2004:
                        bindingsWithTableSizeError.push(bindingObject);
                        break;
                    default:
                        bindingsWithGenericError.push(bindingObject);
                        break;
                }
            }
            else
                bindingsWithSyncOk.push(bindingObject);
        }

        if (bindingsWithSyncOk.length === bindingsReport.length) {
            status = 'ok';
            messages.push('Sync ok.');              
        }
        else {
            status = 'error';
            if (notFoundBindings.length > 0)
                messages.push(
                    'Cannot find the following binding(s): '
                    + getBindingListString(notFoundBindings)
                    + '.');
            if (bindingsWithNotFoundStataData.length > 0)
                messages.push(
                    'Cannot find the Stata data for the following binding(s): '
                    + getBindingListString(bindingsWithNotFoundStataData)
                    + '.');
            if (bindingsWithTableSizeError.length > 0)
                messages.push(
                    'Table size is too small for the following binding(s): '
                    + getBindingListString(bindingsWithTableSizeError)
                    + '.');
            if (bindingsWithGenericError.length > 0)
                messages.push(
                    'Cannot sync the following binding(s): '
                    + getBindingListString(bindingsWithGenericError)
                    + '.');
        }
    } 
    
    return {
        status: status,
        messages: messages
    };    
}

function stataValueToText(value, decimals, missingValues) {
    if (missingValues === 'special_ieee754')
        return value.toFixed(decimals);
    else {
        var missingType;
        switch (value) {
            case 8.9884656743115795E+307: // .
                missingType = '.';
                break;
            case 8.990660123939097E+307:  // a
                missingType = 'a';
                break;            
            case 8.9928545735666145E+307: // b
                missingType = 'b';
                break;            
            case 8.995049023194132E+307:  // c
                missingType = 'c';
                break;            
            case 8.9972434728216494E+307: // d
                missingType = 'd';
                break;            
            case 8.9994379224491669E+307: // e
                missingType = 'e';
                break;            
            case 9.0016323720766844E+307: // f
                missingType = 'f';
                break;            
            case 9.0038268217042019E+307: // g
                missingType = 'g';
                break;            
            case 9.0060212713317193E+307: // h
                missingType = 'h';
                break;            
            case 9.0082157209592368E+307: // i
                missingType = 'i';
                break;            
            case 9.0104101705867543E+307: // j
                missingType = 'j';
                break;            
            case 9.0126046202142718E+307: // k
                missingType = 'k';
                break;            
            case 9.0147990698417892E+307: // l
                missingType = 'l';
                break;            
            case 9.0169935194693067E+307: // m
                missingType = 'm';
                break;            
            case 9.0191879690968242E+307: // n
                missingType = 'n';
                break;            
            case 9.0213824187243417E+307: // o
                missingType = 'o';
                break;            
            case 9.0213824187243417E+307: // p
                missingType = 'p';
                break;            
            case 9.0257713179793766E+307: // q
                missingType = 'q';
                break;            
            case 9.0279657676068941E+307: // r
                missingType = 'r';
                break;            
            case 9.0301602172344116E+307: // s
                missingType = 's';
                break;            
            case 9.032354666861929E+307:  // t
                missingType = 't';
                break;            
            case 9.0345491164894465E+307: // u
                missingType = 'u';
                break;            
            case 9.036743566116964E+307:  // v
                missingType = 'v';
                break;            
            case 9.0389380157444815E+307: // w
                missingType = 'w';
                break;            
            case 9.041132465371999E+307:  // x
                missingType = 'x';
                break;            
            case 9.0433269149995164E+307: // y
                missingType = 'y';
                break;            
            case 9.0455213646270339E+307: // z
                missingType = 'z';
                break;                
            default:                      // no missing value
                missingType = null;
                break; 
        }
        
        if (missingType === null)
            return value.toFixed(decimals);
        else {
            switch (missingValues) {
                case "special_letters":
                    return missingType;
                case "special_pletters":
                    return '(' + missingType + ')';                
                case "string_-":
                    return "-";
                case "special_dot":
                    return ".";
                case "string_m":
                    return "m";
                case "string_NA":
                    return "NA";
                case "string_NaN":
                    return "NaN";
                default:
                    return value.toFixed(decimals);                    
            }
        }
    }    
}