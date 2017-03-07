/* global Office, Word, fabric, swire */

'use strict';

(function () {       
    var bindingsList;
    
    Office.initialize = function (/* reason */) {
        $(document).ready(function () {               
            // Init Fabric components
            initFabricComponents();
            
            bindingsList = $('#bindingsList');
                   
            Office.context.document.addHandlerAsync(
                Office.EventType.DocumentSelectionChanged,
                onDocumentSelectionChanged);  
                   
            listBindings();
        });
    };
    
    
    function initFabricComponents() {       
    }
    
    function initFabricListBindings() {
        var ListElements = document.querySelectorAll(".ms-List");
        for (var i = 0; i < ListElements.length; i++) {
            new fabric['List'](ListElements[i]);
        }        
    }
    
    function onDocumentSelectionChanged(eventArgs) {
        $('#bindingsList li').removeClass('is-unread');
        console.log(eventArgs);
    }
    
    function onBindingSelectionChanged(eventArgs) {
        var bindingId = eventArgs.binding.id;
        
        // Unselect all items
        $('#bindingsList li').removeClass('is-unread');
        
        // Selected item
        var bindingSelectionItem = $('#bindingsList li[data-binding="' + bindingId + '"]');
        bindingSelectionItem.addClass('is-unread');
        
        // Scroll to item
        $('html, body').animate({scrollTop: bindingSelectionItem.offset().top}, 200);        
    }    

    function listBindings() {        
        Office.context.document.bindings.getAllAsync(function (asyncResult) {
            for (var i in asyncResult.value) {
                var binding = asyncResult.value[i];                
                var bindingId = binding.id;
                var bindingListItem = createBindingListItem(bindingId);
                bindingsList.append(bindingListItem);
                asyncResult.value[i].addHandlerAsync(Office.EventType.BindingSelectionChanged, onBindingSelectionChanged);
            }
            initFabricListBindings();
        });
    }    
    
    function createBindingListItem(bindingId) {
        var bindingProperties = getBindingProperties(bindingId);
        
        var listItem = $('<li data-binding="' + bindingId  + '" class="ms-ListItem is-selectable" tabindex="0"></li>');
        listItem.click(onClickBindingListItem);
        
        // Primary text
        var primaryText = $('<span class="ms-ListItem-primaryText"></span>');
        primaryText.text(bindingProperties.name);
        listItem.append(primaryText);
        
        // Type
        var typeText = $('<span class="ms-ListItem-tertiaryText"></span>');
        typeText.text('Type: ' + bindingProperties.type);
        listItem.append(typeText);
        
        // Decimals
        var decimalsText = $('<span class="ms-ListItem-tertiaryText"></span>');
        decimalsText.text('Decimals: ' + bindingProperties.decimals);
        listItem.append(decimalsText);
        
        // Selection checkbox
        var checkbox = $('<div class="ms-ListItem-selectionTarget"></div>');
        listItem.append(checkbox);
        
        // Actions
        var actions = $('<div class="ms-ListItem-actions bindingActions"></div>');
        listItem.append(actions);
        
        // Delete button
        var deleteButton = $('<div class="ms-ListItem-action" title="Delete binding"><i class="ms-Icon ms-Icon--Delete"></i></div>');
        deleteButton.click(onDeleteBindingButtonClicked);
        actions.append(deleteButton);
        
        // Sync data button
        var syncDataButton = $('<div class="ms-ListItem-action" title="Sync data"><i class="ms-Icon ms-Icon--Sync"></i></div>');
        syncDataButton.click(onSyncDataButtonClicked);                
        actions.append(syncDataButton);        
        
        return listItem;
    }    
    
    function onDeleteBindingButtonClicked() {
        var listItem = $(this).closest('li.ms-ListItem');
        var bindingId = listItem.data('binding');
        Office.context.document.bindings.releaseByIdAsync(bindingId, function (asyncResult) { 
            console.log("Release binding status: " + asyncResult.status); 
            listItem.fadeOut(200, function() {
                listItem.remove();
            });
        });        
    }
    
    function onSyncDataButtonClicked() {
        var listItem = $(this).closest('li.ms-ListItem');
        var bindingId = listItem.data('binding');
        var bindingProperties = getBindingProperties(bindingId);
        
        var request;
        if (bindingProperties.type === 'scalar')
            request = {
                job: [
                    {
                        method: 'com.stata.sfi.Scalar.getValue',
                        args: [bindingProperties.name]
                    }
                ]
            };
        else if (bindingProperties.type === 'matrix')
            request = {
                job: [
                    {
                        method: '$getMatrix',
                        args: {
                            name: bindingProperties.name
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
                    console.log('SWire returned an error');
                    return;
                }                
                if (response.output[0].status !== 'ok') {
                    console.log('SWire returned an error');
                    return;                    
                }                
                if (response.output[0].output === null) {
                    console.log('Not existing data');
                    return;
                }
                
                // Stata data
                var data = response.output[0].output;
                
                // Update document
                if (bindingProperties.type === 'scalar')
                    syncScalarData(bindingId, data, bindingProperties.decimals);
                else if (bindingProperties.type === 'matrix')
                    syncMatrixData(bindingId, data, bindingProperties.decimals);
            },
            error: function (/* jqXHR, textStatus, errorThrown */) {
                console.log('Cannot communicate with Stata');
            }
        });        
    }
    
    function syncScalarData(bindingId, scalarValue, decimals) {
        console.log('syncBindingData() with binding id: ' + bindingId);
        
        var text = scalarValue.toFixed(decimals);
        Office.select('bindings#' + bindingId, function onError() {console.log('pippo errore');}).setDataAsync(text, function(asyncResult) {
            if (asyncResult.status === "failed")
                console.log('Error: ' + asyncResult.error.message);
            else
                console.log('Bound data: ' + asyncResult.value);               
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
    
    function onClickBindingListItem() {
        var listItem = $(this).closest('li.ms-ListItem');
        var bindingId = listItem.data('binding');
        console.log(bindingId);
    }

    function getBindingProperties(bindingId) {
        var numericProperties = ['id', 'decimals'];
        var bindingIdArray = bindingId.split('.');
        var bindingProperties = {};
        var i = 0;
        while (i<bindingIdArray.length-1) {
            var key = bindingIdArray[i];
            var value = bindingIdArray[i+1];
            if ($.inArray(key, numericProperties) !== -1)
                value = +value;
            bindingProperties[key] = value;
            i = i + 2;
        } 
        return bindingProperties;
    }    
})();