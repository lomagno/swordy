/* global fabric, Office, swire */

function BindingsList(elementId) {
    this.update = function() {        
        Office.context.document.bindings.getAllAsync(function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                // Number of bindings
                m_nBindings = asyncResult.value.length;
                
                // Empty bindings list
                m_container.empty();                
                
                // Create items
                for (var i in asyncResult.value) {
                    var binding = asyncResult.value[i];                
                    var bindingId = binding.id;
                    var bindingListItem = createItem(bindingId);
                    m_container.append(bindingListItem);
                    binding.addHandlerAsync(Office.EventType.BindingSelectionChanged, onBindingSelectionChanged);
                    binding.addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
                }
                m_self.sort(m_sortBy, m_order);
                m_self.filter(m_filterString);
                initFabric();      
            }
            else
                console.error('Cannot update the bindings list');
        });
    };    
    
    this.sort = function(sortBy, order) {
        var bindingsListItems = m_container.find('li');
        bindingsListItems.sort(function(a, b) {
            var bindingId1 = $(a).data('binding');
            var bindingProperties1 = getBindingProperties(bindingId1);
            var property1 = bindingProperties1[sortBy];
            var bindingId2 = $(b).data('binding');
            var bindingProperties2 = getBindingProperties(bindingId2); 
            var property2 = bindingProperties2[sortBy];
            if (order === 'asc')
                return property1.localeCompare(property2);
            else if (order === 'desc')
                return property2.localeCompare(property1);
        });        
        bindingsListItems.detach().appendTo(m_container);        
    };       
    
    this.filter = function(filterString) {
        m_filterString = filterString;
        var items = m_container.find('li');
        if (filterString !== '') {
            items.each(function() {
                var item = $(this);
                var bindingId = item.data('binding');
                var bindingProperties = getBindingProperties(bindingId);
                var bindingName = bindingProperties.name;
                if (bindingName.indexOf(filterString) === -1)
                    item.hide();
                else
                    item.show();
            });
        }
        else
            items.each(function() {
                $(this).show();
            });            
    };
    
    this.selectDeselectAll = function() {
        var items = m_container.find('li');
        if (items.not('.is-selected').length > 0)
            items.addClass('is-selected');
        else
            items.removeClass('is-selected');
    };
    
    this.deleteSelected = function() {
        var selectedItems = m_container.find('li.is-selected');
        selectedItems.each(function() {
            var listItem = $(this);
            var bindingId = listItem.data('binding');
            Office.context.document.bindings.releaseByIdAsync(bindingId, function (asyncResult) { 
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded)                
                    removeItem(listItem);
                else
                    console.error('Error deleting binding');
            }); 
        });
    };
    
    function createItem(bindingId) {
        var bindingProperties = getBindingProperties(bindingId);
        
        // List item
        var listItem = $('<li data-binding="' + bindingId  + '" class="ms-ListItem is-selectable" tabindex="0"></li>'); // TODO: what tabindex stands for?
        
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
        
        var msg = $('<div class="ms-MessageBar ms-MessageBar--success ms-u-slideDownIn20"><div class="ms-MessageBar-content"><div class="ms-MessageBar-icon"><i class="ms-Icon ms-Icon--Completed"></i></div><div class="ms-MessageBar-text"><span class="mb-text">Sync was successful asdjklas jdlkasjd lkasjd lsjdlkas jdklasjdl kasjdkl asjdlk jasldk jasldjsdjlsjdlkas jdlkas jdl</span><br /><a class="mb-close-link ms-Link" href="#">Close this</a></div></div></div>');
        msg.show();
        listItem.append(msg);
        
        // Selection checkbox
        var checkbox = $('<div class="ms-ListItem-selectionTarget"></div>');
        //checkbox.click(onBindingChecked); // TODO: what to do with this?
        listItem.append(checkbox);
        
        // Actions
        var actionsContainer = $('<div class="ms-ListItem-actions bindingActions"></div>');
        listItem.append(actionsContainer);
        
        // Delete button
        var deleteButton = $('<div class="ms-ListItem-action" title="Delete binding"><i class="ms-Icon ms-Icon--Delete"></i></div>');
        deleteButton.click(function() {onDeleteBindingButtonClicked(listItem);});
        actionsContainer.append(deleteButton);
        
        // Sync data button
        var syncDataButton = $('<div class="ms-ListItem-action" title="Sync data"><i class="ms-Icon ms-Icon--Sync"></i></div>');
        syncDataButton.click(function() {onSyncDataButtonClicked(listItem);});   
        actionsContainer.append(syncDataButton);                
        
        return listItem;
    } 
    
    function onDeleteBindingButtonClicked(listItem) {
        var bindingId = listItem.data('binding');
        Office.context.document.bindings.releaseByIdAsync(bindingId, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
                removeItem(listItem);
            else
                console.error('Cannot update the bindings list');
        });        
    }      
    
    function onBindingSelectionChanged(eventArgs) {
        // Binding ID
        var bindingId = eventArgs.binding.id;
        
        // Unselect all items
        m_container.find('li').removeClass('is-unread');
        
        // Selected item
        var bindingSelectionItem = m_container.find('li[data-binding="' + bindingId + '"]');
        bindingSelectionItem.addClass('is-unread');
        
        // Scroll to item
        $('html, body').animate({scrollTop: bindingSelectionItem.offset().top}, 200);        
    } 
    
    function onBindingDataChanged(eventArgs) {
        var bindingId = eventArgs.binding.id;
        var bindingListItem = m_container.find('li[data-binding="' + bindingId + '"]');
        Office.context.document.bindings.getAllAsync(
            {asyncContext: m_nBindings},
            function(asyncResult) { 
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    if (asyncResult.value.length < asyncResult.asyncContext)
                        removeItem(bindingListItem);
                }
                else
                    console.error('Cannot update the bindings list');
            }
        );                
    }    
    
    function onSyncDataButtonClicked(item) {
        var bindingId = item.data('binding');
        var bindingProperties = getBindingProperties(bindingId);
        
        // SWire request
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
            success: function (swireEncodedResponse) {
                // Decode response
                var response = swire.decode(swireEncodedResponse);
                
                // Check errors
                if (response.status !== 'ok') {
                    console.error('SWire returned an error');
                    return;
                }                
                if (response.output[0].status !== 'ok') {
                    console.error('SWire returned an error');
                    return;                    
                }                
                if (response.output[0].output === null) {
                    console.error('Not existing data');
                    return;
                }
                
                // Stata data
                var data = response.output[0].output;
                
                // Update document
                if (bindingProperties.type === 'scalar')
                    syncScalarData(bindingId, data, bindingProperties.decimals, null);
                else if (bindingProperties.type === 'matrix')
                    syncMatrixData(bindingId, data, bindingProperties.decimals);
            },
            error: function (/* jqXHR, textStatus, errorThrown */) {
                console.error('Cannot communicate with Stata'); // TODO: manage this
                //mErrorMsg.showMessage('Cannot communicate with Stata'); // TODO: manage this
            }
        });        
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
    
    function removeItem(listItem) {
        listItem.fadeOut(200, function() {
            listItem.remove();
            m_nBindings--;
        });        
    }
    
    function initFabric() {
        new fabric['List'](m_container[0]);
    }
    
    function onDocumentSelectionChanged(/* eventArgs */) { // TODO: is eventArgs useful?
        m_container.find('li').removeClass('is-unread');
    }       
    
    // Variables
    var
        m_self = this,
        m_container = $('#' + elementId),
        m_nBindings = 0,
        m_sortBy = 'name',
        m_order = 'asc',
        m_filterString = '';
    
    // Document selection changed event
    Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onDocumentSelectionChanged);     
    
    // Update list
    this.update();
}