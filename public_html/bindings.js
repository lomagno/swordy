/* global Office, Word, fabric, swire */

'use strict';

(function () {   
    var bindingTypeDropdown,
        dataNameTextEdit,
        dataNameLabel,
        successMsg,
        successMsgText,
        errorMsg,
        errorMsgText,
        dataNameErrorMsg,
        decimalsTextEdit,
        decimalsErrorMsg,
        bindButton,
        bindingsList,    
        sortMenu,
        sortMenuIsVisible = false,
        sortByMenuItems,
        orderMenuItems,
        commandBarElement,
        nBindings = 0,
        isDataNameValid = false,
        isDecimalsValid = true,
        stataNameRx = new RegExp(/^[a-zA-Z_][a-zA-Z_0-9]{0,31}$/);        
    
    Office.initialize = function (/* reason */) {
        $(document).ready(function () {               
            // Init Fabric components
            initFabricComponents();                        
            
            // Binding type dropdown
            bindingTypeDropdown = $('#bindingTypeDropdown');
            bindingTypeDropdown.change(onBindingTypeChanged);
            bindingTypeDropdown.find('option[value="scalar"]').prop('selected', true);            
            
            // Data name text edit
            dataNameTextEdit = $('#dataNameTextEdit');
            dataNameTextEdit.on('input', onDataNameTextEditChanged);
            dataNameErrorMsg = $('#dataNameErrorMsg');                        
            
            // Data name label
            dataNameLabel = $('#dataNameLabel');
            
            // Decimals text edit
            decimalsTextEdit = $('#decimalsTextEdit');
            decimalsTextEdit.on('input', onDecimalsTextEditChanged);
            decimalsErrorMsg = $('#decimalsErrorMsg');            

            // Success message
            successMsg = $('#success-msg');            
            $('#success-msg-close-link').click(function(event) {
                event.preventDefault();
                hideSuccessMsg();
            });
            successMsgText = $('#success-msg-text');            
            
            // Error message
            errorMsg = $('#error-msg');            
            $('#error-msg-close-link').click(function(event) {
                event.preventDefault();
                hideErrorMsg();
            });
            errorMsgText = $('#error-msg-text');                        
            
            // Bind button
            bindButton = $('#bindButton');
            bindButton.click(onBindButtonClicked);                       
            
            // Bindings list
            bindingsList = $('#bindingsList');
            
            // Sort button
            var sortButton = $('#sort-button');
            sortButton.unbind('click');
            sortButton.click(onSortButtonClicked);
            
            // "Sort" menu
            sortMenu = $('#sort-menu');
            
            // "Sort by" menu items
            sortByMenuItems = $('.sort-by-menu-item');
            sortByMenuItems.click(onSortByMenuItemClicked);
            
            // "Order" menu items
            orderMenuItems = $('.order-menu-item');
            orderMenuItems.click(onOrderMenuItemClicked);
            
            // Manage pivot button
            $('#manage-pivot-button').click(function() {
                setInterval(function() {commandBarElement._doResize();}, 500);
            });            

            Office.context.document.addHandlerAsync(
                Office.EventType.DocumentSelectionChanged,
                onDocumentSelectionChanged);
                   
            updateListBindings();               
        });
    };
    
    function onSortButtonClicked() {
        if (sortMenuIsVisible)
            hideSortMenu();
        else
            showSortMenu();
    }
    
    function showSortMenu() {
        sortMenu.show();     
        sortMenuIsVisible = true;        
    }
    
    function hideSortMenu() {
        sortMenu.hide();     
        sortMenuIsVisible = false;         
    }
    
    function onSortByMenuItemClicked() {
        var menuItem = $(this);
        sortByMenuItems.removeClass('is-selected');
        menuItem.addClass('is-selected');
        hideSortMenu();
        var order = orderMenuItems.filter('.is-selected').text();
        if (order === 'Ascending')
            order = 'asc';
        else if (order === 'Descending')
            order = 'desc';
        var sortBy = menuItem.text();
        if (sortBy === 'Name')
            sortBy = 'name';
        else if (sortBy === 'Type')
            sortBy = 'type';
        sortBindingsList(sortBy, order);
    }
    
    function onOrderMenuItemClicked() {
        var menuItem = $(this);
        orderMenuItems.removeClass('is-selected');
        menuItem.addClass('is-selected');
        hideSortMenu();
        var order = menuItem.text();
        if (order === 'Ascending')
            order = 'asc';
        else if (order === 'Descending')
            order = 'desc';
        var sortBy = sortByMenuItems.filter('.is-selected').text();
        if (sortBy === 'Name')
            sortBy = 'name';
        else if (sortBy === 'Type')
            sortBy = 'type';
        sortBindingsList(sortBy, order);        
    }  
    
    function sortBindingsList(property, order) {
        var bindingsListItems = $('#bindingsList li');
        bindingsListItems.sort(function(a, b) {
            var bindingId1 = $(a).data('binding');
            var bindingProperties1 = getBindingProperties(bindingId1);
            var property1 = bindingProperties1[property];
            var bindingId2 = $(b).data('binding');
            var bindingProperties2 = getBindingProperties(bindingId2); 
            var property2 = bindingProperties2[property];
            if (order === 'asc')
                return property1.localeCompare(property2);
            else if (order === 'desc')
                return property2.localeCompare(property1);
        });        
        bindingsListItems.detach().appendTo(bindingsList);        
    }
    
    function onBindButtonClicked() {
        hideAllMsg();
        
        Office.context.document.bindings.getAllAsync(function (asyncResult) {
            var innerIdArray = [];
            for (var i in asyncResult.value) {
                var bindingId = asyncResult.value[i].id;
                var innerId = bindingId.split('.')[1];
                innerIdArray.push(innerId);
            }
            innerIdArray.sort();
            var newBindingInnerId = 1;
            for (var i in innerIdArray) {
                if (newBindingInnerId !== +innerIdArray[i])
                    break;
                else
                    newBindingInnerId++;                    
            }
            var bindingType = getBindingType();
            console.log('bindingType = ' + bindingType);
            var dataName = dataNameTextEdit.val().trim();
            var newBindingId =
                'id.' + newBindingInnerId +
                '.type.' + bindingType +
                '.name.' + dataName +
                '.decimals.' + decimalsTextEdit.val().trim();
            var bindingTypeEnum;
            if (bindingType === 'scalar')
                bindingTypeEnum = Office.BindingType.Text;
            else if (bindingType === 'matrix')
                bindingTypeEnum = Office.BindingType.Table;                
            Office.context.document.bindings.addFromSelectionAsync(bindingTypeEnum, {id: newBindingId}, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed)
                    showErrorMsg('Can not create new binding: have you selected a portion of text or a table?');
                else {
                    dataNameTextEdit.val('');
                    isDataNameValid = false;
                    bindButton.prop('disabled', true);
                    showSuccessMsg('The binding for the ' + bindingType + ' "' + dataName + '" was created.');
                    var bindingListItem = createBindingListItem(newBindingId);
                    bindingsList.append(bindingListItem);
                    nBindings++;
                    sortBindingsList('name', 'asc');
                    asyncResult.value.addHandlerAsync(Office.EventType.BindingSelectionChanged, onBindingSelectionChanged);
                    asyncResult.value.addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
                }
            });
        });;
    }
    
    function initFabricComponents() {        
        // Init pivots
        var PivotElements = document.querySelectorAll(".ms-Pivot");
        for (var i = 0; i < PivotElements.length; i++) {
            new fabric['Pivot'](PivotElements[i]);
        }

        // Init command bars
        var CommandBarElements = document.querySelectorAll(".ms-CommandBar");
        for (var i = 0; i < CommandBarElements.length; i++) {
            commandBarElement = new fabric['CommandBar'](CommandBarElements[i]);            
        }

        // Init dropdowns
        var DropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');
        for (var i = 0; i < DropdownHTMLElements.length; ++i)
            new fabric['Dropdown'](DropdownHTMLElements[i]);

        // Init text fields
        var TextFieldElements = document.querySelectorAll(".ms-TextField");
        for (var i = 0; i < TextFieldElements.length; i++) {
            new fabric['TextField'](TextFieldElements[i]);
        }
        
        // Init lists
        var ListElements = document.querySelectorAll(".ms-List");
        for (var i = 0; i < ListElements.length; i++) {
            new fabric['List'](ListElements[i]);
        }        
    }
    
    function getBindingType() {
        return bindingTypeDropdown.find('option:checked').val();
    }
    
    function onBindingTypeChanged() {
        var bindingType = getBindingType();
        if (bindingType === 'scalar')
            dataNameLabel.text('Scalar name');
        else if (bindingType === 'matrix')
            dataNameLabel.text('Matrix name');
    }
    
    function onDataNameTextEditChanged() {
        hideAllMsg();
        
        // Validate scalar name
        var text = $(this).val().trim();
        if (text === '') {
            isDataNameValid = false;
            dataNameErrorMsg.text('A Stata scalar name is required');
            dataNameErrorMsg.show();            
        }
        else if (!stataNameRx.test(text)) {
            isDataNameValid = false;
            dataNameErrorMsg.text('Not valid Stata scalar name');
            dataNameErrorMsg.show();
        }
        else {
            isDataNameValid= true;
            dataNameErrorMsg.hide();
        }
        
        updateBindButtonStatus();
    } 
    
    function onDecimalsTextEditChanged() {
        hideAllMsg();
        
        // Validate decimals
        var text = $(this).val().trim();        
        if (text === '') {
            isDecimalsValid = false;
            decimalsErrorMsg.text('Decimals must be set');
            decimalsErrorMsg.show();
        }
        else if (!($.isNumeric(text) && isInteger(text))) {
            isDecimalsValid = false;
            decimalsErrorMsg.text('An integer number must be entered');
            decimalsErrorMsg.show();
        }
        else if (+text < 0 || +text > 20) {
            isDecimalsValid = false;
            decimalsErrorMsg.text('A integer value between 0 and 20 is required');
            decimalsErrorMsg.show();
        }
        else {
            isDecimalsValid = true;
            decimalsErrorMsg.hide();
        }
        
        updateBindButtonStatus();
    }       
    
    function updateBindButtonStatus() {
        if (isDataNameValid && isDecimalsValid)
            bindButton.prop('disabled', false);
        else
            bindButton.prop('disabled', true);
    }
    
    function showErrorMsg(msg) {
        errorMsgText.text(msg);
        errorMsg.show();    
    }
    
    function hideErrorMsg() {
        errorMsgText.text('');
        errorMsg.hide();
    }
    
    function showSuccessMsg(msg) {
        successMsgText.text(msg);
        successMsg.show();
    }
    
    function hideSuccessMsg() {
        successMsgText.text('');
        successMsg.hide();        
    }
    
    function hideAllMsg() {
        hideErrorMsg();
        hideSuccessMsg();
    }
    
    // This function is required for recent version of IE, because
    // the Number.isInteger function is not supported
    function isInteger(num){
        var numCopy = parseFloat(num);
        return !isNaN(numCopy) && numCopy == numCopy.toFixed();
    }
    
    function onDocumentSelectionChanged(eventArgs) {
        $('#bindingsList li').removeClass('is-unread');
        console.log(eventArgs);
    }
    
    function onBindingSelectionChanged(eventArgs) {
        console.log('onBindingSelectionChanged()');
        var bindingId = eventArgs.binding.id;
        
        // Unselect all items
        $('#bindingsList li').removeClass('is-unread');
        
        // Selected item
        var bindingSelectionItem = $('#bindingsList li[data-binding="' + bindingId + '"]');
        bindingSelectionItem.addClass('is-unread');
        
        // Scroll to item
        $('html, body').animate({scrollTop: bindingSelectionItem.offset().top}, 200);        
    }   
    
    function onBindingDataChanged(eventArgs) {
        var bindingId = eventArgs.binding.id;
        console.error('onBindingDataChanged(): bindingId = ' + bindingId + '; nBindings = ' + nBindings);
        var bindingListItem = $('#bindingsList li[data-binding="' + bindingId + '"]');
        Office.context.document.bindings.getAllAsync(
            {asyncContext: nBindings},
            function(asyncResult) {  
                if (asyncResult.value.length < asyncResult.asyncContext) {
                    bindingListItem.remove();
                    nBindings--;
            }
        });                
    }

    function updateListBindings() {        
        bindingsList.empty();
        Office.context.document.bindings.getAllAsync(function(asyncResult) {
            nBindings = asyncResult.value.length;
            for (var i in asyncResult.value) {
                var binding = asyncResult.value[i];                
                var bindingId = binding.id;
                var bindingListItem = createBindingListItem(bindingId);
                bindingsList.append(bindingListItem);
                asyncResult.value[i].addHandlerAsync(Office.EventType.BindingSelectionChanged, onBindingSelectionChanged);
                asyncResult.value[i].addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
            }
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
                nBindings--;
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