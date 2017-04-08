/* global Office, Word, fabric, swire */

'use strict';

(function () {   
    var bindingTypeDropdown,
        dataNameTextEdit,
        dataNameLabel,
        cSuccessMsg,
        cErrorMsg,        
        mSuccessMsg,
        mErrorMsg,
        dataNameErrorMsg,
        decimalsTextEdit,
        decimalsErrorMsg,
        bindButton,
        searchBox,
        searchBoxText = '',
        cancelSearchButton,
        bindingsList,
        fabricBindingsList,
        sortMenu,
        sortMenuIsVisible = false,
        sortByMenuItems,
        orderMenuItems,
        deleteSelectedBindingsButton,
        commandBarElement, // TODO: what is this?
        nBindings = 0,
        isDataNameValid = false,
        isDecimalsValid = true,
        toBeSynchronizedBindings = [],
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
            
            // "Create" success message
            cSuccessMsg = new MessageBar('create-success-msg');
            
            // "Create" error message
            cErrorMsg = new MessageBar('create-error-msg');            
            
            // Bind button
            bindButton = $('#bindButton');
            bindButton.click(onBindButtonClicked);                       
            
            // Bindings list
            bindingsList = $('#bindingsList');
            
            // Search box
            searchBox = $('#search-box');
            searchBox.bind('input', onSearchBoxChanged);
            searchBox.focus(onSearchBoxGainedFocus);
            
            // Select/deselect all bindings button
            $('#select-deselect-all-bindings-button').click(onSelectDeselectAllBindingsButtonClicked);
            
            // Cancel search
            cancelSearchButton = $('#cancel-search-button');
            cancelSearchButton.click(onCancelSearchButtonClicked);            
            
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
            
            // Delete button
            deleteSelectedBindingsButton = $('#deleteSelectedBindingsButton');
            deleteSelectedBindingsButton.click(onDeleteSelectedBindingsButton);
            
            // Sync button
            $('#sync-button').click(onSyncButtonClicked);
            
            // "Manage" success message
            mSuccessMsg = new MessageBar('manage-success-msg');
            
            // "Manage" error message
            mErrorMsg = new MessageBar('manage-error-msg');

            Office.context.document.addHandlerAsync(
                Office.EventType.DocumentSelectionChanged,
                onDocumentSelectionChanged);
                   
            updateBindingsList();               
        });
    };    
    
    function onSelectDeselectAllBindingsButtonClicked() {
        var items = bindingsList.find('.ms-ListItem');
        if (items.not('.is-selected').length > 0)
            items.addClass('is-selected');
        else
            items.removeClass('is-selected');
    }    
    
    function closeManageMesssages() {
        mSuccessMsg.close();
        mErrorMsg.close();
    }
    
    function onSyncButtonClicked() {
        closeManageMesssages();
        
        toBeSynchronizedBindings = [];
        var requestedData = [];
        Office.context.document.bindings.getAllAsync(function (asyncResult) { 
            for (var i in asyncResult.value) {
                var bindingId = asyncResult.value[i].id;
                toBeSynchronizedBindings.push(bindingId);
                var bindingProperties = getBindingProperties(bindingId);
                requestedData.push({
                    name: bindingProperties.name,
                    type: bindingProperties.type
                });
            }            
            
            var swireRequest = {                
                job: [
                    {
                        method: '$getData',
                        args: {
                            data: requestedData
                        }                
                    }
                ]
            };
            
            $.ajax({
                url: 'https://localhost:50000',
                data: swire.encode(swireRequest),
                method: "POST",
                success: function(swireEncodedResponse) {
                    // Decode response
                    var response = swire.decode(swireEncodedResponse);

                    // Check errors
                    if (response.status !== 'ok') {
                        toBeSynchronizedBindings = [];
                        console.error('SWire returned an error');
                        return;
                    }                
                    if (response.output[0].status !== 'ok') {
                        toBeSynchronizedBindings = [];
                        console.error('SWire returned an error');
                        return;                    
                    }                
                    
                    // Stata data
                    var retrievedData  = response.output[0].output.data;

                    // Update document
                    for (var i in toBeSynchronizedBindings) {
                        var bindingId = toBeSynchronizedBindings[i];
                        var bindingProperties = getBindingProperties(bindingId);
                        if (requestedData[i].type === 'scalar')
                            syncScalarData(bindingId, retrievedData[i], bindingProperties.decimals, onSyncCompleted);
                    }
                },
                error: function() {
                    toBeSynchronizedBindings = [];
                    mErrorMsg.showMessage('Cannot communicate with Stata');
                }
            });            
        });
    }
    
    function onSyncCompleted() {
        mSuccessMsg.showMessage('Sync completed');
    }
    
    function onSearchBoxChanged() {
        console.log('onSearchBoxChanged()');
        searchBoxText = $(this).val().trim();
        filterBindings(searchBoxText);
    }
    
    function onSearchBoxGainedFocus() {
        searchBox.val(searchBoxText);
    }
    
    function onCancelSearchButtonClicked() {
        searchBox.val('');
        searchBoxText = '';
        bindingsList.find('.ms-ListItem').each(function() {
            $(this).show();
        });
    }
    
    function filterBindings(text) {
        bindingsList.find('.ms-ListItem').each(function() {
            var listItem = $(this);
            var bindingId = listItem.data('binding');
            var bindingProperties = getBindingProperties(bindingId);
            var bindingName = bindingProperties.name;
            if (bindingName.indexOf(text) === -1)
                listItem.hide();
            else
                listItem.show();
        });        
    }
    
    function onDeleteSelectedBindingsButton() {
        var selectedItems = bindingsList.find('.ms-ListItem.is-selected');
        selectedItems.each(function() {
            var listItem = $(this);
            var bindingId = listItem.data('binding');
            Office.context.document.bindings.releaseByIdAsync(bindingId, function (asyncResult) { 
                console.log("Release binding status: " + asyncResult.status); // TODO: manage error 
                listItem.fadeOut(200, function() {
                    listItem.remove();
                    nBindings--;
                });
            }); 
        });
    }
    
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
        closeAllCreateMsg();
        
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
                    cErrorMsg.showMessage('Can not create new binding: have you selected a portion of text or a table?');
                else {
                    nBindings++;
                    dataNameTextEdit.val('');
                    isDataNameValid = false;
                    bindButton.prop('disabled', true);
                    updateBindingsList();
                    cSuccessMsg.showMessage('The binding for the ' + bindingType + ' "' + dataName + '" was created.');
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
        closeAllCreateMsg();
        
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
    
    function closeAllCreateMsg() {
        cSuccessMsg.close();
        cErrorMsg.close();
    }
    
    // This function is required for recent version of IE, because
    // the Number.isInteger function is not supported
    function isInteger(num){
        var numCopy = parseFloat(num);
        return !isNaN(numCopy) && numCopy == numCopy.toFixed();
    }
    
    function onDocumentSelectionChanged(eventArgs) {
        $('#bindingsList li').removeClass('is-unread');
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
    
    function onBindingDataChanged(eventArgs) {
        var bindingId = eventArgs.binding.id;
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

    function updateBindingsList() {
        // Empty bindings list
        bindingsList.empty();
        
        // Create bindings list
        Office.context.document.bindings.getAllAsync(function(asyncResult) {            
            for (var i in asyncResult.value) {
                var binding = asyncResult.value[i];                
                var bindingId = binding.id;
                var bindingListItem = createBindingListItem(bindingId);
                bindingsList.append(bindingListItem);
                asyncResult.value[i].addHandlerAsync(Office.EventType.BindingSelectionChanged, onBindingSelectionChanged);
                asyncResult.value[i].addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
            }
            var bindingsListElement = document.getElementById('bindingsList');
            fabricBindingsList = new fabric['List'](bindingsListElement);
        });
    }
    
    function createBindingListItem(bindingId) {
        var bindingProperties = getBindingProperties(bindingId);
        
        var listItem = $('<li data-binding="' + bindingId  + '" class="ms-ListItem is-selectable" tabindex="0"></li>');
        
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
        //checkbox.click(onBindingChecked); // TODO: what to do with this?
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
        syncDataButton.click(onIndividualSyncDataButtonClicked);                
        actions.append(syncDataButton);                
        
        return listItem;
    }        
    
    /*
    function onBindingChecked() {  
        var isSelected = $(this).parent('.ms-ListItem').hasClass('is-selected');
        var deltaChecked = isSelected ? -1 : 1;
        if (bindingsList.find('.ms-ListItem.is-selected').length + deltaChecked > 0)
            deleteSelectedBindingsButton.prop('disabled', false);
        else
            deleteSelectedBindingsButton.prop('disabled', true);
    }
    
    function updateDeleteSelectedBindingsButtonStatus() {
        if (bindingsList.find('.ms-ListItem.is-selected').length > 0)
            deleteSelectedBindingsButton.prop('disabled', false);
        else
            deleteSelectedBindingsButton.prop('disabled', true);        
    }
    */
    
    function onDeleteBindingButtonClicked() {
        var listItem = $(this).closest('li.ms-ListItem');
        var bindingId = listItem.data('binding');
        Office.context.document.bindings.releaseByIdAsync(bindingId, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
                listItem.fadeOut(200, function() {
                    listItem.remove();
                    nBindings--;
                });
        });        
    }
    
    function onIndividualSyncDataButtonClicked() {
        var listItem = $(this).closest('li.ms-ListItem');
        var bindingId = listItem.data('binding');
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
                toBeSynchronizedBindings = [];
                toBeSynchronizedBindings.push(bindingId);
                if (bindingProperties.type === 'scalar')
                    syncScalarData(bindingId, data, bindingProperties.decimals, onIndividualSyncDataCompleted);
                else if (bindingProperties.type === 'matrix')
                    syncMatrixData(bindingId, data, bindingProperties.decimals);
            },
            error: function (/* jqXHR, textStatus, errorThrown */) {
                mErrorMsg.showMessage('Cannot communicate with Stata');
            }
        });        
    }
    
    function onIndividualSyncDataCompleted() {
        // TODO: inform the user
        console.log('ok');
    }
    
    function syncScalarData(bindingId, scalarValue, decimals, onComplete) {        
        var text = scalarValue.toFixed(decimals);
        Office.select('bindings#' + bindingId, function() {console.log('pippo errore');}).setDataAsync(text, {asyncContext: bindingId}, function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
            {
                // Remove binding from the toBeSynchronizedBindings array
                toBeSynchronizedBindings.splice(toBeSynchronizedBindings.indexOf(asyncResult.asyncContext), 1);
                
                // Execute callback if the toBeSynchronizedBindings array is void
                if (toBeSynchronizedBindings.length === 0)
                    onComplete();                
            }
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