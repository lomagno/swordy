/* global Office, Word, fabric, swire */

'use strict';

(function () {
    var m_bindingsList,
        bindingTypeDropdown,
        m_scalarNameTextField,
        m_matrixNameTextField,
        m_startingRowTextField,
        m_startingColumnTextField,
        m_decimalsTextField,
        m_decimalsForColumnsTextField,
        m_missingValuesDropdown,
        cSuccessMsg,
        cErrorMsg,
        mSuccessMsg,
        mErrorMsg,
        m_bindButton,
        searchBox,
        searchBoxText = '',
        cancelSearchButton,
        sortMenu,
        sortMenuIsVisible = false,
        sortByMenuItems,
        orderMenuItems,
        m_checkUncheckAllBindingsButton,
        m_deleteSelectedBindingsButton,
        m_syncSelectedBindingsButton,
        commandBarElement, // TODO: what is this?
        stataNameRx = new RegExp(/^[a-zA-Z_][a-zA-Z_0-9]{0,31}$/),
        m_decimalsListRx = new RegExp(/^\s*([0-9]|1[0-9]|20)\s*(,\s*([0-9]|1[0-9]|20))*\s*$/);

    Office.initialize = function () {
        $(document).ready(function () {            
            // Init Fabric components
            initFabricComponents();

            // Binding type dropdown
            bindingTypeDropdown = $('#bindingTypeDropdown');
            var fabricBindingTypeDropdown = new fabric['Dropdown'](bindingTypeDropdown[0]);
            $(fabricBindingTypeDropdown._dropdownItems[1].newItem).click(); // Select "scalar"
            bindingTypeDropdown.find('.ms-Dropdown-select').change(onBindingTypeChanged);
            new FieldWithHelp('bindingTypeDropdown');

            // Bind button
            m_bindButton = $('#bindButton');
            new fabric['Button'](m_bindButton[0], onBindButtonClicked);

            // Validators
            var integerNumberValidator = function (text) {
                if (!($.isNumeric(text) && isInteger(text)))
                    return {
                        isValid: false,
                        errorMessage: 'An integer number must be entered'
                    };
                else
                    return {isValid: true};
            };
            var startingRowColumnRangeValidator = function (text) {
                if (+text < 1 || +text > 99999)
                    return {
                        isValid: false,
                        errorMessage: 'A integer value between 1 and 99999 is required'
                    };
                else
                    return {isValid: true};
            };

            // Scalar name text field
            m_scalarNameTextField = new TextField({
                elementId: 'scalarNameTextField',
                validators: [
                    function (text) {
                        if (text === '')
                            return {
                                isValid: false,
                                errorMessage: 'A Stata data name is required'
                            };
                        else
                            return {isValid: true};
                    },
                    function (text) {
                        if (!stataNameRx.test(text))
                            return {
                                isValid: false,
                                errorMessage: 'Not valid Stata data name'
                            };
                        else
                            return {isValid: true};
                    }
                ],
                onErrorStatusChanged: updateBindButtonStatus
            });
            m_scalarNameTextField.setValue('', false);
            new FieldWithHelp('scalarNameTextField');
            
            // Matrix name text field
            m_matrixNameTextField = new TextField({
                elementId: 'matrixNameTextField',
                validators: [
                    function (text) {
                        if (text === '')
                            return {
                                isValid: false,
                                errorMessage: 'A Stata data name is required'
                            };
                        else
                            return {isValid: true};
                    },
                    function (text) {
                        if (!stataNameRx.test(text))
                            return {
                                isValid: false,
                                errorMessage: 'Not valid Stata data name'
                            };
                        else
                            return {isValid: true};
                    }
                ],
                onErrorStatusChanged: updateBindButtonStatus
            });
            m_matrixNameTextField.setValue('', false);
            new FieldWithHelp('matrixNameTextField');
            m_matrixNameTextField.hide();

            // Starting row text field
            m_startingRowTextField = new TextField({
                elementId: 'startingRowTextField',
                validators: [
                    function (text) {
                        if (text === '')
                            return {
                                isValid: false,
                                errorMessage: 'A starting row must be set'
                            };
                        else
                            return {isValid: true};
                    },
                    integerNumberValidator,
                    startingRowColumnRangeValidator
                ],
                onErrorStatusChanged: updateBindButtonStatus
            });
            m_startingRowTextField.setValue('1');            
            new FieldWithHelp('startingRowTextField');
            m_startingRowTextField.hide();

            // Starting column text field
            m_startingColumnTextField = new TextField({
                elementId: 'startingColumnTextField',
                validators: [
                    function (text) {
                        if (text === '')
                            return {
                                isValid: false,
                                errorMessage: 'A starting row must be set'
                            };
                        else
                            return {isValid: true};
                    },
                    integerNumberValidator,
                    startingRowColumnRangeValidator
                ],
                onErrorStatusChanged: updateBindButtonStatus
            });
            m_startingColumnTextField.setValue('1');            
            new FieldWithHelp('startingColumnTextField');
            m_startingColumnTextField.hide();

            // Decimals text field
            m_decimalsTextField = new TextField({
                elementId: 'decimalsTextField',
                validators: [
                    function (text) {
                        if (text === '')
                            return {
                                isValid: false,
                                errorMessage: 'Decimals must be set'
                            };
                        else
                            return {isValid: true};
                    },
                    integerNumberValidator,
                    function (text) {
                        if (+text < 0 || +text > 20) {
                            return {
                                isValid: false,
                                errorMessage: 'An integer value between 0 and 20 is required'
                            };
                        } else
                            return {isValid: true};
                    }
                ],
                onErrorStatusChanged: updateBindButtonStatus
            });
            m_decimalsTextField.setValue('3');
            new FieldWithHelp('decimalsTextField');
            
            // Decimals for columns text field            
            m_decimalsForColumnsTextField = new TextField({
                elementId: 'decimalsForColumnsTextField',
                validators: [
                    function (text) {
                        if (text === '')
                            return {
                                isValid: false,
                                errorMessage: 'Decimals must be set'
                            };
                        else
                            return {isValid: true};
                    },
                    function (text) {
                        if (!m_decimalsListRx.test(text))
                            return {
                                isValid: false,
                                errorMessage:
                                    'Not valid decimals list: must be a list of integers between 0 and 20 separated by commas.'
                                    + ' Example: 4, 1, 5.'
                            };
                        else
                            return {isValid: true};
                    }
                ],
                onErrorStatusChanged: updateBindButtonStatus
            });
            m_decimalsForColumnsTextField.setValue('3');
            m_decimalsForColumnsTextField.hide();
            new FieldWithHelp('decimalsForColumnsTextField');
            
            // Missing values dropdown
            m_missingValuesDropdown = $('#missingValuesDropdown');
            var fabricMissingValuesDropdown = new fabric['Dropdown'](m_missingValuesDropdown[0]);
            $(fabricMissingValuesDropdown._dropdownItems[1].newItem).click(); // Select "special_letters"
            new FieldWithHelp('missingValuesDropdown');            

            // "Create" success message
            cSuccessMsg = new MessageBar('create-success-msg');

            // "Create" error message
            cErrorMsg = new MessageBar('create-error-msg');

            // Search box
            searchBox = $('#search-box');
            searchBox.bind('input', onSearchBoxChanged);
            searchBox.focus(onSearchBoxGainedFocus);

            // Check/uncheck all bindings button
            m_checkUncheckAllBindingsButton = new CommandBarButton({
                elementId: 'check-uncheck-all-bindings-button',
                onClick: onCheckUncheckAllBindingsButtonClicked
            });
            m_checkUncheckAllBindingsButton.setEnabled(false);

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
            $('#manage-pivot-button').click(function () {
                setInterval(function () {
                    commandBarElement._doResize();
                }, 500);
            });

            // Delete selected bindings button
            m_deleteSelectedBindingsButton = new CommandBarButton({
                elementId: 'deleteSelectedBindingsButton',
                onClick: onDeleteSelectedBindingsButtonClicked
            });
            m_deleteSelectedBindingsButton.setEnabled(false);

            // Sync selected bindings button
            m_syncSelectedBindingsButton = new CommandBarButton({
                elementId: 'sync-selected-bindings-button',
                onClick: onSyncSelectedBindingsButtonClicked
            });
            m_syncSelectedBindingsButton.setEnabled(false);

            $('#refresh-bindings-list-button').click(onRefreshBindingsListButtonClicked);

            // "Manage" success message
            mSuccessMsg = new MessageBar('manage-success-msg');

            // "Manage" error message
            mErrorMsg = new MessageBar('manage-error-msg');

            // Bindings list
            m_bindingsList = new BindingsList({
                elementId: 'bindingsList',
                onListStatusChanged: onListStatusChanged
            });
        });
    };

    function onRefreshBindingsListButtonClicked() {
        m_bindingsList.update(function (status) {
            if (status === 'ok')
                mSuccessMsg.showMessage('The bindings list was correctly refreshed.');
            else
                mErrorMsg.showMessage('There was an error while refreshing the bindings list.');
        });
    }

    function onListStatusChanged(status) {
        m_deleteSelectedBindingsButton.setEnabled(status.selection !== 'nothing');
        m_syncSelectedBindingsButton.setEnabled(status.selection !== 'nothing');
        m_checkUncheckAllBindingsButton.setEnabled(status.populated);
    }

    function onDeleteSelectedBindingsButtonClicked() {
        m_bindingsList.deleteCheckedItems(function (report) {
            var erroneousBindingReleases = report.erroneousBindingReleases;
            if (erroneousBindingReleases.length === 0)
                mSuccessMsg.showMessage('Delete succeeded.');
            else {
                var errBindings = '';
                for (var i in erroneousBindingReleases) {
                    var bindingProperties = getBindingProperties(erroneousBindingReleases[i]);
                    if (i > 0)
                        errBindings += ', ';
                    errBindings += bindingProperties.name + ' (' + bindingProperties.type + ')';
                }
                mErrorMsg.showMessage('Cannot delete the following bindings: ' + errBindings
                        + '. Maybe these bindings are no longer in the Word document'
                        + '. Please try to refresh the bindings list.');
            }
        });
    }

    function onCheckUncheckAllBindingsButtonClicked() {
        m_bindingsList.checkUncheckAll();
    }

    function closeManageMesssages() {
        mSuccessMsg.close();
        mErrorMsg.close();
    }

    function onSyncSelectedBindingsButtonClicked() {
        closeManageMesssages(); // TODO: it this needed?        

        // Selected binding IDs
        var selectedBindingIds = [];
        var checkedItems = m_bindingsList.getCheckedItems();
        for (var i in checkedItems) {
            var item = checkedItems[i];
            selectedBindingIds.push(item.getBindingId());
        }

        syncBindings({
            bindingIds: selectedBindingIds,
            onComplete: onSyncCompleted
        });
    }

    function onSyncCompleted(report) {
        var textualReport = getTextualReport(report);
        var status = textualReport.status;
        var messages = textualReport.messages;
        if (status === 'ok')
            mSuccessMsg.showMessage(messages[0]);
        else {
            if (messages.length === 1)
                mErrorMsg.showMessage(messages[0]);
            else {
                mErrorMsg.reset();
                mErrorMsg.appendList();
                for (var i in messages)
                    mErrorMsg.appendListItem(messages[i]);
                mErrorMsg.show();
            }
        }             
    }    

    function onSearchBoxChanged() {
        searchBoxText = $(this).val().trim();
        m_bindingsList.filter(searchBoxText);
    }

    function onSearchBoxGainedFocus() {
        searchBox.val(searchBoxText);
    }

    function onCancelSearchButtonClicked() {
        searchBox.val('');
        searchBoxText = '';
        m_bindingsList.filter('');
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
        var orderFromMenu = orderMenuItems.filter('.is-selected').text();
        var order;
        if (orderFromMenu === 'Ascending')
            order = 'asc';
        else if (orderFromMenu === 'Descending')
            order = 'desc';
        var sortByFromMenu = menuItem.text();
        var sortBy;
        if (sortByFromMenu === 'Name')
            sortBy = 'name';
        else if (sortByFromMenu === 'Type')
            sortBy = 'type';
        m_bindingsList.sort(sortBy, order);
    }

    function onOrderMenuItemClicked() {
        var menuItem = $(this);
        orderMenuItems.removeClass('is-selected');
        menuItem.addClass('is-selected');
        hideSortMenu();
        var orderFromMenu = menuItem.text();
        var order;
        if (orderFromMenu === 'Ascending')
            order = 'asc';
        else if (orderFromMenu === 'Descending')
            order = 'desc';
        var sortByFromMenu = sortByMenuItems.filter('.is-selected').text();
        var sortBy;
        if (sortByFromMenu === 'Name')
            sortBy = 'name';
        else if (sortByFromMenu === 'Type')
            sortBy = 'type';
        m_bindingsList.sort(sortBy, order);
    }

    function onBindButtonClicked() {
        closeAllCreateMsg();

        Office.context.document.bindings.getAllAsync(function (asyncResult) {
            // TODO: change this strategy for binding inner ID generation?
            // New binding inner ID
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
            var dataName;          
            var newBindingId;
            var bindingTypeEnum;
            if (bindingType === 'scalar') {
                dataName = m_scalarNameTextField.getValue().trim();
                newBindingId =
                        'id.' + newBindingInnerId +
                        '.type.' + bindingType +
                        '.name.' + dataName +
                        '.decimals.' + m_decimalsTextField.getValue().trim() +
                        '.missings.' + m_missingValuesDropdown.find('option:checked').val();
                bindingTypeEnum = Office.BindingType.Text;
            } else if (bindingType === 'matrix') {
                dataName = m_matrixNameTextField.getValue().trim();
                var decimalsList = splitCommaSeparatedValues(m_decimalsForColumnsTextField.getValue());
                var encodedDecimals = decimalsList.join('_');
                newBindingId =
                        'id.' + newBindingInnerId +
                        '.type.' + bindingType +
                        '.name.' + dataName +
                        '.startingRow.' + (m_startingRowTextField.getValue().trim() - 1) +
                        '.startingColumn.' + (m_startingColumnTextField.getValue().trim() - 1) +
                        '.decimals.' + encodedDecimals +
                        '.missings.' + m_missingValuesDropdown.find('option:checked').val();
                bindingTypeEnum = Office.BindingType.Table;
            }

            // Add binding
            Office.context.document.bindings.addFromSelectionAsync(
                    bindingTypeEnum,
                    {id: newBindingId},
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            var binding = asyncResult.value;
                            if (bindingType === 'scalar')
                                m_scalarNameTextField.setValue('');
                            else if (bindingType === 'matrix')
                                m_matrixNameTextField.setValue('');
                            m_bindButton.prop('disabled', true);
                            m_bindingsList.addItem(binding, true);
                            cSuccessMsg.showMessage('The binding for the ' + bindingType + ' "' + dataName + '" was created.');
                        } else {
                            if (bindingType === 'scalar')
                                cErrorMsg.showMessage('Can not create new binding: did you select a portion of text?');
                            else if (bindingType === 'matrix')
                                cErrorMsg.showMessage('Can not create new binding: did you select an entire table?');
                        }
                    }
            );
        });
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
    }

    function getBindingType() {
        return bindingTypeDropdown.find('option:checked').val();
    }

    function onBindingTypeChanged() {
        var bindingType = getBindingType();
        if (bindingType === 'scalar') {
            m_scalarNameTextField.show();
            m_decimalsTextField.show();
            m_matrixNameTextField.hide();
            m_decimalsForColumnsTextField.hide();
            m_startingRowTextField.hide();
            m_startingColumnTextField.hide();
        } else if (bindingType === 'matrix') {
            m_matrixNameTextField.show();
            m_decimalsForColumnsTextField.show();            
            m_startingRowTextField.show();
            m_startingColumnTextField.show();
            m_scalarNameTextField.hide();
            m_decimalsTextField.hide();
        }
        updateBindButtonStatus();
    }

    function updateBindButtonStatus() {
        var bindingType = getBindingType();
        if (bindingType === 'scalar') {
            if (m_scalarNameTextField.isValid() && m_decimalsTextField.isValid())
                m_bindButton.prop('disabled', false);
            else
                m_bindButton.prop('disabled', true);
        }
        else if (bindingType === 'matrix') {
            if (m_matrixNameTextField.isValid()
                && m_decimalsForColumnsTextField.isValid()
                && m_startingRowTextField.isValid()
                && m_startingColumnTextField.isValid()    
            )
                m_bindButton.prop('disabled', false);
            else
                m_bindButton.prop('disabled', true);        
        }
    }

    function closeAllCreateMsg() {
        cSuccessMsg.close();
        cErrorMsg.close();
    }

    // TODO: delete this unused function?
    function getAsyncErrorMessage(code) {
        switch (code) {
            case 2004:
                return 'The table size is too small';
            case 3002: // TODO: insert in the right place
                return 'Not existing binding';
        }
    }
})();