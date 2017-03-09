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
        });
    };
    
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
                }
            });
        });;
    }
    
    function initFabricComponents() {
        // Init dropdowns
        var DropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');
        for (var i = 0; i < DropdownHTMLElements.length; ++i)
            new fabric['Dropdown'](DropdownHTMLElements[i]);

        // Init text fields
        var TextFieldElements = document.querySelectorAll(".ms-TextField");
        for (var i = 0; i < TextFieldElements.length; i++) {
            new fabric['TextField'](TextFieldElements[i]);
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
    
    // TODO: delete this?
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
})();