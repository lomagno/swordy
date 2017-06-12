/*
 * pars:
 * - bindingId
 * - onDeleteButtonClicked
 * - onCheckedStatusChanged
 */
/* global swire */

function BindingListItem(pars) {        
    this.getGui = function() {
        return m_gui;
    };
    
    this.getBindingId = function() {
        return m_bindingId;
    };
    
    this.getName = function() {
        return m_name;
    };
    
    this.getType = function() {
        return m_type;
    };
    
    this.getDecimals = function() {
        return m_decimals;
    };            
    
    this.getMissingValues = function() {
        return m_missingValues;
    };
    
    this.show = function() {
        m_gui.show();
    };
    
    this.hide = function() {
        m_gui.hide();
    };
    
    this.fadeOut = function(onComplete) {
        m_gui.fadeOut(200, onComplete);
    };
    
    this.setChecked = function(status) {
        if (status)
            m_gui.addClass('is-selected');            
        else
            m_gui.removeClass('is-selected');
    };
    
    this.isChecked = function() {
        return m_gui.hasClass('is-selected');
    };
    
    this.setHighlighted = function(status) {
        if (status)
            m_gui.addClass('is-unread');
        else
            m_gui.removeClass('is-unread');
    };
    
    function createSuccessMessageBarHtml() {
        return '<div id="success-msg" class="ms-MessageBar ms-MessageBar--success ms-u-slideDownIn20">' +
            '<div class="ms-MessageBar-content">' +
                '<div class="ms-MessageBar-icon">' +
                    '<i class="ms-Icon ms-Icon--Completed"></i>' +
                '</div>' +
                '<div class="ms-MessageBar-text">' +
                    '<div class="mb-content"></div>' +
                    '<a class="mb-close-link ms-Link" href="#">Close this</a> ' +
                '</div>' +
            '</div>' +
        '</div>';
    }
    
    function createErrorMessageBarHtml() {
        return '<div id="error-msg" class="ms-MessageBar ms-MessageBar--error ms-u-slideDownIn20">' +
                '<div class="ms-MessageBar-content">' +
                    '<div class="ms-MessageBar-icon">' +
                        '<i class="ms-Icon ms-Icon--StatusErrorFull"></i>' +
                    '</div>' +
                    '<div class="ms-MessageBar-text">' +
                        '<div class="mb-content"></div>' +
                        '<a class="mb-close-link ms-Link" href="#">Close this</a>' +
                    '</div>' +
                '</div>' +
            '</div>';
    }
    
    function onSyncDataButtonClicked() {
        syncBindings({
            bindingIds: [m_bindingId],
            onComplete: onSyncCompleted
        });          
    } 
    
    function onSyncCompleted(report) {
        var textualReport = getTextualReport(report);
        var status = textualReport.status;
        var messages = textualReport.messages;
        if (status === 'ok') {
            m_successMessageBar.showMessage(messages[0]);
        }
        else {
            if (messages.length === 1)
                m_errorMessageBar.showMessage(messages[0]);
            else {
                m_errorMessageBar.reset();
                m_errorMessageBar.appendList();
                for (var i in messages)
                    m_errorMessageBar.appendListItem(messages[i]);
                m_errorMessageBar.show();
            }
        }
    }
    
    var
        m_self = this,
        m_gui,
        m_bindingId = pars.bindingId,
        m_name,
        m_type,
        m_decimals,
        m_missingValues,
        m_successMessageBar,
        m_errorMessageBar;
    
    (function() {
        var bindingProperties = getBindingProperties(pars.bindingId);
        m_name = bindingProperties.name;
        m_type = bindingProperties.type;
        m_decimals = bindingProperties.decimals;
        m_missingValues = bindingProperties.missings;
        
        // List item
        m_gui = $('<li data-binding="' + m_bindingId  + '" class="ms-ListItem is-selectable" tabindex="0"></li>'); // TODO: what tabindex stands for?
        
        // Primary text
        var primaryText = $('<span class="ms-ListItem-primaryText"></span>');
        primaryText.text(m_name);
        m_gui.append(primaryText);
        
        // Type
        var typeText = $('<span class="ms-ListItem-tertiaryText"></span>');
        typeText.text('Type: ' + m_type);
        m_gui.append(typeText);
        
        // Decimals
        var decimalsText = $('<span class="ms-ListItem-tertiaryText"></span>');
        var decimalsContent;
        if (m_type === 'scalar')
            decimalsContent = m_decimals;
        else
            decimalsContent = m_decimals.join(', ');
        decimalsText.text('Decimals: ' + decimalsContent);
        m_gui.append(decimalsText);
        
        // Missing values
        var missingValuesText = $('<span class="ms-ListItem-tertiaryText"></span>');
        var missingsTypeText = '';
        
        if (m_missingValues.indexOf('special_') === 0) { // starts with "special_"
            switch (m_missingValues) {
                case 'special_ieee754':
                    missingsTypeText =  'IEEE 754';
                    break;
                case "special_letters":
                    missingsTypeText =  'letters';
                    break;
                case "special_pletters":
                    missingsTypeText =  'letters with parentheses';
                    break;
                case "special_dot":
                    missingsTypeText =  ".";
                    break;
                default:
                    missingsTypeText = 'unknown';
                    break;
            }                
        }
        else if (m_missingValues.indexOf('string_') === 0) // starts with "string_"
            missingsTypeText =  m_missingValues.substring(7);        

        missingValuesText.text('Missings: ' + missingsTypeText);
        m_gui.append(missingValuesText);
        
        // Selection checkbox
        var checkbox = $('<div class="ms-ListItem-selectionTarget"></div>');
        checkbox.click(function(e) {
            e.stopImmediatePropagation();
            m_self.setChecked(!(m_self.isChecked()));
            pars.onCheckedStatusChanged();
        });
        m_gui.append(checkbox);
        
        // Actions
        var actionsContainer = $('<div class="ms-ListItem-actions bindingActions"></div>');
        m_gui.append(actionsContainer);
        
        // Delete button
        var deleteButton = $('<div class="ms-ListItem-action" title="Delete binding"><i class="ms-Icon ms-Icon--Delete"></i></div>');
        deleteButton.click(function() {pars.onDeleteButtonClicked(m_self);});
        actionsContainer.append(deleteButton);
        
        // Sync data button
        var syncDataButton = $('<div class="ms-ListItem-action" title="Sync data"><i class="ms-Icon ms-Icon--SetAction"></i></div>');
        syncDataButton.click(onSyncDataButtonClicked);   
        actionsContainer.append(syncDataButton);
        
        // Success message bar
        var successMessageBarHtml = $(createSuccessMessageBarHtml());
        m_gui.append(successMessageBarHtml);
        m_successMessageBar = new MessageBar(successMessageBarHtml);
        
        // Error message bar
        var errorMessageBarHtml = $(createErrorMessageBarHtml());
        m_gui.append(errorMessageBarHtml);
        m_errorMessageBar = new MessageBar(errorMessageBarHtml);        
    })();
}
