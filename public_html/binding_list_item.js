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
    
    function onSyncDataButtonClicked() {
        // SWire request
        var request;
        if (m_type === 'scalar')
            request = {
                job: [
                    {
                        method: 'com.stata.sfi.Scalar.getValue',
                        args: [m_name]
                    }
                ]
            };
        else if (m_type === 'matrix')
            request = {
                job: [
                    {
                        method: '$getMatrix',
                        args: {
                            name: m_name
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
                
                // Stata data
                var data = response.output[0].output;
                
                // Not found data
                var bindingsWithNoFoundData = [];
                if (data === null)
                    bindingsWithNoFoundData.push(m_bindingId);
                
                // Report
                var report = {
                    count: data === null ? 0 : 1,
                    notFound: bindingsWithNoFoundData,
                    syncOk: [],
                    syncNotOk: [],
                    syncNotOkErrorCodes: []
                };         
                
                if (data !== null) {                                    
                    // Update document
                    if (m_type === 'scalar')
                        syncScalarData({
                            bindingId: m_bindingId,
                            scalarValue: data,
                            report: report,
                            onComplete: onSyncCompleted                            
                        });                    
                    else if (m_type === 'matrix')
                        syncMatrixData({
                            bindingId: m_bindingId,
                            matrixData: data,
                            report: report,
                            onComplete: onSyncCompleted                              
                        }); 
                }
                else
                    onSyncCompleted(report);
            },
            error: function (/* jqXHR, textStatus, errorThrown */) {
                console.error('Cannot communicate with Stata'); // TODO: manage this
                //mErrorMsg.showMessage('Cannot communicate with Stata'); // TODO: manage this
            }
        });        
    } 
    
    function onSyncCompleted(report) {
        console.log(report);
    }
    
    var
        m_self = this,
        m_gui,
        m_bindingId = pars.bindingId,
        m_name,
        m_type,
        m_decimals;
    
    (function() {
        var bindingProperties = getBindingProperties(pars.bindingId);
        m_name = bindingProperties.name;
        m_type = bindingProperties.type;
        m_decimals = bindingProperties.decimals;
        
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
        decimalsText.text('Decimals: ' + m_decimals);
        m_gui.append(decimalsText);
        
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
    })();
}
