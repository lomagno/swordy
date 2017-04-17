/* global Office, fabric */

'use strict';

(function () {      
    var
        m_syncDocumentButton,
        m_successMsg,
        m_errorMsg,
        m_synchingSpinner,
        m_fabricSynchingSpinner;
    
    Office.initialize = function (/* reason */) {
        $(document).ready(function () { 
            // Sync document button
            m_syncDocumentButton = $('#syncDocumentButton');
            new fabric['Button'](m_syncDocumentButton[0], onSyncDocumentButtonClicked);            
            m_syncDocumentButton.hide();
            
            // "Create" success message
            m_successMsg = new MessageBar('success-msg');

            // "Create" error message
            m_errorMsg = new MessageBar('error-msg'); 
            
            // Spinner
            m_synchingSpinner = $('#synchingSpinner');
            m_fabricSynchingSpinner = new fabric['Spinner'](m_synchingSpinner[0]);
            
            syncDocument();
        });
    };  
    
    function syncDocument() { 
        m_syncDocumentButton.hide();
        m_synchingSpinner.show();
        m_fabricSynchingSpinner.start();
        Office.context.document.bindings.getAllAsync(function (asyncResult) {
            var bindingIds = [];
            for (var i in asyncResult.value)
                bindingIds.push(asyncResult.value[i].id);
            syncBindings({
                bindingIds: bindingIds,
                onComplete: onSyncComplete
            });
        });
    }
    
    function onSyncComplete(report) {        
        var textualReport = getTextualReport(report);
        var status = textualReport.status;
        var messages = textualReport.messages;
        if (status === 'ok') {
            m_successMsg.showMessage(messages[0]);
        }
        else {
            if (messages.length === 1)
                m_errorMsg.showMessage(messages[0]);
            else {
                m_errorMsg.reset();
                m_errorMsg.appendList();
                for (var i in messages)
                    m_errorMsg.appendListItem(messages[i]);
                m_errorMsg.show();
            }
        }
        m_syncDocumentButton.show();
        m_synchingSpinner.hide();
        m_fabricSynchingSpinner.stop();
    }
    
    function onSyncDocumentButtonClicked() {
        syncDocument();
    }
})();