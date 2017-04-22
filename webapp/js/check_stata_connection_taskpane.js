/* global Office, fabric, swire */

'use strict';

(function () {
    var
        m_checkStataConnectionButton,
        m_successMsg,
        m_errorMsg,
        m_checkStataConnectionSpinner,
        m_fabricCheckStataConnectionSpinner,
        m_troubleshooting;

    // The initialize function is run each time the page is loaded.
    Office.initialize = function () {
        $(document).ready(function () {
            // Test Stata connection button
            m_checkStataConnectionButton = $('#checkStataConnectionButton');
            new fabric['Button'](m_checkStataConnectionButton[0], onTestStataConnectionButtonClicked);
            
            // "Create" success message
            m_successMsg = new MessageBar('success-msg');

            // "Create" error message
            m_errorMsg = new MessageBar('error-msg'); 
            
            // Spinner
            m_checkStataConnectionSpinner = $('#checkStataConnectionSpinner');
            m_fabricCheckStataConnectionSpinner = new fabric['Spinner'](m_checkStataConnectionSpinner[0]);
            
            // Troubleshooting
            m_troubleshooting = $('#troubleshooting');
            m_troubleshooting.hide();
            
            // Troubleshooting list items
            m_troubleshooting.find('li').click(function(event) {
                var item = $(this);
                if (!$(event.target).hasClass('troubleshooting-item'))    
                    return;
                var instructions = item.find('.troubleshooting-item-instructions');
                instructions.toggleClass('troubleshooting-item-hidden');
                item.toggleClass('troubleshooting-item-open');
            });
            
            checkStataConnection();            
        });
    };
    
    function onTestStataConnectionButtonClicked() {
        m_checkStataConnectionButton.hide();
        m_fabricCheckStataConnectionSpinner.start();
        m_checkStataConnectionSpinner.show();
        checkStataConnection();
    }
    
    function showErrorLayout() {
        m_checkStataConnectionButton.show();
        m_fabricCheckStataConnectionSpinner.stop();
        m_checkStataConnectionSpinner.hide();
        m_troubleshooting.show();
    }
    
    function checkStataConnection() {
        // SWire request
        var swireRequest = {                
            job: [
                {
                    method: '$ping'   
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
                    showErrorLayout();
                    m_errorMsg.showMessage('SWordy can not connect to Stata.');
                    return;
                }                
                if (response.output[0].status !== 'ok') {
                    showErrorLayout();
                    m_errorMsg.showMessage('SWordy can not connect to Stata.');
                    return;                   
                } 
                
                m_errorMsg.close();
                m_fabricCheckStataConnectionSpinner.stop();
                m_checkStataConnectionSpinner.hide();
                m_checkStataConnectionButton.show();
                m_successMsg.showMessage('SWordy can successfully connnect to Stata by using the SWire HTTPS server.');
                m_troubleshooting.hide();
            },
            error: function() {
                showErrorLayout();
                m_errorMsg.showMessage('SWordy can not connect to Stata.');
            }            
        });
    }
})();
