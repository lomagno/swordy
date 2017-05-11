/* global Office, fabric */

"use strict";

(function () {    
    var
        m_topicDropdownElement,
        m_topicContainers;
        
    // The initialize function is run each time the page is loaded.
    Office.initialize = function () {
        $(document).ready(function () {
            // Topic dropdown
            m_topicDropdownElement = $('#topicDropdown');
            new fabric['Dropdown'](m_topicDropdownElement[0]);
            m_topicDropdownElement.find('.ms-Dropdown-select').change(onTopicChanged); 
            
            // Topic containers
            m_topicContainers = $('.topic-container');
            
            // Show initial topic
            showTopic('about');
        });
    };
    
    function showTopic(topic) {
        m_topicContainers.hide();
        m_topicContainers.filter('[data-topic=' + topic + ']').show();
    }
    
    function onTopicChanged() {
        var topic = m_topicDropdownElement.find('option:checked').val();        
        showTopic(topic);        
    }
})();
