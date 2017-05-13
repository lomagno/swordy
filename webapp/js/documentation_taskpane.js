/* global Office, fabric */

// "use strict";

(function () {    
    var
        m_topicDropdownElement,
        m_topicDropdownItemsMap = new Map,
        m_topicContainers;
        
    // The initialize function is run each time the page is loaded.
    Office.initialize = function () {
        $(document).ready(function () {
            var initialTopic = 'local_installation';
            
            // Topic dropdown
            m_topicDropdownElement = $('#topicDropdown');
            var fabricTopicDropdown = new fabric['Dropdown'](m_topicDropdownElement[0]);
            for (var i=1; i<fabricTopicDropdown._dropdownItems.length; ++i) {
                var key = fabricTopicDropdown._dropdownItems[i].oldOption.value;
                var value = fabricTopicDropdown._dropdownItems[i].newItem;
                m_topicDropdownItemsMap.set(key, value);
            }
            setTopicDropdownItem(initialTopic);
            m_topicDropdownElement.find('.ms-Dropdown-select').change(onTopicChanged);             
            
            // Topic containers
            m_topicContainers = $('.topic-container');
            
            // Topic links
            $('.topic-link').click(onTopicLinkClicked);
            
            // Show initial topic
            showTopic(initialTopic);
        });
    };
    
    function setTopicDropdownItem(topic) {
        $(m_topicDropdownItemsMap.get(topic)).click();
    }
    
    function showTopic(topic) {
        m_topicContainers.hide();
        m_topicContainers.filter('[data-topic=' + topic + ']').show();
    }
    
    function onTopicChanged() {
        var topic = m_topicDropdownElement.find('option:checked').val();        
        showTopic(topic);        
    }
    
    function onTopicLinkClicked() {
        var topicLink = $(this);
        var topic = topicLink.data('topic-target');
        setTopicDropdownItem(topic);
    }
})();
