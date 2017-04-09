/* global fabric, Office */

function BindingsList(elementId) {
    this.update = function() {        
        Office.context.document.bindings.getAllAsync(function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                // Number of bindings
                m_nBindings = asyncResult.value.length;
                
                // Empty bindings list
                m_container.empty();                
                
                // Create items
                for (var i in asyncResult.value) {
                    var binding = asyncResult.value[i];                
                    var bindingId = binding.id;
                    var bindingListItem = createItem(bindingId);
                    m_container.append(bindingListItem);
                    binding.addHandlerAsync(Office.EventType.BindingSelectionChanged, onBindingSelectionChanged);
                    binding.addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
                }
                m_self.sort(m_sortBy, m_order);
                initFabric();      
            }
            else
                console.error('Cannot update the bindings list');
        });
    };    
    
    this.sort = function(sortBy, order) {
        var bindingsListItems = m_container.find('li');
        bindingsListItems.sort(function(a, b) {
            var bindingId1 = $(a).data('binding');
            var bindingProperties1 = getBindingProperties(bindingId1);
            var property1 = bindingProperties1[sortBy];
            var bindingId2 = $(b).data('binding');
            var bindingProperties2 = getBindingProperties(bindingId2); 
            var property2 = bindingProperties2[sortBy];
            if (order === 'asc')
                return property1.localeCompare(property2);
            else if (order === 'desc')
                return property2.localeCompare(property1);
        });        
        bindingsListItems.detach().appendTo(m_container);        
    };            
    
    function createItem(bindingId) {
        var bindingProperties = getBindingProperties(bindingId);
        
        // List item
        var listItem = $('<li data-binding="' + bindingId  + '" class="ms-ListItem is-selectable" tabindex="0"></li>'); // TODO: what tabindex stands for?
        
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
        var actionsContainer = $('<div class="ms-ListItem-actions bindingActions"></div>');
        listItem.append(actionsContainer);
        
        // Delete button
        var deleteButton = $('<div class="ms-ListItem-action" title="Delete binding"><i class="ms-Icon ms-Icon--Delete"></i></div>');
        deleteButton.click(function() {onDeleteBindingButtonClicked(listItem);});
        actionsContainer.append(deleteButton);
        
        // Sync data button
        var syncDataButton = $('<div class="ms-ListItem-action" title="Sync data"><i class="ms-Icon ms-Icon--Sync"></i></div>');
        //syncDataButton.click(onIndividualSyncDataButtonClicked);                
        actionsContainer.append(syncDataButton);                
        
        return listItem;
    } 
    
    function onDeleteBindingButtonClicked(listItem) {
        var bindingId = listItem.data('binding');
        Office.context.document.bindings.releaseByIdAsync(bindingId, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
                removeItem(listItem);
            else
                console.error('Cannot update the bindings list');
        });        
    }      
    
    function onBindingSelectionChanged(eventArgs) {
        // Binding ID
        var bindingId = eventArgs.binding.id;
        
        // Unselect all items
        m_container.find('li').removeClass('is-unread');
        
        // Selected item
        var bindingSelectionItem = m_container.find('li[data-binding="' + bindingId + '"]');
        bindingSelectionItem.addClass('is-unread');
        
        // Scroll to item
        $('html, body').animate({scrollTop: bindingSelectionItem.offset().top}, 200);        
    } 
    
    function onBindingDataChanged(eventArgs) {
        var bindingId = eventArgs.binding.id;
        var bindingListItem = m_container.find('li[data-binding="' + bindingId + '"]');
        Office.context.document.bindings.getAllAsync(
            {asyncContext: m_nBindings},
            function(asyncResult) { 
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    if (asyncResult.value.length < asyncResult.asyncContext)
                        removeItem(bindingListItem);
                }
                else
                    console.error('Cannot update the bindings list');
            }
        );                
    }      
    
    function removeItem(listItem) {
        listItem.fadeOut(200, function() {
            listItem.remove();
            m_nBindings--;
        });        
    }
    
    function initFabric() {
        new fabric['List'](m_container[0]);
    }
    
    function onDocumentSelectionChanged(/* eventArgs */) { // TODO: is eventArgs useful?
        m_container.find('li').removeClass('is-unread');
    }       
    
    // Variables
    var
        m_self = this,
        m_container = $('#' + elementId),
        m_nBindings = 0,
        m_sortBy = 'name',
        m_order = 'asc';
    
    // Document selection changed event
    Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onDocumentSelectionChanged);     
    
    // Update list
    this.update();
}