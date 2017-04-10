/* global fabric, Office, swire */

/*
 * pars:
 * - elementId
 * - onListStatusChanged
 */
function BindingsList(pars) {
    this.update = function() {
        Office.context.document.bindings.getAllAsync(function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {                
                // Empty bindings list                
                m_container.empty();
                m_items = [];
                
                // Create items
                for (var i in asyncResult.value) {
                    var binding = asyncResult.value[i]; 
                    m_self.addItem(binding, false);
                }
                
                // Sort and filter
                m_self.sort(m_sortBy, m_order);
                m_self.filter(m_filterString);
                
                updateListStatus();
            }
            else
                console.error('Cannot update the bindings list');
        });
    };
    
    this.addItem = function(binding, sortAndFilter) {
        var item = new BindingListItem({
            bindingId: binding.id,
            onDeleteButtonClicked: onDeleteItemButtonClicked,
            onCheckedStatusChanged: onItemCheckedStatusChanged
        });
        m_container.append(item.getGui());
        new fabric['ListItem'](item.getGui()[0]);
        m_items.push(item);
        m_nBindings++;        
        binding.addHandlerAsync(Office.EventType.BindingSelectionChanged, onBindingSelectionChanged);
        binding.addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);        
        
        if (sortAndFilter) {
            m_self.sort(m_sortBy, m_order);
            m_self.filter(m_filterString); 
            updateListStatus();
        }
    };
    
    this.sort = function(sortBy, order) {
        m_sortBy = sortBy;
        m_order = order;
        
        // Sort array of items
        m_items.sort(function(item1, item2) {
            var property1, property2;
            if (sortBy === 'name') {
                property1 = item1.getName();
                property2 = item2.getName();
            }
            else if (sortBy === 'type') {
                property1 = item1.getType();
                property2 = item2.getType();                
            }
            if (order === 'asc')
                return property1.localeCompare(property2);
            else
                return property2.localeCompare(property1);
        });
        
        // Detach items from the DOM
        for (var i in m_items) {
            m_items[i].getGui().detach();
        }        
        
        // Populate list of items
        for (var i in m_items)
            m_container.append(m_items[i].getGui());
    };     
    
    this.filter = function(filterString) {
        m_filterString = filterString;
        if (filterString !== '') {
            for (var i in m_items) {
                var item = m_items[i];
                var name = item.getName();
                if (name.indexOf(filterString) === -1)
                    item.hide();
                else
                    item.show();
            }
        }
        else {
            for (var i in m_items)
                m_items[i].show();
        }
    };

    this.checkUncheckAll = function() {
        var newSelectionStatus = !isAllChecked();
        for (var i in m_items)
            m_items[i].setChecked(newSelectionStatus);
    };
    
    this.getCheckedItems = function() {
        var checkedItems = [];
        for (var i in m_items) {
            var item = m_items[i];
            if (item.isChecked())
                checkedItems.push(item);
        }
        return checkedItems;
    };
    
    this.deleteCheckedItems = function() {
        var checkedItems = m_self.getCheckedItems();
        for (var i in checkedItems) {
            var item = checkedItems[i];
            var bindingId = item.getBindingId();
            Office.context.document.bindings.releaseByIdAsync(
                bindingId, {asyncContext: item}, function (asyncResult) { 
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded)                
                        removeItem(asyncResult.asyncContext);
                    else
                        console.error('Error deleting binding');
                }
            );             
        }
    };
    
    function onItemCheckedStatusChanged() {
        updateListStatus();
    }
    
    function getListStatus() {
        if (m_items.length === 0)
            return {
                populated: false,
                selection: 'nothing'
            };
        ;
        
        var somethingIsChecked = false;
        var areAllChecked = true;
        for (var i in m_items) {
            if (m_items[i].isChecked())
                somethingIsChecked = true;            
            else
                areAllChecked = false;
        }       
        
        var selection;
        if (areAllChecked)
            selection = 'all';
        else if (somethingIsChecked && !areAllChecked)
            selection = 'something';
        else
            selection = 'nothing';
        
        return {
            populated: true,
            selection: selection            
        };
    }
    
    function isAllChecked() {
        for (var i in m_items)
            if (m_items[i].isChecked() === false)
                return false;
        return true;
    }    
    
    function onDeleteItemButtonClicked(item) {
        var bindingId = item.getBindingId();
        Office.context.document.bindings.releaseByIdAsync(bindingId, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
                removeItem(item);
            else
                console.error('Cannot update the bindings list');
        }); 
    }      
    
    function onBindingSelectionChanged(eventArgs) {
        // Binding ID
        var bindingId = eventArgs.binding.id;
        
        // Unhighlight all items
        unhighlightAllItems();
        
        // Corresponding item
        var item = getItemByBindingId(bindingId);
        if (item === null)
            return;
        
        // Highlight item
        item.setHighlighted(true);
        
        // Scroll to item
        $('html, body').animate({scrollTop: item.getGui().offset().top}, 200);        
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
    
    function getItemByBindingId(bindingId) {
        for (var i in m_items) {
            if (m_items[i].getBindingId() === bindingId)
                return m_items[i];
        }
        return null;
    }
    
    function removeItem(item) {
        item.fadeOut(function() {
            // Remove item from array of items
            m_items.splice(m_items.indexOf(item), 1);
            m_nBindings--;
            
            updateListStatus();
        });        
    }   
    
    function onDocumentSelectionChanged(/* eventArgs */) { // TODO: is eventArgs useful?
        unhighlightAllItems();
    }       
    
    function unhighlightAllItems() {
        for (var i in m_items)
            m_items[i].setHighlighted(false);
    }
    
    function updateListStatus() {
        var m_oldListStatus = m_listStatus;
        m_listStatus = getListStatus();
        if (m_listStatus.populated !== m_oldListStatus.populated
            || m_listStatus.selection !== m_oldListStatus.selection)
            pars.onListStatusChanged(m_listStatus);        
    }
    
    // Variables
    var
        m_self = this,
        m_container = $('#' + pars.elementId),
        m_items = [],
        m_nBindings = 0,
        m_sortBy = 'name',
        m_order = 'asc',
        m_filterString = '',
        m_listStatus = {
            populated: false,
            selection: 'nothing'
        };
    
    // Document selection changed event
    Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onDocumentSelectionChanged);     
    
    // Update list
    this.update();
}