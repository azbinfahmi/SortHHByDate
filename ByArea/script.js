var selectedArea = {}
function showRowsAndSheets(){
    selectedArea = {}
    for(let i=0; i < selectedItems.length; i++){
        selectedArea[selectedItems[i]] = Area[selectedItems[i]]
    }
    displayArea(selectedArea)
}