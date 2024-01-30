let index_1 = 0, index_2 = 0, RangeDate =[]

function toggleDateRange(){
    const rangeSelect = document.querySelector('.Range');
    const dateRangeRadio = document.getElementById('dateRange');
    const singleDateRadio = document.getElementById('singleDate');
    const dateSelect = document.getElementById('dateSelect');
    const dateSelect_2 = document.getElementById('dateSelect_2');

    // Toggle the display property based on the selected radio button
    rangeSelect.style.display = dateRangeRadio.checked ? 'block' : 'none';

    if (singleDateRadio.checked) {
        dateSelect_2.value = dateSelect.value;
    }
}

function handleDateSelectChange(selectElement) {
    RangeDate =[]
    var selectedDateString_2 = document.getElementById('dateSelect_2');
    const selectedValue = selectElement.value;
    selectedDateString_2.innerHTML =''
    //cari index ke berapa
    for (var i = 0; i< uniqueDates.length; i++){
        if(selectedValue == uniqueDates[i]){
            index_1 = Number(i)
            break;
        }
    }
    console.log('index_1',index_1)
    //display value untuk date  Range
    for(var i = 0; i< uniqueDates.length; i++){
        if( i >= index_1){
            const option = document.createElement('option');
            option.value = uniqueDates[i];
            option.textContent = uniqueDates[i];
            selectedDateString_2.appendChild(option);
        }
    }
}

function handleDateSelectChange_2(selectElement){
    RangeDate =[]
    const selectedValue = selectElement.value;
    for (var i = 0; i< uniqueDates.length; i++){
        if( i >= index_1){
            RangeDate.push(formattedStringToExcelDate(uniqueDates[i]))
        }
        if(selectedValue == uniqueDates[i]){
            index_2 = i
            break;
        }
    }
    RangeDate = RangeDate.flat()
    //console.log('RangeDate: ',RangeDate)    
}
