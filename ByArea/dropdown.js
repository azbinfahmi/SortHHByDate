// script.js
const items = ["Apple", "Banana", "Cherry", "Date", "Elderberry", "Fig", "Grape"];
var selectedItems = [];

document.addEventListener('click', function(event) {
  const dropdown = document.getElementById("dropdownList");
  if (!event.target.matches('#searchInput') && !event.target.matches('.dropdown-content div')) {
    dropdown.classList.remove('show');
  }
});

function toggleDropdown() {
  const dropdown = document.getElementById("dropdownList");
  dropdown.classList.toggle('show');
}

function populateDropdown() {
  const list = document.getElementById("dropdownList");
  items.forEach(item => {
    const div = document.createElement("div");
    div.textContent = item;
    div.onclick = function() { toggleItemSelection(this, item); };
    list.appendChild(div);
  });
}

function toggleItemSelection(div, item) {
  const index = selectedItems.indexOf(item);
  if (index > -1) {
    selectedItems.splice(index, 1); // Remove item from list
    div.classList.remove('selected');
  } else {
    selectedItems.push(item); // Add item to list
    div.classList.add('selected');
  }
  updateInputField();
}

function updateInputField() {
  const input = document.getElementById("searchInput");
  input.value = selectedItems.join(', '); // Join array into string separated by commas
}

populateDropdown();
