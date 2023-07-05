var customers = [];

document
    .getElementById('import-button')
    .addEventListener('change', function () {
        var fr = new FileReader();
        fr.onload = function () {
            var csvData = this.result;
            var jsonObj = convertCsvToJson(csvData);

            localStorage.setItem('customers', JSON.stringify(jsonObj));
            customers = JSON.parse(localStorage.getItem('customers'));

            if (customers.length > 0) {
                populateDropdown(customers);
            }
        };
        fr.readAsText(this.files[0]);
    });

function convertCsvToJson(csvData) {
    var lines = csvData.split('\n');
    var headers = lines[0].split(',');
    var jsonArray = [];

    for (var i = 1; i < lines.length; i++) {
        var currentLine = lines[i].split(',');
        var jsonObject = {};

        for (var j = 0; j < headers.length; j++) {
            var propertyName = convertToCamelCase(headers[j]);
            jsonObject[propertyName] = currentLine[j];
        }

        jsonArray.push(jsonObject);
    }

    return jsonArray;
}

function convertToCamelCase(str) {
    return str
        .toLowerCase()
        .replace(/[^a-zA-Z0-9]+(.)/g, (_, chr) => chr.toUpperCase());
}

function populateDropdown(customers) {
    // Get the select element
    var select = document.getElementById('customer-select');

    // Clear existing options
    select.innerHTML = '';

    // Add default placeholder option
    var defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.disabled = true;
    defaultOption.selected = true;
    defaultOption.textContent = 'Select a customer';
    select.appendChild(defaultOption);

    // Populate the dropdown with customer names and data attributes
    customers.forEach(function (customer) {
        var option = document.createElement('option');
        option.text = customer.companyName;
        option.value = customer.microsoftId || customer.primaryDomainName;
        option.dataset.primaryDomainName = customer.primaryDomainName;
        option.dataset.microsoftId = customer.microsoftId;
        select.add(option);
    });

    // Attach event listener for search input
    var searchInput = document.getElementById('search-input');
    searchInput.addEventListener('input', function () {
        var searchValue = this.value.trim().toLowerCase();
        filterDropdown(select, searchValue);
        openDropdown();
    });
}

function filterDropdown(select, searchValue) {
    var options = select.options;
    for (var i = 0; i < options.length; i++) {
        var option = options[i];
        var text = option.text.toLowerCase();
        if (text.indexOf(searchValue) > -1) {
            option.style.display = '';
        } else {
            option.style.display = 'none';
        }
    }
}

function openUrlButton(url) {
    var selectedOption = select.options[select.selectedIndex];
    var selectedMicrosoftId = selectedOption.dataset.microsoftId;
    var selectedPrimaryDomainName = selectedOption.dataset.primaryDomainName;
    url = url.replace('<microsoftId>', encodeURIComponent(selectedMicrosoftId));
    url = url.replace('<primaryDomainName>', encodeURIComponent(selectedPrimaryDomainName));
    window.open(url, '_blank');
}

// Get the select element
var select = document.getElementById('customer-select');

// Define the URLLink custom element extending from HTMLButtonElement
class URLLink extends HTMLButtonElement {
    constructor() {
        super();
    }

    connectedCallback() {
        var url = this.getAttribute('url');
        this.addEventListener('click', function () {
            openUrlButton(url);
        });
        this.style.cursor = 'pointer';
    }
}
// Define the custom element
customElements.define('url-link', URLLink, { extends: 'button' });

// Open the dropdown
function openDropdown() {
    select.parentNode.classList.add('open');
    select.click(); // Trigger click event on the select element
}

// Close the dropdown
function closeDropdown() {
    select.parentNode.classList.remove('open');
}

// Add event listener for customer select
select.addEventListener('change', function () {
    var selectedOption = this.options[this.selectedIndex];
    if (selectedOption.value !== '') {
        var urlButtons = document.querySelectorAll('button[is="url-link"]');
        urlButtons.forEach(function (button) {
            button.hidden = false;
        });
    } else {
        var urlButtons = document.querySelectorAll('button[is="url-link"]');
        urlButtons.forEach(function (button) {
            button.hidden = true;
        });
    }
});

// Check if customers exist and populate dropdown on startup
customers = JSON.parse(localStorage.getItem('customers'));
if (customers && customers.length > 0) {
    populateDropdown(customers);
}
