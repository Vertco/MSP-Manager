var customers = [];
var selectedCustomer = "";

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
                populateDatalist(customers);
            }
        };
        fr.readAsText(this.files[0]);
    });

function getSelectedCustomer(value) {
    return customers.find(function (customer) {
        return customer.companyName === value;
    });
}

document
    .getElementById('customer-select')
    .addEventListener('input', function () {
        selectedCustomer = getSelectedCustomer(this.value);
    });

// Define the URLLink custom element extending from HTMLButtonElement
class URLLink extends HTMLButtonElement {
    constructor() {
        super();
    }

    connectedCallback() {
        this.style.cursor = 'pointer';

        // Retrieve the initial URL attribute
        var url = this.getAttribute('url');

        // Update the URL whenever a different customer is selected
        var updateUrl = () => {
            selectedCustomer = getSelectedCustomer(document.getElementById('customer-select').value);
            if (selectedCustomer) {
                var updatedUrl = url.replace(/<microsoftId>/g, encodeURIComponent(selectedCustomer.microsoftId));
                updatedUrl = updatedUrl.replace(/<primaryDomainName>/g, encodeURIComponent(selectedCustomer.primaryDomainName));
                this.setAttribute('url', updatedUrl);
            }
        };

        // Handle the click event
        this.addEventListener('click', function () {
            updateUrl();
            window.open(this.getAttribute('url'), '_blank');
        });

        // Handle changes in the customer selection
        document.getElementById('customer-select').addEventListener('input', updateUrl);
    }
}

// Define the custom element
customElements.define('url-link', URLLink, { extends: 'button' });


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

function populateDatalist() {
    var datalist = document.getElementById('customer-list');

    datalist.innerHTML = '';

    customers.forEach(function (customer) {
        var option = document.createElement('option');
        option.value = customer.companyName;

        datalist.appendChild(option);
    });
}

// Check if customers exist and populate datalist on startup
customers = JSON.parse(localStorage.getItem('customers'));
if (customers && customers.length > 0) {
    populateDatalist();
}
