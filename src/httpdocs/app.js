// Variables definition
var customers = [];
var selectedCustomer = "";

// Check if customers exist and populate datalist on startup
customers = JSON.parse(localStorage.getItem('customers'));
if (customers && customers.length > 0) {
    populateDatalist(customers);
}

// Function to read the currently selected customer from the dropdown
function getSelectedCustomer(value) {
    return customers.find(function (customer) {
        return customer.companyName === value;
    });
}

// Function to convert CSV to JSON data
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

// Function to convert text to camelcase
function convertToCamelCase(str) {
    return str
        .toLowerCase()
        .replace(/[^a-zA-Z0-9]+(.)/g, (_, chr) => chr.toUpperCase());
}

// Function for populating the datalist for the customer dropdown
function populateDatalist(customers) {
    var datalist = document.getElementById("customer-list");

    datalist.innerHTML = "";

    if (customers && customers.length > 0) {
        customers.forEach(function (customer) {
            var option = document.createElement("option");
            option.value = customer.companyName;

            datalist.appendChild(option);
        });
    }
}

// Function to handle the file opening
function handleFileOpen(blob) {
    const reader = new FileReader();

    reader.onload = function (e) {
        const csvData = e.target.result;
        const jsonData = convertCsvToJson(csvData);

        jsonData.sort(function (a, b) {
            var nameA = a.companyName.toUpperCase();
            var nameB = b.companyName.toUpperCase();
            if (nameA < nameB) {
                return -1;
            }
            if (nameA > nameB) {
                return 1;
            }
            return 0;
        });

        localStorage.setItem("customers", JSON.stringify(jsonData));
        const customers = JSON.parse(localStorage.getItem("customers"));

        if (customers.length > 0) {
            populateDatalist(customers);
        }
    };

    reader.readAsText(blob);
}

// Function to add event listener for the import button 
function addImportListener() {
    document
        .getElementById('import-button')
        .addEventListener('change', function () {
            var fr = new FileReader();
            fr.onload = function () {
                var csvData = this.result;
                var jsonObj = convertCsvToJson(csvData);

                jsonObj.sort(function (a, b) {
                    var nameA = a.companyName.toUpperCase();
                    var nameB = b.companyName.toUpperCase();
                    if (nameA < nameB) {
                        return -1;
                    }
                    if (nameA > nameB) {
                        return 1;
                    }
                    return 0;
                });

                localStorage.setItem('customers', JSON.stringify(jsonObj));
                customers = JSON.parse(localStorage.getItem('customers'));

                if (customers.length > 0) {
                    populateDatalist(customers);
                }
            };
            fr.readAsText(this.files[0]);
        });
}

// Check if the File Handler API is supported
if ("launchQueue" in window) {
    // Handle the file when the PWA is launched with a file
    console.debug("Launch queue detected");
    window.launchQueue.setConsumer(async (launchParams) => {
        if (launchParams.files.length > 0) {
            const fileHandle = launchParams.files[0];
            const file = await fileHandle.getFile();
            handleFileOpen(file);
        }
    });
    addImportListener();
} else {
    // Fallback for browsers without File Handler API
    console.debug("No launch queue detected");
    addImportListener();
}

// Add event listener for customer dropdown
document
    .getElementById('customer-select')
    .addEventListener('input', function () {
        selectedCustomer = getSelectedCustomer(this.value);
    });

// Add event listener for clear button
document
    .getElementById('clear-button')
    .addEventListener('click', function () {
        document.getElementById('customer-select').value = "";

    });

// Define the mspButton custom element extending from HTMLButtonElement
class mspButton extends HTMLButtonElement {
    constructor() {
        super();
    }

    connectedCallback() {
        this.style.cursor = 'pointer';

        // Retrieve the initial URL attribute
        var url = this.getAttribute('url');
        var selectedCustomer = null;

        // Update the URL whenever a different customer is selected
        var updateUrl = () => {
            var previousCustomer = selectedCustomer;
            selectedCustomer = getSelectedCustomer(document.getElementById('customer-select').value);

            if (selectedCustomer) {
                var updatedUrl = url.replace(/<microsoftId>/g, encodeURIComponent(selectedCustomer.microsoftId));
                updatedUrl = updatedUrl.replace(/<primaryDomainName>/g, encodeURIComponent(selectedCustomer.primaryDomainName));
                this.setAttribute('url', updatedUrl);
            }

            // Enable the element when selectedCustomer changes from empty to not empty
            if (!previousCustomer && selectedCustomer) {
                this.disabled = false;
            }
        };

        // Handle the click event
        this.addEventListener('click', function () {
            updateUrl();
            window.open(this.getAttribute('url'), '_blank');
        });

        // Handle changes in the customer selection
        document.getElementById('customer-select').addEventListener('input', updateUrl);

        // Disable the element initially
        this.disabled = true;
    }
}

// Define the mspButton element
customElements.define('msp-button', mspButton, { extends: 'button' });