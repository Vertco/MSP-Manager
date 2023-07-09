var customers = [];
var selectedCustomer = "";

// Check if customers exist and populate datalist on startup
customers = JSON.parse(localStorage.getItem('customers'));
if (customers && customers.length > 0) {
    populateDatalist(customers);
}

async function handleFiles(files) {
    for (const file of files) {
        file.handle = file;
    }
}

if ("launchQueue" in window) {
    window.launchQueue.setConsumer((launchParams) => {
        if (launchParams.files && launchParams.files.length) {
            var fr = new FileReader();
            fr.onload = function () {
                var csvData = this.result;
                var jsonObj = convertCsvToJson(csvData); // Make sure to define convertCsvToJson

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
            fr.readAsText(launchParams.files[0].handle);
            handleFiles(launchParams.files);
        }
    });
}

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

document
    .getElementById('clear-button')
    .addEventListener('click', function () {
        document.getElementById('customer-select').value = "";

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

function populateDatalist() {
    var datalist = document.getElementById('customer-list');

    datalist.innerHTML = '';

    customers.forEach(function (customer) {
        var option = document.createElement('option');
        option.value = customer.companyName;

        datalist.appendChild(option);
    });
}

// Define the URLLink custom element extending from HTMLButtonElement
class URLLink extends HTMLButtonElement {
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

// Define the custom element
customElements.define('url-link', URLLink, { extends: 'button' });