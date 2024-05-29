// Variables definition
var customers = [];
var selectedCustomer = '';

// Constants definition
const defaultOrder = [["admincenter", "entra", "intune", "exchange", "sharepoint", "azure", "teams", "defender", "purview"], []];
const draggables = document.querySelectorAll('.draggable')
const containers = document.querySelectorAll('.container')
const dropOverlay = document.getElementById('drop-overlay');

// Check if customers exist and populate datalist on startup
customers = JSON.parse(localStorage.getItem('customers'));
if (customers && customers.length > 0) {
    populateDatalist(customers);
}

// Ckeck if savedOrder exists and restore order on startup
const savedOrder = localStorage.getItem('savedOrder');
if (savedOrder) {
    const containerIds = JSON.parse(savedOrder);
    containers.forEach((container, index) => {
        containerIds[index].forEach(draggableId => {
            const draggable = document.getElementById(draggableId);
            if (draggable) {
                container.appendChild(draggable);
            } else {
                console.warn(`Draggable element with ID ${draggableId} not found.`);
            }
        });
    });
}

// Function to read the currently selected customer from the dropdown
function getSelectedCustomer(value) {
    const companyName = value.split(" | ")[0]; // Split and get the companyName
    if (customers) {
        return customers.find(function (customer) {
            return customer.companyName === companyName;
        });
    }
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

            // Rename the columns as per the requirements
            if (propertyName === 'customerid') {
                propertyName = 'microsoftId';
            } else if (propertyName === 'defaultdomainname') {
                propertyName = 'primaryDomainName';
            } else if (propertyName === 'displayname') {
                propertyName = 'companyName';
            }

            // Filter the columns so that only microsoftId, companyName and primaryDomainName remain
            if (['microsoftId', 'companyName', 'primaryDomainName'].includes(propertyName)) {
                jsonObject[propertyName] = currentLine[j].replace(/^"|"$/g, '').trim();
            }
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
    var datalist = document.getElementById('customer-list');

    datalist.innerHTML = '';

    if (customers && customers.length > 0) {
        customers.forEach(function (customer) {
            var option = document.createElement('option');
            option.value = `${customer.companyName} | ${customer.primaryDomainName}`;

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

        localStorage.setItem('customers', JSON.stringify(jsonData));
        customers = JSON.parse(localStorage.getItem('customers'));

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

// Function to copy tenant ID to clipboard
function copyTenantID() {
    if (selectedCustomer) {
        tenantId = selectedCustomer.microsoftId;
        navigator.clipboard.writeText(tenantId);
        document.querySelector('#tenant-id>p').innerHTML = 'Copied!';
        setTimeout(function () {
            document
                .querySelector('#tenant-id>p').innerHTML = selectedCustomer.microsoftId
        }, 1000);
    }
}

// Function to copy tenant domain to clipboard
function copyTenantDomain() {
    if (selectedCustomer) {
        tenantDomain = selectedCustomer.primaryDomainName;
        navigator.clipboard.writeText(tenantDomain);
        document.querySelector('#tenant-domain>p').innerHTML = 'Copied!';
        setTimeout(function () {
            document
                .querySelector('#tenant-domain>p').innerHTML = selectedCustomer.primaryDomainName
        }, 1000);
    }
}

// Check if the File Handler API is supported
if ('launchQueue' in window) {
    // Handle the file when the PWA is launched with a file
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
    addImportListener();
}

// Add event listener for customer dropdown
document
    .getElementById('customer-select')
    .addEventListener('input', function () {
        selectedCustomer = getSelectedCustomer(this.value);
        document.getElementById('tenant-id').disabled = false;
        document.getElementById('tenant-domain').disabled = false;
        tenantIdButton = document.querySelector('#tenant-id>p');
        tenantDomainButton = document.querySelector('#tenant-domain>p');
        if (selectedCustomer) {
            tenantIdButton.innerHTML = selectedCustomer.microsoftId;
            tenantDomainButton.innerHTML = selectedCustomer.primaryDomainName;
        } else {
            document.getElementById('tenant-id').disabled = true;
            tenantIdButton.innerHTML = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx";
            document.getElementById('tenant-domain').disabled = true;
            tenantDomainButton.innerHTML = "contoso.com";
        }
    });

// Define the mspButton custom element extending from HTMLButtonElement
class mspButton extends HTMLButtonElement {
    constructor() {
        super();
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
            } else {
                this.disabled = true;
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

        // Handle dropping element
        this.addEventListener('dragend', () => {
            updateUrl();
        });

        // Handle changes in the customer selection
        document.getElementById('customer-select').addEventListener('input', updateUrl);

        // Disable the element initially
        this.disabled = true;
    }
}

// Define the mspButton element
customElements.define('msp-button', mspButton, { extends: 'button' });

// Function to save the order of the MSP buttons
function saveOrder() {
    const containerIds = Array.from(containers).map(container => {
        return Array.from(container.children).map(draggable => draggable.id);
    });
    localStorage.setItem('savedOrder', JSON.stringify(containerIds));
}

// Add event listeners for dragging buttons
draggables.forEach(draggable => {
    draggable.addEventListener('dragstart', () => {
        draggable.classList.add('dragging')
    })

    draggable.addEventListener('dragend', () => {
        draggable.classList.remove('dragging');
        saveOrder(); // Save the order when drag ends
    })
})

// Add event listeners to allow for dropping
containers.forEach(container => {
    container.addEventListener('dragover', e => {
        e.preventDefault()
        const afterElement = getDragAfterElement(container, e.clientY)
        const draggable = document.querySelector('.dragging')
        if (afterElement == null) {
            container.appendChild(draggable)
        } else {
            container.insertBefore(draggable, afterElement)
        }
    })
})

function getDragAfterElement(container, y) {
    const draggableElements = [...container.querySelectorAll('.draggable:not(.dragging)')]

    return draggableElements.reduce((closest, child) => {
        const box = child.getBoundingClientRect()
        const offset = y - box.top - box.height / 2
        if (offset < 0 && offset > closest.offset) {
            return { offset: offset, element: child }
        } else {
            return closest
        }
    }, { offset: Number.NEGATIVE_INFINITY }).element
}