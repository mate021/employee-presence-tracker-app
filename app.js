// DOM elements
const employeesInput = document.getElementById('employees');
const cardLoginInput = document.getElementById('card-login');
const officialLeavesInput = document.getElementById('official-leaves');
const calculateBtn = document.getElementById('calculate-btn');
const employeesStatus = document.getElementById('employees-status').querySelector('span');
const cardLoginStatus = document.getElementById('card-login-status').querySelector('span');
const officialLeavesStatus = document.getElementById('official-leaves-status').querySelector('span');
const resultsBody = document.getElementById('results-body');
const sortableHeaders = document.querySelectorAll('.sortable');
const directorateKpis = document.getElementById('directorate-kpis');

// Track upload status
let employeesUploaded = false;
let cardLoginUploaded = false;
let officialLeavesUploaded = false;

// Store results data for sorting
let resultsData = [];
let currentSortColumn = null;
let currentSortDirection = 'asc';

// Calculate button event listener
calculateBtn.addEventListener('click', async function() {
    // Check if files are selected
    if (employeesInput.files.length === 0) {
        alert('Please select an employees file');
        return;
    }
    if (cardLoginInput.files.length === 0) {
        alert('Please select a card login file');
        return;
    }
    if (officialLeavesInput.files.length === 0) {
        alert('Please select an official leaves file');
        return;
    }
    
    // Upload all files first
    try {
        await uploadFile(employeesInput.files[0], 'employees');
        await uploadFile(cardLoginInput.files[0], 'card-login');
        await uploadFile(officialLeavesInput.files[0], 'official-leaves');
        
        // Then calculate presence
        calculatePresence();
    } catch (error) {
        console.error('Error during upload process:', error);
        alert('An error occurred during the upload process. Please try again.');
    }
});

// Function to upload a file to the server
function uploadFile(file, type) {
    return new Promise((resolve, reject) => {
        const formData = new FormData();
        formData.append('file', file);
        
        // Show loading indicator
        const statusElement = getStatusElement(type);
        statusElement.textContent = 'Uploading...';
        
        fetch(`/api/upload/${type}`, {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.json();
        })
        .then(data => {
            console.log(`${type} upload successful:`, data);
            statusElement.textContent = 'Uploaded';
            statusElement.classList.add('uploaded');
            
            // Update upload status
            updateUploadStatus(type, true);
            resolve(data);
        })
        .catch(error => {
            console.error(`Error uploading ${type} file:`, error);
            statusElement.textContent = 'Upload failed';
            statusElement.classList.remove('uploaded');
            
            // Update upload status
            updateUploadStatus(type, false);
            reject(error);
        });
    });
}

// Helper function to get the status element for a file type
function getStatusElement(type) {
    if (type === 'employees') {
        return employeesStatus;
    } else if (type === 'card-login') {
        return cardLoginStatus;
    } else if (type === 'official-leaves') {
        return officialLeavesStatus;
    }
    return null;
}

// Helper function to update upload status
function updateUploadStatus(type, status) {
    if (type === 'employees') {
        employeesUploaded = status;
    } else if (type === 'card-login') {
        cardLoginUploaded = status;
    } else if (type === 'official-leaves') {
        officialLeavesUploaded = status;
    }
}

// Function to calculate employee presence
function calculatePresence() {
    // Show loading indicator
    resultsBody.innerHTML = '<tr><td colspan="7">Loading results...</td></tr>';
    
    // Fetch calculation results from the server
    fetch('/api/calculate')
    .then(response => {
        if (!response.ok) {
            throw new Error('Network response was not ok');
        }
        return response.json();
    })
    .then(data => {
        // Display results
        displayResults(data.employees);
    })
    .catch(error => {
        console.error('Error calculating presence:', error);
        resultsBody.innerHTML = '<tr><td colspan="7">Error calculating results. Please try again.</td></tr>';
        alert('Failed to calculate presence. Please try again.');
    });
}

// Function to display results in the table
function displayResults(results) {
    // Store results for sorting
    resultsData = results;
    
    // Clear previous results
    resultsBody.innerHTML = '';
    
    if (results.length === 0) {
        resultsBody.innerHTML = '<tr><td colspan="7">No results found</td></tr>';
        return;
    }
    
    // Calculate and display KPIs
    calculateDirectorateKPIs(results);
    
    // Display results
    renderResultsTable(resultsData);
    
    // Add event listeners to sortable headers
    setupSortListeners();
}

// Function to calculate and display directorate KPIs
function calculateDirectorateKPIs(results) {
    // Clear previous KPIs
    directorateKpis.innerHTML = '';
    
    // Group employees by directorate
    const directorates = {};
    
    results.forEach(employee => {
        const directorate = employee.directorate || 'Unknown';
        
        if (!directorates[directorate]) {
            directorates[directorate] = {
                totalUsage: 0,
                validEmployees: 0
            };
        }
        
        // Only include employees with valid building usage (not 'N/A' or '0%')
        const buildingUsage = employee.buildingUsage;
        if (buildingUsage !== 'N/A' && buildingUsage !== '0%' && buildingUsage !== '0.0%') {
            // Extract numeric value from percentage string
            const usageValue = parseFloat(buildingUsage);
            if (!isNaN(usageValue)) {
                directorates[directorate].totalUsage += usageValue;
                directorates[directorate].validEmployees++;
            }
        }
    });
    
    // Calculate average and create KPI elements
    Object.keys(directorates).sort().forEach(directorate => {
        const { totalUsage, validEmployees } = directorates[directorate];
        
        // Calculate average (avoid division by zero)
        let averageUsage = 0;
        if (validEmployees > 0) {
            averageUsage = totalUsage / validEmployees;
        }
        
        // Determine usage class based on percentage
        let usageClass = '';
        if (averageUsage >= 70) {
            usageClass = 'high-usage';
        } else if (averageUsage >= 50) {
            usageClass = 'medium-usage';
        } else {
            usageClass = 'low-usage';
        }
        
        // Create KPI element
        const kpiElement = document.createElement('div');
        kpiElement.className = `directorate-kpi ${usageClass}`;
        kpiElement.innerHTML = `
            <div class="directorate-name">${directorate}</div>
            <div class="usage-value">${averageUsage.toFixed(1)}%</div>
        `;
        
        directorateKpis.appendChild(kpiElement);
    });
}

// Function to render the results table
function renderResultsTable(data) {
    resultsBody.innerHTML = '';
    
    data.forEach(result => {
        const row = document.createElement('tr');
        
        row.innerHTML = `
            <td>${result.name || ''}</td>
            <td>${result.directorate || ''}</td>
            <td>${result.department || ''}</td>
            <td>${result.loginCount || 0}</td>
            <td>${result.leaveCount || 0}</td>
            <td>${result.daysHome || 0}</td>
            <td>${result.buildingUsage || ''}</td>
        `;
        
        resultsBody.appendChild(row);
    });
}

// Function to set up sort listeners
function setupSortListeners() {
    sortableHeaders.forEach(header => {
        header.addEventListener('click', () => {
            const column = header.dataset.column;
            sortTable(column);
        });
    });
}

// Function to sort the table
function sortTable(column) {
    // Toggle sort direction if clicking the same column
    if (currentSortColumn === column) {
        currentSortDirection = currentSortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        currentSortColumn = column;
        currentSortDirection = 'asc';
    }
    
    // Sort the data
    resultsData.sort((a, b) => {
        let valueA = a[column];
        let valueB = b[column];
        
        // Handle numeric values
        if (column === 'loginCount' || column === 'leaveCount' || column === 'daysHome') {
            valueA = Number(valueA) || 0;
            valueB = Number(valueB) || 0;
        }
        // Handle building usage percentage
        else if (column === 'buildingUsage') {
            valueA = parseFloat(valueA) || 0;
            valueB = parseFloat(valueB) || 0;
        }
        // Handle text values
        else {
            valueA = (valueA || '').toLowerCase();
            valueB = (valueB || '').toLowerCase();
        }
        
        if (valueA < valueB) {
            return currentSortDirection === 'asc' ? -1 : 1;
        }
        if (valueA > valueB) {
            return currentSortDirection === 'asc' ? 1 : -1;
        }
        return 0;
    });
    
    // Update the table display
    renderResultsTable(resultsData);
    
    // Update sort indicators
    updateSortIndicators(column);
}

// Function to update sort indicators in the table headers
function updateSortIndicators(column) {
    sortableHeaders.forEach(header => {
        // Remove existing indicators
        header.classList.remove('sort-asc', 'sort-desc');
        
        // Add indicator to current sort column
        if (header.dataset.column === column) {
            header.classList.add(`sort-${currentSortDirection}`);
        }
    });
}