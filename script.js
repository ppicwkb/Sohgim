
        // Configuration
        const SHEET_ID = '1pQQoIcrQxM38AW01r9l6T8JZBtwmUVQBZH-eU3iWpaw';
        const API_KEY = 'AIzaSyBhiKDVDH4fle5_EqAIaA05YjpxVMEBYZM';
        const RANGE = 'A:N';

        // Global variables
        let allData = [];
        let filteredData = [];
        let headers = [];

        // DOM elements
        const loadingState = document.getElementById('loadingState');
        const dataContainer = document.getElementById('dataContainer');
        const cardContainer = document.getElementById('cardContainer');
        const tableHeader = document.getElementById('tableHeader');
        const tableBody = document.getElementById('tableBody');
        const smartSearch = document.getElementById('smartSearch');
        const tradedItemIdInput = document.getElementById('tradedItemIdInput');
        const tradedItemNameInput = document.getElementById('tradedItemNameInput');
        const batchInput = document.getElementById('batchInput');
        const containerIdInput = document.getElementById('containerIdInput');
        const activeFilters = document.getElementById('activeFilters');
        const clearFilters = document.getElementById('clearFilters');
        const resultsSummary = document.getElementById('resultsSummary');
        const resultsCount = document.getElementById('resultsCount');
        const noResults = document.getElementById('noResults');
        const errorMessage = document.getElementById('errorMessage');
        
        
        const exportPDF = document.getElementById('exportPDF');
        const totalQty = document.getElementById('totalQty');
        const baseQty = document.getElementById('baseQty');
        const tableViewBtn = document.getElementById('tableViewBtn');
        const cardViewBtn = document.getElementById('cardViewBtn');
        const cardModal = document.getElementById('cardModal');
        const closeModal = document.getElementById('closeModal');
        const modalContent = document.getElementById('modalContent');
        const lastModified = document.getElementById('lastModified');
        const activeRecords = document.getElementById('activeRecords');

        // Global variables
        let currentView = 'table';

        // Get last modified date from Drive API
        async function getLastModifiedDate() {
            try {
                const url = `https://www.googleapis.com/drive/v3/files/${SHEET_ID}?fields=modifiedTime&key=${API_KEY}`;
                const response = await fetch(url);
                
                if (response.ok) {
                    const data = await response.json();
                    const modifiedDate = new Date(data.modifiedTime);
                    lastModified.textContent = modifiedDate.toLocaleString();
                } else {
                    lastModified.textContent = 'Unable to fetch';
                }
            } catch (error) {
                console.error('Error fetching last modified date:', error);
                lastModified.textContent = 'Unable to fetch';
            }
        }

        // Load data from Google Sheets
        async function loadSheetData() {
            try {
                const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${RANGE}?key=${API_KEY}`;
                const response = await fetch(url);
                
                if (!response.ok) {
                    throw new Error('Failed to fetch data');
                }
                
                const data = await response.json();
                
                if (data.values && data.values.length > 0) {
                    headers = data.values[0];
                    allData = data.values.slice(1);
                    filteredData = [...allData];
                    
                    setupTable();
                    populateFilters();
                    displayData();
                    setupEventListeners();
                    getLastModifiedDate();
                    
                    loadingState.classList.add('hidden');
                    if (currentView === 'table') {
                        dataContainer.classList.remove('hidden');
                    } else {
                        cardContainer.classList.remove('hidden');
                    }
                    resultsSummary.classList.remove('hidden');
                } else {
                    throw new Error('No data found');
                }
                }
             catch (error) 
             {
                console.error('Error loading data:', error);
                loadingState.classList.add('hidden');
                errorMessage.classList.remove('hidden');
            }
        }

        // Setup table headers
        function setupTable() {
            tableHeader.innerHTML = '';
            headers.forEach(header => {
                const th = document.createElement('th');
                th.className = 'px-6 py-3 text-left text-xs font-medium uppercase tracking-wider';
                th.textContent = header;
                tableHeader.appendChild(th);
            });
        }

        // Populate filter suggestions
        function populateFilters() {
            // Store all unique values for suggestions
            window.filterData = {
                tradedItemId: [...new Set(allData.map(row => row[getColumnIndex('Traded Item ID')]).filter(val => val))].sort(),
                tradedItemName: [...new Set(allData.map(row => row[getColumnIndex('Traded Item Name')]).filter(val => val))].sort(),
                batch: [...new Set(allData.map(row => row[getColumnIndex('Batch')]).filter(val => val))].sort(),
                containerId: [...new Set(allData.map(row => row[getColumnIndex('Container Id')]).filter(val => val))].sort()
            };
        }

        // Get column index by header name
        function getColumnIndex(headerName) {
            return headers.findIndex(header => 
                header.toLowerCase().includes(headerName.toLowerCase())
            );
        }

        // Setup event listeners
        function setupEventListeners() {
            smartSearch.addEventListener('input', debounce(applyFilters, 300));
            
            // Setup suggestion inputs
            setupSuggestionInput(tradedItemIdInput, 'tradedItemIdSuggestions', 'tradedItemId');
            setupSuggestionInput(tradedItemNameInput, 'tradedItemNameSuggestions', 'tradedItemName');
            setupSuggestionInput(batchInput, 'batchSuggestions', 'batch');
            setupSuggestionInput(containerIdInput, 'containerIdSuggestions', 'containerId');
            
          
            // Setup dropdown buttons
            setupDropdownButton('tradedItemIdDropdown', 'tradedItemIdSuggestions', 'tradedItemId', tradedItemIdInput);
            setupDropdownButton('tradedItemNameDropdown', 'tradedItemNameSuggestions', 'tradedItemName', tradedItemNameInput);
            setupDropdownButton('batchDropdown', 'batchSuggestions', 'batch', batchInput);
            setupDropdownButton('containerIdDropdown', 'containerIdSuggestions', 'containerId', containerIdInput);
            
            // View toggle buttons
            tableViewBtn.addEventListener('click', () => switchView('table'));
            cardViewBtn.addEventListener('click', () => switchView('card'));
            
            // Modal events
            closeModal.addEventListener('click', () => cardModal.classList.add('hidden'));
            cardModal.addEventListener('click', (e) => {
                if (e.target === cardModal) {
                    cardModal.classList.add('hidden');
                }
            });
            
            clearFilters.addEventListener('click', clearAllFilters);
            
            
            exportPDF.addEventListener('click', exportToPDF);
            
            // Close suggestions when clicking outside
            document.addEventListener('click', (e) => {
                if (!e.target.closest('.relative')) {
                    hideAllSuggestions();
                }
            });
        }

        // Debounce function for search input
        function debounce(func, wait) {
            let timeout;
            return function executedFunction(...args) {
                const later = () => {
                    clearTimeout(timeout);
                    func(...args);
                };
                clearTimeout(timeout);
                timeout = setTimeout(later, wait);
            };
        }

        // Setup suggestion input functionality
        function setupSuggestionInput(inputElement, suggestionsId, dataKey) {
            const suggestionsElement = document.getElementById(suggestionsId);
            let selectedIndex = -1;

            inputElement.addEventListener('input', debounce(() => {
                const value = inputElement.value.toLowerCase();
                if (value.length === 0) {
                    suggestionsElement.classList.add('hidden');
                    return;
                }

                // Get available values based on current filters
                const availableValues = getAvailableValues(dataKey);
                const filteredSuggestions = availableValues.filter(item => 
                    item.toLowerCase().includes(value)
                ).slice(0, 10);

                if (filteredSuggestions.length > 0) {
                    showSuggestions(suggestionsElement, filteredSuggestions, inputElement);
                } else {
                    suggestionsElement.classList.add('hidden');
                }
                selectedIndex = -1;
            }, 200));

            inputElement.addEventListener('keydown', (e) => {
                const suggestions = suggestionsElement.querySelectorAll('.suggestion-item');
                
                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    selectedIndex = Math.min(selectedIndex + 1, suggestions.length - 1);
                    updateSelectedSuggestion(suggestions, selectedIndex);
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    selectedIndex = Math.max(selectedIndex - 1, -1);
                    updateSelectedSuggestion(suggestions, selectedIndex);
                } else if (e.key === 'Enter') {
                    e.preventDefault();
                    if (selectedIndex >= 0 && suggestions[selectedIndex]) {
                        inputElement.value = suggestions[selectedIndex].textContent;
                        suggestionsElement.classList.add('hidden');
                        applyFilters();
                    }
                } else if (e.key === 'Escape') {
                    suggestionsElement.classList.add('hidden');
                    selectedIndex = -1;
                }
            });

            inputElement.addEventListener('blur', () => {
                setTimeout(() => {
                    suggestionsElement.classList.add('hidden');
                }, 200);
            });

            inputElement.addEventListener('change', applyFilters);
        }

        // Show suggestions dropdown
        function showSuggestions(suggestionsElement, suggestions, inputElement) {
            suggestionsElement.innerHTML = '';
            suggestions.forEach(suggestion => {
                const div = document.createElement('div');
                div.className = 'suggestion-item';
                div.textContent = suggestion;
                div.addEventListener('click', () => {
                    inputElement.value = suggestion;
                    suggestionsElement.classList.add('hidden');
                    applyFilters();
                });
                suggestionsElement.appendChild(div);
            });
            suggestionsElement.classList.remove('hidden');
        }

        // Update selected suggestion highlight
        function updateSelectedSuggestion(suggestions, selectedIndex) {
            suggestions.forEach((suggestion, index) => {
                if (index === selectedIndex) {
                    suggestion.classList.add('selected');
                } else {
                    suggestion.classList.remove('selected');
                }
            });
        }

        // Hide all suggestion dropdowns
        function hideAllSuggestions() {
            document.querySelectorAll('[id$="Suggestions"]').forEach(el => {
                el.classList.add('hidden');
            });
        }

        // Setup dropdown button functionality
        function setupDropdownButton(buttonId, suggestionsId, dataKey, inputElement) {
            const button = document.getElementById(buttonId);
            const suggestionsElement = document.getElementById(suggestionsId);
            
            button.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                
                // Hide other dropdowns first
                document.querySelectorAll('[id$="Suggestions"]').forEach(el => {
                    if (el.id !== suggestionsId) {
                        el.classList.add('hidden');
                    }
                });
                
                // Toggle current dropdown
                if (suggestionsElement.classList.contains('hidden')) {
                    const availableValues = getAvailableValues(dataKey);
                    if (availableValues.length > 0) {
                        showAllSuggestions(suggestionsElement, availableValues, inputElement);
                    }
                } else {
                    suggestionsElement.classList.add('hidden');
                }
            });
        }

        // Show all available suggestions (for dropdown button)
        function showAllSuggestions(suggestionsElement, suggestions, inputElement) {
            suggestionsElement.innerHTML = '';
            suggestions.slice(0, 20).forEach(suggestion => { // Limit to 20 items for performance
                const div = document.createElement('div');
                div.className = 'suggestion-item';
                div.textContent = suggestion;
                div.addEventListener('click', () => {
                    inputElement.value = suggestion;
                    suggestionsElement.classList.add('hidden');
                    applyFilters();
                });
                suggestionsElement.appendChild(div);
            });
            suggestionsElement.classList.remove('hidden');
        }

        // Get available values based on current filters
        function getAvailableValues(dataKey) {
            if (!window.filterData) return [];
            
            // Get currently filtered data to show only relevant suggestions
            let currentFiltered = [...allData];
            
            // Apply other filters to get interconnected suggestions
            const otherFilters = {
                tradedItemId: tradedItemIdInput.value,
                tradedItemName: tradedItemNameInput.value,
                batch: batchInput.value,
                containerId: containerIdInput.value
            };
            
            // Remove current filter from consideration
            delete otherFilters[dataKey];
            
            Object.entries(otherFilters).forEach(([key, value]) => {
                if (value) {
                    const columnIndex = getColumnIndex(key === 'tradedItemId' ? 'Traded Item ID' : 
                                                     key === 'tradedItemName' ? 'Traded Item Name' :
                                                     key === 'batch' ? 'Batch' : 'Container Id');
                    if (columnIndex !== -1) {
                        currentFiltered = currentFiltered.filter(row => row[columnIndex] === value);
                    }
                }
            });
            
            // Get unique values from filtered data
            const columnIndex = getColumnIndex(dataKey === 'tradedItemId' ? 'Traded Item ID' : 
                                             dataKey === 'tradedItemName' ? 'Traded Item Name' :
                                             dataKey === 'batch' ? 'Batch' : 'Container Id');
            
            if (columnIndex !== -1) {
                return [...new Set(currentFiltered.map(row => row[columnIndex]).filter(val => val))].sort();
            }
            
            return window.filterData[dataKey] || [];
        }

        // Apply all filters
        function applyFilters() {
            let filtered = [...allData];

            // Smart search filter
            const searchTerm = smartSearch.value.toLowerCase().trim();
            if (searchTerm) {
                filtered = filtered.filter(row => 
                    row.some(cell => 
                        cell && cell.toString().toLowerCase().includes(searchTerm)
                    )
                );
            }

            // Column-specific filters
            const filters = [
                { element: tradedItemIdInput, columnIndex: getColumnIndex('Traded Item ID') },
                { element: tradedItemNameInput, columnIndex: getColumnIndex('Traded Item Name') },
                { element: batchInput, columnIndex: getColumnIndex('Batch') },
                { element: containerIdInput, columnIndex: getColumnIndex('Container Id') }
            ];

            filters.forEach(filter => {
                if (filter.element.value && filter.columnIndex !== -1) {
                    filtered = filtered.filter(row => 
                        row[filter.columnIndex] === filter.element.value
                    );
                }
            });

            filteredData = filtered;
            displayData();
            updateActiveFilters();
            updateQuantityCards();
        }

        // Update quantity cards
        function updateQuantityCards() {
            const qtyIndex = getColumnIndex('Qty');
            const baseQtyIndex = getColumnIndex('Base Qty');
            
            let totalSum = 0;
            let baseSum = 0;
            
            filteredData.forEach(row => {
                if (qtyIndex !== -1 && row[qtyIndex]) {
                    const value = parseFloat(row[qtyIndex]) || 0;
                    totalSum += value;
                }
                if (baseQtyIndex !== -1 && row[baseQtyIndex]) {
                    const value = parseFloat(row[baseQtyIndex]) || 0;
                    baseSum += value;
                }
            });
            
            totalQty.textContent = totalSum.toLocaleString();
            baseQty.textContent = baseSum.toLocaleString();
            
            
        }

        // Switch between table and card view
        function switchView(view) {
            currentView = view;
            
            if (view === 'table') {
                tableViewBtn.className = 'premium-btn px-6 py-3 rounded-xl text-sm font-semibold transition-all duration-300 flex items-center space-x-2';
                cardViewBtn.className = 'px-6 py-3 rounded-xl text-sm font-semibold transition-all duration-300 text-gray-600 hover:text-gray-800 hover:bg-white hover:shadow-md flex items-center space-x-2';
                dataContainer.classList.remove('hidden');
                cardContainer.classList.add('hidden');
            } else {
                cardViewBtn.className = 'premium-btn px-6 py-3 rounded-xl text-sm font-semibold transition-all duration-300 flex items-center space-x-2';
                tableViewBtn.className = 'px-6 py-3 rounded-xl text-sm font-semibold transition-all duration-300 text-gray-600 hover:text-gray-800 hover:bg-white hover:shadow-md flex items-center space-x-2';
                dataContainer.classList.add('hidden');
                cardContainer.classList.remove('hidden');
            }
            
            displayData();
        }

        // Display filtered data
        function displayData() {
            if (filteredData.length === 0) {
                dataContainer.classList.add('hidden');
                cardContainer.classList.add('hidden');
                noResults.classList.remove('hidden');
                resultsSummary.classList.add('hidden');
                return;
            }

            noResults.classList.add('hidden');
            resultsSummary.classList.remove('hidden');

            if (currentView === 'table') {
                displayTableData();
                dataContainer.classList.remove('hidden');
                cardContainer.classList.add('hidden');
            } else {
                displayCardData();
                cardContainer.classList.remove('hidden');
                dataContainer.classList.add('hidden');
            }

            resultsCount.textContent = `Showing ${filteredData.length} of ${allData.length} records`;
        }

        // Display data in table format
        function displayTableData() {
            tableBody.innerHTML = '';
            const searchTerm = smartSearch.value.toLowerCase().trim();

            filteredData.forEach((row, index) => {
                const tr = document.createElement('tr');
                
                row.forEach(cell => {
                    const td = document.createElement('td');
                    
                    if (searchTerm && cell && cell.toString().toLowerCase().includes(searchTerm)) {
                        td.innerHTML = highlightSearchTerm(cell.toString(), searchTerm);
                    } else {
                        td.textContent = cell || '';
                    }
                    
                    tr.appendChild(td);
                });
                
                tableBody.appendChild(tr);
            });
        }

        // Display data in card format
        function displayCardData() {
            cardContainer.innerHTML = '';
            const searchTerm = smartSearch.value.toLowerCase().trim();

            filteredData.forEach((row, index) => {
                const card = document.createElement('div');
                card.className = 'card cursor-pointer hover:scale-105 transition-all duration-300 border border-white border-opacity-20';
                
                // Get key information for card display
                const tradedItemName = row[getColumnIndex('Traded Item Name')] || ' ';
                const tradedItemId = row[getColumnIndex('Traded Item ID')] || ' ';
                const qty = row[getColumnIndex('Qty')] || '0';
                const baseQty = row[getColumnIndex('Base Qty')] || '0';
                const batch = row[getColumnIndex('Batch')] || ' ';
                const containerId = row[getColumnIndex('Container Id')] || ' ';
                
                card.innerHTML = `
                    <div class="p-4">
                        <div class="flex justify-between items-start mb-4">
                            <div class="flex-1">
                                <h3 class="text-sm font-bold text-gray-900 mb-2 clamp-2 leading-tight">${highlightSearchTerm(tradedItemName, searchTerm)}</h3>
                                <div class="inline-flex items-center px-2 py-1 rounded-full bg-blue-50 border border-blue-200">
                                    <span class="text-xs text-blue-700 font-semibold">ID: ${tradedItemId}</span>
                                </div>
                            </div>
                            <div class="flex flex-col items-end space-y-1">
                             
                                
                            </div>
                        </div>
                        
                        <div class="grid grid-cols-2 gap-3 mb-4">
                            <div class="text-center p-3 bg-gradient-to-br from-blue-50 to-blue-100 rounded-xl border border-blue-200">
                                <p class="text-xs text-blue-600 mb-1 font-bold uppercase tracking-wider">Quantity</p>
                                <p class="text-lg font-bold text-blue-700">${qty}</p>
                            </div>
                            <div class="text-center p-3 bg-gradient-to-br from-green-50 to-green-100 rounded-xl border border-green-200">
                                <p class="text-xs text-green-600 mb-1 font-bold uppercase tracking-wider">KG</p>
                                <p class="text-lg font-bold text-green-700">${baseQty}</p>
                            </div>
                        </div>
                        
                        <div class="space-y-2 mb-4">
                            <div class="flex justify-between items-center p-2 bg-gray-50 rounded-lg border border-gray-200">
                                <span class="text-xs text-gray-600 font-medium">Batch:</span>
                                <span class="text-xs font-bold text-gray-900">${batch}</span>
                            </div>
                            <div class="flex justify-between items-center p-2 bg-gray-50 rounded-lg border border-gray-200">
                                <span class="text-xs text-gray-600 font-medium">Container:</span>
                                <span class="text-xs font-bold text-gray-900">${containerId}</span>
                            </div>
                        </div>
                        
                        <button class="w-full btn bg-gradient-to-r from-indigo-500 to-purple-600 hover:from-indigo-600 hover:to-purple-700 py-2 px-3 rounded-lg text-xs font-semibold hover:scale-105 transition-all duration-300 shadow-lg hover:shadow-indigo-500/25">
                            <span class="flex items-center justify-center space-x-1">
                                <svg class="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/>
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"/>
                                </svg>
                                <span>View Details</span>
                            </span>
                        </button>
                    </div>
                `;
                
                card.addEventListener('click', () => showCardModal(row, index));
                cardContainer.appendChild(card);
            });
        }

        // Show card modal with full details
        function showCardModal(row, index) {
            modalContent.innerHTML = '';
            
            // Create a clean grid layout
            const grid = document.createElement('div');
            grid.className = 'grid grid-cols-1 md:grid-cols-2 gap-4';
            
            headers.forEach((header, headerIndex) => {
                const value = row[headerIndex] || '';
                const div = document.createElement('div');
                div.className = 'bg-white rounded-xl p-4 border border-gray-200 hover:border-gray-300 transition-all duration-200 hover:shadow-md';
                
                // Simple icon selection
                let iconColor = 'bg-gray-500';
                let textColor = 'text-gray-700';
                
                if (header.toLowerCase().includes('qty')) {
                    iconColor = 'bg-blue-500';
                    textColor = 'text-blue-700';
                } else if (header.toLowerCase().includes('id')) {
                    iconColor = 'bg-purple-500';
                    textColor = 'text-purple-700';
                } else if (header.toLowerCase().includes('batch')) {
                    iconColor = 'bg-green-500';
                    textColor = 'text-green-700';
                } else if (header.toLowerCase().includes('container')) {
                    iconColor = 'bg-orange-500';
                    textColor = 'text-orange-700';
                } else if (header.toLowerCase().includes('name')) {
                    iconColor = 'bg-pink-500';
                    textColor = 'text-pink-700';
                }
                
                div.innerHTML = `
                    <div class="flex items-start space-x-3">
                        <div class="w-8 h-8 ${iconColor} rounded-lg flex items-center justify-center flex-shrink-0 mt-1">
                            <div class="w-2 h-2 bg-white rounded-full"></div>
                        </div>
                        <div class="flex-1 min-w-0">
                            <p class="text-xs font-semibold ${textColor} uppercase tracking-wider mb-1">${header}</p>
                            <p class="text-sm font-medium text-gray-900 break-words">${value || 'N/A'}</p>
                        </div>
                    </div>
                `;
                grid.appendChild(div);
            });
            
            modalContent.appendChild(grid);
            cardModal.classList.remove('hidden');
        }

        // Highlight search terms
        function highlightSearchTerm(text, term) {
            if (!term) return text;
            const regex = new RegExp(`(${term})`, 'gi');
            return text.replace(regex, '<span class="highlight">$1</span>');
        }

        // Update active filters display
        function updateActiveFilters() {
            activeFilters.innerHTML = '';
            
            const filters = [
                { label: 'Search', value: smartSearch.value },
                { label: 'Item ID', value: tradedItemIdInput.value },
                { label: 'Item Name', value: tradedItemNameInput.value },
                { label: 'Batch', value: batchInput.value },
                { label: 'Container', value: containerIdInput.value }
            ];

            filters.forEach(filter => {
                if (filter.value) {
                    const badge = document.createElement('span');
                    badge.className = 'badge px-3 py-1 rounded-full text-sm flex items-center gap-2';
                    badge.innerHTML = `
                        ${filter.label}: ${filter.value}
                        <button onclick="removeFilter('${filter.label}')" class="text-white hover:text-gray-200">×</button>
                    `;
                    activeFilters.appendChild(badge);
                }
            });
        }

        // Remove specific filter
        function removeFilter(filterLabel) {
            switch(filterLabel) {
                case 'Search':
                    smartSearch.value = '';
                    break;
                case 'Item ID':
                    tradedItemIdInput.value = '';
                    break;
                case 'Item Name':
                    tradedItemNameInput.value = '';
                    break;
                case 'Batch':
                    batchInput.value = '';
                    break;
                case 'Container':
                    containerIdInput.value = '';
                    break;
            }
            applyFilters();
        }

        // Clear all filters
        function clearAllFilters() {
            smartSearch.value = '';
            tradedItemIdInput.value = '';
            tradedItemNameInput.value = '';
            batchInput.value = '';
            containerIdInput.value = '';
            hideAllSuggestions();
            applyFilters();
        }

        // Export to CSV
        
        

        // Export to PDF
        function exportToPDF() {
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF('l', 'mm', 'a4'); // landscape orientation
            
            // Add title
            doc.setFontSize(16);
            doc.text('Data SOH GIM by PPIC', 14, 15);
            
            // Add export date
            doc.setFontSize(10);
            doc.text(`Export Date: ${new Date().toLocaleDateString()}`, 14, 25);
            
            // Add summary information
            const qtyIndex = getColumnIndex('Qty');
            const baseQtyIndex = getColumnIndex('Base Qty');
            
            let totalSum = 0;
            let baseSum = 0;
            
            filteredData.forEach(row => {
                if (qtyIndex !== -1 && row[qtyIndex]) {
                    totalSum += parseFloat(row[qtyIndex]) || 0;
                }
                if (baseQtyIndex !== -1 && row[baseQtyIndex]) {
                    baseSum += parseFloat(row[baseQtyIndex]) || 0;
                }
            });
            
            doc.text(`Total Records: ${filteredData.length}`, 14, 35);
            doc.text(`Total Qty: ${totalSum.toLocaleString()}`, 14, 42);
            doc.text(`Total Base Qty: ${baseSum.toLocaleString()}`, 14, 49);
            
            // Prepare table data
            const tableData = filteredData.map(row => 
                row.map(cell => cell || '')
            );
            
            // Add table
            doc.autoTable({
                head: [headers],
                body: tableData,
                startY: 55,
                styles: {
                    fontSize: 8,
                    cellPadding: 2
                },
                headStyles: {
                    fillColor: [59, 130, 246],
                    textColor: 255,
                    fontStyle: 'bold'
                },
                alternateRowStyles: {
                    fillColor: [249, 250, 251]
                },
                margin: { top: 55, left: 14, right: 14 },
                tableWidth: 'auto',
                columnStyles: {},
                didDrawPage: function (data) {
                    // Add page numbers
                    doc.setFontSize(8);
                    doc.text(`Page ${data.pageNumber}`, doc.internal.pageSize.width - 20, doc.internal.pageSize.height - 10);
                }
            });
            
            // Save the PDF
            doc.save(`SOH_data_by PPIC_${new Date().toISOString().split('T')[0]}.pdf`);
        }

    
     

        // Initialize the application
        loadSheetData();
       
