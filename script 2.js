
   
        // Global variables
        let allData = [];
        let filteredData = [];
        let currentPage = 1;
        const itemsPerPage = 50;
        let sortColumn = -1;
        let sortDirection = 'asc';

        // API Configuration
        const API_KEY = 'AIzaSyCfnSym_wpPtzWL0P3-kslY1Y14B7nD-34';
        const SPREADSHEET_ID = '1pQQoIcrQxM38AW01r9l6T8JZBtwmUVQBZH-eU3iWpaw';
        const RANGE = 'A:N'; // Will try multiple sheet names

        // Initialize application
        document.addEventListener('DOMContentLoaded', function() {
            loadData();
            setupEventListeners();
            startAutoRefresh(); // Start real-time monitoring
        });

        function setupEventListeners() {
            // Search functionality
            document.getElementById('searchInput').addEventListener('input', function() {
                filterData();
            });

            // Sortable headers
            document.querySelectorAll('.sortable').forEach(header => {
                header.addEventListener('click', function() {
                    const column = parseInt(this.dataset.column);
                    sortData(column);
                });
            });
        }

        async function loadData() {
            try {
                document.getElementById('loadingIndicator').style.display = 'block';
                console.log('Loading data from Google Sheets...');
                
                // Get spreadsheet metadata to check last modified time
                let lastModified = null;
                let spreadsheetTitle = 'Unknown';
                try {
                    const metadataUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}?key=${API_KEY}&fields=properties.title,properties.timeZone,sheets.properties.title`;
                    const metadataResponse = await fetch(metadataUrl);
                    if (metadataResponse.ok) {
                        const metadata = await metadataResponse.json();
                        spreadsheetTitle = metadata.properties?.title || 'Unknown';
                        console.log('Spreadsheet metadata:', metadata);
                        

                    }
                } catch (e) {
                    console.log('Could not fetch metadata:', e.message);
                }
                
                // Get file metadata from Drive API for last modified time
                try {
                    const driveUrl = `https://www.googleapis.com/drive/v3/files/${SPREADSHEET_ID}?key=${API_KEY}&fields=modifiedTime,name,owners`;
                    const driveResponse = await fetch(driveUrl);
                    if (driveResponse.ok) {
                        const driveData = await driveResponse.json();
                        if (driveData.modifiedTime) {
                            lastModified = new Date(driveData.modifiedTime);
                            console.log('Last modified:', lastModified);
                            
                            // Update last modified display
                            const lastModifiedElement = document.getElementById('lastUpdated');
                            if (lastModifiedElement) {
                                const timeAgo = getTimeAgo(lastModified);
                                lastModifiedElement.innerHTML = `
                                    <span class="font-bold">${lastModified.toLocaleString()}</span>
                                    <span class="text-xs opacity-75 ml-2">(${timeAgo})</span>
                                `;
                            }
                            
                            // Update mobile last updated display
                            const mobileLastUpdated = document.getElementById('mobileLastUpdated');
                            if (mobileLastUpdated) {
                                const timeAgo = getTimeAgo(lastModified);
                                mobileLastUpdated.textContent = `${lastModified.toLocaleString()} (${timeAgo})`;
                            }
                        }
                        
                        // Update system status based on how recent the data is
                        if (lastModified) {
                            const minutesAgo = (new Date() - lastModified) / (1000 * 60);
                            const systemStatus = document.getElementById('systemStatus');
                            const statusIndicator = systemStatus.parentElement.querySelector('.animate-pulse');
                            
                            // Mobile status elements
                            const mobileStatusDot = document.getElementById('mobileStatusDot');
                            const mobileStatusText = document.getElementById('mobileStatusText');
                            
                            if (minutesAgo < 5) {
                                // Desktop status
                                systemStatus.textContent = 'LIVE';
                                systemStatus.className = 'text-2xl font-bold text-green-300 drop-shadow-lg';
                                statusIndicator.className = 'w-3 h-3 bg-green-400 rounded-full animate-pulse shadow-lg';
                                
                                // Mobile status
                                if (mobileStatusDot && mobileStatusText) {
                                    mobileStatusDot.className = 'realtime-status-dot live';
                                    mobileStatusText.textContent = 'LIVE';
                                }
                            } else if (minutesAgo < 30) {
                                // Desktop status
                                systemStatus.textContent = 'RECENT';
                                systemStatus.className = 'text-2xl font-bold text-yellow-300 drop-shadow-lg';
                                statusIndicator.className = 'w-3 h-3 bg-yellow-400 rounded-full animate-pulse shadow-lg';
                                
                                // Mobile status
                                if (mobileStatusDot && mobileStatusText) {
                                    mobileStatusDot.className = 'realtime-status-dot recent';
                                    mobileStatusText.textContent = 'RECENT';
                                }
                            } else {
                                // Desktop status
                                systemStatus.textContent = 'STALE';
                                systemStatus.className = 'text-2xl font-bold text-orange-300 drop-shadow-lg';
                                statusIndicator.className = 'w-3 h-3 bg-orange-400 rounded-full animate-pulse shadow-lg';
                                
                                // Mobile status
                                if (mobileStatusDot && mobileStatusText) {
                                    mobileStatusDot.className = 'realtime-status-dot stale';
                                    mobileStatusText.textContent = 'STALE';
                                }
                            }
                        }
                    }
                } catch (e) {
                    console.log('Could not fetch Drive metadata:', e.message);
                }
                
                // Try different sheet names and ranges
                const possibleRanges = [
                    'A:N',
                    'Sheet1!A:N', 
                    'Lembar1!A:N',
                    'Data!A:N',
                    'Inventory!A:N'
                ];
                
                let data = null;
                let successUrl = '';
                
                for (const range of possibleRanges) {
                    try {
                        const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${range}?key=${API_KEY}`;
                        console.log('Trying URL:', url);
                        
                        const response = await fetch(url);
                        
                        if (response.ok) {
                            const testData = await response.json();
                            if (testData.values && testData.values.length > 0) {
                                data = testData;
                                successUrl = url;
                                console.log(`Success with range: ${range}`);
                                break;
                            }
                        }
                    } catch (e) {
                        console.log(`Failed with range ${range}:`, e.message);
                        continue;
                    }
                }
                
                if (!data) {
                    throw new Error('Could not access spreadsheet with any sheet name. Please ensure the spreadsheet is publicly accessible (Anyone with link can view)');
                }
                
                console.log('Data received:', data);
                
                if (data.values && data.values.length > 0) {
                    console.log('Total rows:', data.values.length);
                    
                    // Check if first row looks like headers
                    const hasHeaders = data.values[0].some(cell => 
                        typeof cell === 'string' && 
                        (cell.toLowerCase().includes('location') || 
                         cell.toLowerCase().includes('item') || 
                         cell.toLowerCase().includes('qty'))
                    );
                    
                    // Skip header row if it exists
                    const dataRows = hasHeaders ? data.values.slice(1) : data.values;
                    
                    allData = dataRows.map(row => {
                        // Ensure all columns exist
                        while (row.length < 14) {
                            row.push('');
                        }
                        return row;
                    });
                    
                    console.log('Processed data rows:', allData.length);
                    
                    filteredData = [...allData];
                    updateDashboard();
                    renderTable();
                    updatePagination();
                    updateTimestamp();
                    
                    // Show success message
                    if (allData.length > 0) {
                        console.log('Data loaded successfully!');
                    } else {
                        throw new Error('No data rows found after processing');
                    }
                } else {
                    throw new Error('No data found in spreadsheet or sheet is empty');
                }
                
            } catch (error) {
                console.error('Error loading data:', error);
                document.getElementById('tableBody').innerHTML = `
                    <tr>
                        <td colspan="14" class="px-4 py-8 text-center text-red-600">
                            ❌ Error memuat data: ${error.message}<br>
                            <small class="text-gray-500 mt-2 block">
                                Pastikan spreadsheet dapat diakses publik dan API key valid.<br>
                                Cek console browser untuk detail error.
                            </small>
                        </td>
                    </tr>
                `;
            } finally {
                document.getElementById('loadingIndicator').style.display = 'none';
            }
        }

        function updateDashboard() {
            const totalQty = filteredData.reduce((sum, row) => sum + (parseFloat(row[10]) || 0), 0);
            const totalBaseQty = filteredData.reduce((sum, row) => sum + (parseFloat(row[12]) || 0), 0);
            const uniqueItems = new Set(filteredData.map(row => row[1])).size;
            const lowQtyItems = filteredData.filter(row => (parseFloat(row[10]) || 0) < 10).length;

            // Update desktop dashboard
            document.getElementById('totalQty').textContent = totalQty.toLocaleString();
            document.getElementById('totalBaseQty').textContent = totalBaseQty.toLocaleString();
            document.getElementById('uniqueItems').textContent = uniqueItems.toLocaleString();
            document.getElementById('lowQtyItems').textContent = lowQtyItems.toLocaleString();

            // Update mobile dashboard
            document.getElementById('mobileTotalQty').textContent = totalQty.toLocaleString();
            document.getElementById('mobileTotalBaseQty').textContent = totalBaseQty.toLocaleString();
            document.getElementById('mobileUniqueItems').textContent = uniqueItems.toLocaleString();
            document.getElementById('mobileLowQtyItems').textContent = lowQtyItems.toLocaleString();
            
            // Update top items analysis
            updateTopItemsAnalysis();
        }

        function updateTopItemsAnalysis() {
            // Group items by name and sum quantities
            const itemGroups = {};
            
            filteredData.forEach(row => {
                const itemName = row[2] || 'Unknown Item';
                const itemId = row[1] || 'Unknown ID';
                const qty = parseFloat(row[10]) || 0;
                const baseQty = parseFloat(row[12]) || 0;
                const uom = row[11] || '';
                const baseUom = row[13] || '';
                
                const key = `${itemName}_${itemId}`;
                
                if (!itemGroups[key]) {
                    itemGroups[key] = {
                        name: itemName,
                        id: itemId,
                        totalQty: 0,
                        totalBaseQty: 0,
                        uom: uom,
                        baseUom: baseUom,
                        locations: 0
                    };
                }
                
                itemGroups[key].totalQty += qty;
                itemGroups[key].totalBaseQty += baseQty;
                itemGroups[key].locations += 1;
            });
            
            // Convert to array and sort
            const itemsArray = Object.values(itemGroups);
            
            // Top 10 by quantity (minimized)
            const topByQty = itemsArray
                .sort((a, b) => b.totalQty - a.totalQty)
                .slice(0, 10);
            
            // Top 10 by base quantity (minimized)
            const topByBaseQty = itemsArray
                .sort((a, b) => b.totalBaseQty - a.totalBaseQty)
                .slice(0, 10);
            
            // Render compact top items by quantity
            const topQtyContainer = document.getElementById('topItemsByQty');
            if (topQtyContainer) {
                topQtyContainer.innerHTML = topByQty.map((item, index) => `
                    <div class="flex items-center justify-between p-2 bg-white rounded border border-blue-100 hover:border-blue-200 transition-colors">
                        <div class="flex items-center space-x-2">
                            <div class="w-5 h-5 bg-blue-500 text-white rounded-full flex items-center justify-center text-xs font-bold">
                                ${index + 1}
                            </div>
                            <div class="flex-1 min-w-0">
                                <div class="text-xs font-semibold text-blue-900 truncate" title="${item.name}">
                                    ${item.name.length > 15 ? item.name.substring(0, 15) + '...' : item.name}
                                </div>
                                <div class="text-xs text-blue-600 font-mono opacity-75">
                                    ${item.id} • ${item.locations}x
                                </div>
                            </div>
                        </div>
                        <div class="text-right">
                            <div class="text-sm font-bold text-blue-900">
                                ${item.totalQty.toLocaleString()}
                            </div>
                            <div class="text-xs text-blue-600 opacity-75">
                                ${item.uom}
                            </div>
                        </div>
                    </div>
                `).join('');
            }
            
            // Render compact top items by base quantity
            const topBaseQtyContainer = document.getElementById('topItemsByBaseQty');
            if (topBaseQtyContainer) {
                topBaseQtyContainer.innerHTML = topByBaseQty.map((item, index) => `
                    <div class="flex items-center justify-between p-2 bg-white rounded border border-green-100 hover:border-green-200 transition-colors">
                        <div class="flex items-center space-x-2">
                            <div class="w-5 h-5 bg-green-500 text-white rounded-full flex items-center justify-center text-xs font-bold">
                                ${index + 1}
                            </div>
                            <div class="flex-1 min-w-0">
                                <div class="text-xs font-semibold text-green-900 truncate" title="${item.name}">
                                    ${item.name.length > 15 ? item.name.substring(0, 15) + '...' : item.name}
                                </div>
                                <div class="text-xs text-green-600 font-mono opacity-75">
                                    ${item.id} • ${item.locations}x
                                </div>
                            </div>
                        </div>
                        <div class="text-right">
                            <div class="text-sm font-bold text-green-900">
                                ${item.totalBaseQty.toLocaleString()}
                            </div>
                            <div class="text-xs text-green-600 opacity-75">
                                ${item.baseUom}
                            </div>
                        </div>
                    </div>
                `).join('');
            }
            
            // Render mobile top items (compact top 10)
            const mobileTopContainer = document.getElementById('mobileTopItems');
            if (mobileTopContainer) {
                const combinedTop = topByQty.slice(0, 10); // Show top 10 for mobile
                mobileTopContainer.innerHTML = combinedTop.map((item, index) => `
                    <div class="flex items-center justify-between p-2 bg-white rounded border border-purple-100 hover:border-purple-200 transition-colors">
                        <div class="flex items-center space-x-2">
                            <div class="w-5 h-5 bg-purple-500 text-white rounded-full flex items-center justify-center text-xs font-bold">
                                ${index + 1}
                            </div>
                            <div class="flex-1 min-w-0">
                                <div class="text-xs font-semibold text-purple-900 truncate" title="${item.name}">
                                    ${item.name.length > 12 ? item.name.substring(0, 12) + '...' : item.name}
                                </div>
                                <div class="text-xs text-purple-600 font-mono opacity-75">
                                    ${item.id}
                                </div>
                            </div>
                        </div>
                        <div class="text-right">
                            <div class="text-sm font-bold text-purple-900">
                                ${item.totalQty.toLocaleString()}
                            </div>
                            <div class="text-xs text-purple-600 opacity-75">
                                ${item.uom}
                            </div>
                        </div>
                    </div>
                `).join('');
            }
        }
function filterData() {
    const input = document.getElementById('searchInput');
    const suggestionList = document.getElementById('suggestionList');
    const searchTerm = input.value.toLowerCase();

    // Jika kolom pencarian kosong
    if (searchTerm === '') {
        filteredData = [...allData];
        suggestionList.classList.add('hidden');
    } else {
        // Filter berdasarkan nama atau ID
        filteredData = allData.filter(row => {
            const itemName = (row[2] || '').toLowerCase();
            const itemId = (row[1] || '').toLowerCase();
            return itemName.includes(searchTerm) || itemId.includes(searchTerm);
        });

        // Tampilkan semua hasil (tanpa dibatasi 10)
        const suggestions = filteredData.map(row => {
            const name = highlightMatch(row[2] || '', searchTerm);
            const id = highlightMatch(row[1] || '', searchTerm);
            return `
                <li class="px-4 py-2 hover:bg-neutral-100 cursor-pointer text-sm text-neutral-700"
                    onclick="selectSuggestion('${row[2]}')">
                    <div class="flex flex-col">
                        <span class="font-medium">${name}</span>
                        <span class="text-xs text-neutral-500">ID: ${id}</span>
                    </div>
                </li>
            `;
        }).join('');

        // Jika tidak ada hasil
        suggestionList.innerHTML = suggestions || 
            '<li class="px-4 py-2 text-neutral-400 text-sm">No results found</li>';
        
        suggestionList.classList.remove('hidden');
    }

    // Update tampilan utama
    currentPage = 1;
    updateDashboard();
    renderTable();
    updatePagination();
}

// Klik suggestion → isi input & jalankan filter ulang
function selectSuggestion(value) {
    document.getElementById('searchInput').value = value;
    document.getElementById('suggestionList').classList.add('hidden');
    filterData();
}

// Saat input fokus → tampilkan suggestion jika ada
function showSuggestions() {
    const list = document.getElementById('suggestionList');
    if (list.innerHTML.trim() !== '') {
        list.classList.remove('hidden');
    }
}

// Sembunyikan suggestion saat blur (delay agar klik terdeteksi)
function hideSuggestionsDelayed() {
    setTimeout(() => document.getElementById('suggestionList').classList.add('hidden'), 150);
}

// Fungsi highlight teks yang cocok
function highlightMatch(text, term) {
    const regex = new RegExp(`(${term})`, 'gi');
    return text.replace(regex, '<span class="font-semibold text-primary-600">$1</span>');
}

        function sortData(column) {
            if (sortColumn === column) {
                sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
            } else {
                sortColumn = column;
                sortDirection = 'asc';
            }

            // Update header indicators
            document.querySelectorAll('.sortable').forEach(header => {
                header.classList.remove('sort-asc', 'sort-desc');
            });
            
            const currentHeader = document.querySelector(`[data-column="${column}"]`);
            currentHeader.classList.add(sortDirection === 'asc' ? 'sort-asc' : 'sort-desc');

            filteredData.sort((a, b) => {
                let aVal = a[column] || '';
                let bVal = b[column] || '';
                
                // Try to parse as numbers for numeric columns
                if (column === 10 || column === 12) { // Qty and Base Qty columns
                    aVal = parseFloat(aVal) || 0;
                    bVal = parseFloat(bVal) || 0;
                } else {
                    aVal = aVal.toString().toLowerCase();
                    bVal = bVal.toString().toLowerCase();
                }
                
                if (aVal < bVal) return sortDirection === 'asc' ? -1 : 1;
                if (aVal > bVal) return sortDirection === 'asc' ? 1 : -1;
                return 0;
            });

            renderTable();
        }

        function renderTable() {
            const tbody = document.getElementById('tableBody');
            const mobileContainer = document.getElementById('mobileCardContainer');
            const startIndex = (currentPage - 1) * itemsPerPage;
            const endIndex = startIndex + itemsPerPage;
            const pageData = filteredData.slice(startIndex, endIndex);

            if (pageData.length === 0) {
                tbody.innerHTML = `
                    <tr>
                        <td colspan="14" class="px-8 py-12 text-center">
                            <div class="flex flex-col items-center space-y-4">
                                <svg class="w-16 h-16 text-neutral-300" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1" d="M20 13V6a2 2 0 00-2-2H6a2 2 0 00-2 2v7m16 0v5a2 2 0 01-2 2H6a2 2 0 01-2-2v-5m16 0h-2.586a1 1 0 00-.707.293l-2.414 2.414a1 1 0 01-.707.293h-3.172a1 1 0 01-.707-.293l-2.414-2.414A1 1 0 006.586 13H4"></path>
                                </svg>
                                <div class="text-center">
                                    <h3 class="text-lg font-semibold text-neutral-700">No Data Found</h3>
                                    <p class="text-neutral-500 mt-1">No inventory records match your search criteria</p>
                                </div>
                            </div>
                        </td>
                    </tr>
                `;
                mobileContainer.innerHTML = `
                    <div class="text-center py-12">
                        <div class="flex flex-col items-center space-y-4">
                            <svg class="w-16 h-16 text-neutral-300" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1" d="M20 13V6a2 2 0 00-2-2H6a2 2 0 00-2 2v7m16 0v5a2 2 0 01-2 2H6a2 2 0 01-2-2v-5m16 0h-2.586a1 1 0 00-.707.293l-2.414 2.414a1 1 0 01-.707.293h-3.172a1 1 0 01-.707-.293l-2.414-2.414A1 1 0 006.586 13H4"></path>
                            </svg>
                            <div class="text-center">
                                <h3 class="text-lg font-semibold text-neutral-700">No Data Found</h3>
                                <p class="text-neutral-500 mt-1">No inventory records match your search criteria</p>
                            </div>
                        </div>
                    </div>
                `;
                return;
            }

            // Render desktop table
            tbody.innerHTML = pageData.map(row => {
                const qty = parseFloat(row[10]) || 0;
                let rowClass = '';
                let statusBadge = '';
                
                if (qty < 10) {
                    rowClass = 'status-critical';
                    statusBadge = '<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-danger-100 text-danger-800">Critical</span>';
                } else if (qty >= 10 && qty <= 50) {
                    rowClass = 'status-warning';
                    statusBadge = '<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-warning-100 text-warning-800">Low</span>';
                } else {
                    rowClass = 'status-normal';
                    statusBadge = '<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-success-100 text-success-800">Normal</span>';
                }

                return `
                    <tr class="${rowClass}">
                        <td class="px-6 py-4 text-sm font-medium text-neutral-900">${row[0] || '-'}</td>
                        <td class="px-6 py-4 text-sm text-neutral-700 font-mono">${row[1] || '-'}</td>
                        <td class="px-6 py-4 text-sm font-medium text-neutral-900">${row[2] || '-'}</td>
                        <td class="px-6 py-4 text-sm text-neutral-700">${row[3] || '-'}</td>
                        <td class="px-6 py-4 text-sm text-neutral-700">${row[4] || '-'}</td>
                        <td class="px-6 py-4 text-sm text-neutral-700">${row[5] || '-'}</td>
                        <td class="px-6 py-4 text-sm text-neutral-700">${row[6] || '-'}</td>
                        <td class="px-6 py-4 text-sm text-neutral-700">${row[7] || '-'}</td>
                        <td class="px-6 py-4 text-sm text-neutral-700 font-mono">${row[8] || '-'}</td>
                        <td class="px-6 py-4 text-sm text-neutral-700">${row[9] || '-'}</td>
                        <td class="px-6 py-4 text-sm">
                            <div class="flex items-center space-x-3">
                                <span class="text-lg font-bold text-neutral-900">${qty.toLocaleString()}</span>
                                ${statusBadge}
                            </div>
                        </td>
                        <td class="px-6 py-4 text-sm text-neutral-700">${row[11] || '-'}</td>
                        <td class="px-6 py-4 text-sm font-semibold text-neutral-900">${(parseFloat(row[12]) || 0).toLocaleString()}</td>
                        <td class="px-6 py-4 text-sm text-neutral-700">${row[13] || '-'}</td>
                    </tr>
                `;
            }).join('');

            // Render mobile cards
            mobileContainer.innerHTML = pageData.map(row => {
                const qty = parseFloat(row[10]) || 0;
                const baseQty = parseFloat(row[12]) || 0;
                let cardClass = '';
                let statusBadge = '';
                let statusClass = '';
                
                if (qty < 10) {
                    cardClass = 'low-qty';
                    statusBadge = 'CRITICAL';
                    statusClass = 'mobile-status-critical';
                } else if (qty >= 10 && qty <= 50) {
                    cardClass = 'medium-qty';
                    statusBadge = 'LOW STOCK';
                    statusClass = 'mobile-status-warning';
                } else {
                    cardClass = 'high-qty';
                    statusBadge = 'NORMAL';
                    statusClass = 'mobile-status-normal';
                }

                return `
                    <div class="mobile-item ${cardClass}">
                        <div class="mobile-header">
                            <div class="mobile-item-title">${row[2] || 'Unknown Item'}</div>
                            <div class="mobile-item-id">ID: ${row[1] || '-'}</div>
                        </div>
                        
                        <div class="mobile-qty-container">
                            <div class="mobile-qty">${qty.toLocaleString()}</div>
                            <div class="mobile-status-badge ${statusClass}">${statusBadge}</div>
                        </div>
                        
                        <div class="mobile-field">
                            <span class="mobile-label">Location</span>
                            <span class="mobile-value" title="${row[0] || '-'}">${row[0] || '-'}</span>
                        </div>
                        <div class="mobile-field">
                            <span class="mobile-label">Pallet</span>
                            <span class="mobile-value" title="${row[3] || '-'}">${row[3] || '-'}</span>
                        </div>
                        <div class="mobile-field">
                            <span class="mobile-label">Pallet No</span>
                            <span class="mobile-value" title="${row[4] || '-'}">${row[4] || '-'}</span>
                        </div>
                        <div class="mobile-field">
                            <span class="mobile-label">Batch</span>
                            <span class="mobile-value" title="${row[8] || '-'}">${row[8] || '-'}</span>
                        </div>
                        <div class="mobile-field">
                            <span class="mobile-label">Container</span>
                            <span class="mobile-value" title="${row[9] || '-'}">${row[9] || '-'}</span>
                        </div>
                        <div class="mobile-field">
                            <span class="mobile-label">UOM</span>
                            <span class="mobile-value" title="${row[11] || '-'}">${row[11] || '-'}</span>
                        </div>
                        <div class="mobile-field">
                            <span class="mobile-label">Base Qty</span>
                            <span class="mobile-value" title="${baseQty.toLocaleString()} ${row[13] || ''}">${baseQty.toLocaleString()} ${row[13] || ''}</span>
                        </div>
                        <div class="mobile-field">
                            <span class="mobile-label">Trans Date</span>
                            <span class="mobile-value" title="${row[5] || '-'}">${row[5] || '-'}</span>
                        </div>
                        <div class="mobile-field">
                            <span class="mobile-label">Prod Date</span>
                            <span class="mobile-value" title="${row[6] || '-'}">${row[6] || '-'}</span>
                        </div>
                        <div class="mobile-field">
                            <span class="mobile-label">Exp Date</span>
                            <span class="mobile-value" title="${row[7] || '-'}">${row[7] || '-'}</span>
                        </div>
                    </div>
                `;
            }).join('');

            // Update mobile pagination
            updateMobilePagination();
        }

        function updatePagination() {
            const totalPages = Math.ceil(filteredData.length / itemsPerPage);
            const startIndex = (currentPage - 1) * itemsPerPage + 1;
            const endIndex = Math.min(currentPage * itemsPerPage, filteredData.length);

            document.getElementById('showingStart').textContent = filteredData.length > 0 ? startIndex : 0;
            document.getElementById('showingEnd').textContent = endIndex;
            document.getElementById('totalRecords').textContent = filteredData.length;

            // Generate pagination buttons
            const paginationNav = document.getElementById('paginationNav');
            let paginationHTML = '';

            // Previous button
            paginationHTML += `
                <button onclick="previousPage()" ${currentPage === 1 ? 'disabled' : ''} 
                        class="pagination-corporate relative inline-flex items-center px-4 py-2 rounded-lg text-sm font-semibold ${currentPage === 1 ? 'cursor-not-allowed opacity-50' : ''}">
                    <svg class="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 19l-7-7 7-7"></path>
                    </svg>
                    Prev
                </button>
            `;

            // Page numbers
            const maxVisiblePages = 5;
            let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
            let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);
            
            if (endPage - startPage + 1 < maxVisiblePages) {
                startPage = Math.max(1, endPage - maxVisiblePages + 1);
            }

            for (let i = startPage; i <= endPage; i++) {
                paginationHTML += `
                    <button onclick="goToPage(${i})" 
                            class="pagination-corporate relative inline-flex items-center px-4 py-2 rounded-lg text-sm font-semibold ${i === currentPage ? 'active' : ''}">
                        ${i}
                    </button>
                `;
            }

            // Next button
            paginationHTML += `
                <button onclick="nextPage()" ${currentPage === totalPages ? 'disabled' : ''} 
                        class="pagination-corporate relative inline-flex items-center px-4 py-2 rounded-lg text-sm font-semibold ${currentPage === totalPages ? 'cursor-not-allowed opacity-50' : ''}">
                    Next
                    <svg class="w-4 h-4 ml-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"></path>
                    </svg>
                </button>
            `;

            paginationNav.innerHTML = paginationHTML;
        }

        function previousPage() {
            if (currentPage > 1) {
                currentPage--;
                renderTable();
                updatePagination();
            }
        }

        function nextPage() {
            const totalPages = Math.ceil(filteredData.length / itemsPerPage);
            if (currentPage < totalPages) {
                currentPage++;
                renderTable();
                updatePagination();
            }
        }

        function goToPage(page) {
            currentPage = page;
            renderTable();
            updatePagination();
        }

        function updateMobilePagination() {
            const totalPages = Math.ceil(filteredData.length / itemsPerPage);
            const mobilePagination = document.getElementById('mobilePagination');
            
            if (totalPages <= 1) {
                mobilePagination.innerHTML = '';
                return;
            }
            
            let paginationHTML = '';
            
            // Previous button
            paginationHTML += `
                <button onclick="previousPage()" ${currentPage === 1 ? 'disabled' : ''} 
                        class="mobile-page-btn ${currentPage === 1 ? '' : 'hover:bg-gray-50'}">
                    ‹ Prev
                </button>
            `;
            
            // Current page info
            paginationHTML += `
                <span class="mobile-page-info">
                    ${currentPage} / ${totalPages}
                </span>
            `;
            
            // Next button
            paginationHTML += `
                <button onclick="nextPage()" ${currentPage === totalPages ? 'disabled' : ''} 
                        class="mobile-page-btn ${currentPage === totalPages ? '' : 'hover:bg-gray-50'}">
                    Next ›
                </button>
            `;
            
            mobilePagination.innerHTML = paginationHTML;
        }

        function updateTimestamp() {
            const now = new Date();
            const timeString = now.toLocaleString('en-US', {
                year: 'numeric',
                month: 'short',
                day: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            });
            document.getElementById('lastUpdated').textContent = timeString;
        }

        function getTimeAgo(date) {
            const now = new Date();
            const diffInSeconds = Math.floor((now - date) / 1000);
            
            if (diffInSeconds < 60) {
                return `${diffInSeconds} seconds ago`;
            } else if (diffInSeconds < 3600) {
                const minutes = Math.floor(diffInSeconds / 60);
                return `${minutes} minute${minutes > 1 ? 's' : ''} ago`;
            } else if (diffInSeconds < 86400) {
                const hours = Math.floor(diffInSeconds / 3600);
                return `${hours} hour${hours > 1 ? 's' : ''} ago`;
            } else {
                const days = Math.floor(diffInSeconds / 86400);
                return `${days} day${days > 1 ? 's' : ''} ago`;
            }
        }

        // Auto-refresh functionality
        let autoRefreshInterval;
        
        function startAutoRefresh() {
            // Check for updates every 30 seconds
            autoRefreshInterval = setInterval(async () => {
                try {
                    // Check if spreadsheet was modified
                    const driveUrl = `https://www.googleapis.com/drive/v3/files/${SPREADSHEET_ID}?key=${API_KEY}&fields=modifiedTime`;
                    const response = await fetch(driveUrl);
                    if (response.ok) {
                        const data = await response.json();
                        const lastModified = new Date(data.modifiedTime);
                        
                        // Update time ago display
                        const lastModifiedElement = document.getElementById('lastUpdated');
                        if (lastModifiedElement && lastModified) {
                            const timeAgo = getTimeAgo(lastModified);
                            const currentText = lastModifiedElement.innerHTML;
                            if (currentText.includes('font-bold')) {
                                // Update only the time ago part
                                lastModifiedElement.innerHTML = currentText.replace(/\(.*?\)/, `(${timeAgo})`);
                            }
                        }
                        
                        // Update mobile time ago display
                        const mobileLastUpdated = document.getElementById('mobileLastUpdated');
                        if (mobileLastUpdated && lastModified) {
                            const timeAgo = getTimeAgo(lastModified);
                            const currentText = mobileLastUpdated.textContent;
                            if (currentText.includes('(')) {
                                // Update only the time ago part
                                mobileLastUpdated.textContent = currentText.replace(/\(.*?\)/, `(${timeAgo})`);
                            }
                        }
                        
                        // Update system status
                        const minutesAgo = (new Date() - lastModified) / (1000 * 60);
                        const systemStatus = document.getElementById('systemStatus');
                        const statusIndicator = systemStatus.parentElement.querySelector('.animate-pulse');
                        
                        // Mobile status elements
                        const mobileStatusDot = document.getElementById('mobileStatusDot');
                        const mobileStatusText = document.getElementById('mobileStatusText');
                        
                        if (minutesAgo < 5) {
                            // Desktop status
                            systemStatus.textContent = 'LIVE';
                            systemStatus.className = 'text-2xl font-bold text-green-300 drop-shadow-lg';
                            statusIndicator.className = 'w-3 h-3 bg-green-400 rounded-full animate-pulse shadow-lg';
                            
                            // Mobile status
                            if (mobileStatusDot && mobileStatusText) {
                                mobileStatusDot.className = 'realtime-status-dot live';
                                mobileStatusText.textContent = 'LIVE';
                            }
                        } else if (minutesAgo < 30) {
                            // Desktop status
                            systemStatus.textContent = 'RECENT';
                            systemStatus.className = 'text-2xl font-bold text-yellow-300 drop-shadow-lg';
                            statusIndicator.className = 'w-3 h-3 bg-yellow-400 rounded-full animate-pulse shadow-lg';
                            
                            // Mobile status
                            if (mobileStatusDot && mobileStatusText) {
                                mobileStatusDot.className = 'realtime-status-dot recent';
                                mobileStatusText.textContent = 'RECENT';
                            }
                        } else {
                            // Desktop status
                            systemStatus.textContent = 'STALE';
                            systemStatus.className = 'text-2xl font-bold text-orange-300 drop-shadow-lg';
                            statusIndicator.className = 'w-3 h-3 bg-orange-400 rounded-full animate-pulse shadow-lg';
                            
                            // Mobile status
                            if (mobileStatusDot && mobileStatusText) {
                                mobileStatusDot.className = 'realtime-status-dot stale';
                                mobileStatusText.textContent = 'STALE';
                            }
                        }
                    }
                } catch (error) {
                    console.log('Auto-refresh check failed:', error);
                }
            }, 30000); // 30 seconds
        }

        function stopAutoRefresh() {
            if (autoRefreshInterval) {
                clearInterval(autoRefreshInterval);
                autoRefreshInterval = null;
            }
        }

        function refreshData() {
            const refreshBtn = document.getElementById('refreshBtn');
            const systemStatus = document.getElementById('systemStatus');
            
            // Update button state
            refreshBtn.disabled = true;
            refreshBtn.innerHTML = `
                <!-- Background Animation -->
                <div class="absolute inset-0 bg-gradient-to-r from-white/0 via-white/20 to-white/0 -translate-x-full group-hover:translate-x-full transition-transform duration-700"></div>
                
                <!-- Loading Content -->
                <div class="relative flex items-center space-x-3">
                    <div class="p-2 bg-white/20 rounded-lg backdrop-blur-sm">
                        <svg class="w-5 h-5 animate-spin" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"></path>
                        </svg>
                    </div>
                    <div class="flex flex-col items-start">
                        <span class="text-white font-black text-base">Refreshing...</span>
                        <span class="text-white/70 text-xs font-medium">Loading data</span>
                    </div>
                </div>
                
                <!-- Status Indicator -->
                <div class="absolute top-2 right-2">
                    <div class="w-2 h-2 bg-yellow-400 rounded-full animate-pulse"></div>
                </div>
            `;
            
            // Update system status to show refreshing
            systemStatus.textContent = 'REFRESHING';
            systemStatus.className = 'text-2xl font-bold text-yellow-300 drop-shadow-lg animate-pulse';
            
            loadData().then(() => {
                // Reset system status to online
                systemStatus.textContent = 'ONLINE';
                systemStatus.className = 'text-2xl font-bold text-white drop-shadow-lg';
                
                // Reset button state
                refreshBtn.disabled = false;
                refreshBtn.innerHTML = `
                    <!-- Background Animation -->
                    <div class="absolute inset-0 bg-gradient-to-r from-white/0 via-white/20 to-white/0 -translate-x-full group-hover:translate-x-full transition-transform duration-700"></div>
                    
                    <!-- Content -->
                    <div class="relative flex items-center space-x-3">
                        <div class="p-2 bg-white/20 rounded-lg backdrop-blur-sm group-hover:bg-white/30 transition-colors">
                            <svg class="w-5 h-5 transition-transform duration-300 group-hover:rotate-180" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"></path>
                            </svg>
                        </div>
                        <div class="flex flex-col items-start">
                            <span class="text-white font-black text-base">Refresh Data</span>
                            <span class="text-white/70 text-xs font-medium">Update inventory</span>
                        </div>
                    </div>
                    
                    <!-- Status Indicator -->
                    <div class="absolute top-2 right-2">
                        <div class="w-2 h-2 bg-green-400 rounded-full animate-pulse"></div>
                    </div>
                `;
            }).catch(() => {
                // Set system status to error
                systemStatus.textContent = 'ERROR';
                systemStatus.className = 'text-2xl font-bold text-red-300 drop-shadow-lg';
                
                // Reset button state with error
                refreshBtn.disabled = false;
                refreshBtn.innerHTML = `
                    <!-- Background Animation -->
                    <div class="absolute inset-0 bg-gradient-to-r from-white/0 via-white/20 to-white/0 -translate-x-full group-hover:translate-x-full transition-transform duration-700"></div>
                    
                    <!-- Content -->
                    <div class="relative flex items-center space-x-3">
                        <div class="p-2 bg-white/20 rounded-lg backdrop-blur-sm group-hover:bg-white/30 transition-colors">
                            <svg class="w-5 h-5 transition-transform duration-300 group-hover:rotate-180" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"></path>
                            </svg>
                        </div>
                        <div class="flex flex-col items-start">
                            <span class="text-white font-black text-base">Refresh Data</span>
                            <span class="text-white/70 text-xs font-medium">Update inventory</span>
                        </div>
                    </div>
                    
                    <!-- Status Indicator -->
                    <div class="absolute top-2 right-2">
                        <div class="w-2 h-2 bg-red-400 rounded-full animate-pulse"></div>
                    </div>
                `;
            });
        }



        function exportToPDF() {
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF('l', 'mm', 'a4'); // landscape orientation
            
            doc.setFontSize(16);
            doc.text('Stok GIM', 14, 15);
            doc.setFontSize(10);
            doc.text(`Generated on: ${new Date().toLocaleString()}`, 14, 25);
            
            const headers = [['Location ID', 'Item ID', 'Item Name', 'Pallet', 'Pallet No', 'Trans Date', 'Prod Date', 'Exp Date', 'Batch', 'Container', 'Qty', 'UOM', 'Base Qty', 'Base UOM']];
            
            const data = filteredData.map(row => [
                row[0] || '', row[1] || '', row[2] || '', row[3] || '', row[4] || '',
                row[5] || '', row[6] || '', row[7] || '', row[8] || '', row[9] || '',
                row[10] || '', row[11] || '', row[12] || '', row[13] || ''
            ]);
            
            doc.autoTable({
                head: headers,
                body: data,
                startY: 30,
                styles: { fontSize: 8 },
                headStyles: { fillColor: [66, 139, 202] },
                didDrawCell: function(data) {
                    // Highlight rows based on Qty
                    if (data.column.index === 10 && data.section === 'body') {
                        const qty = parseFloat(data.cell.text[0]) || 0;
                        if (qty < 10) {
                            data.cell.styles.fillColor = [252, 228, 236]; // Light red
                        } else if (qty >= 10 && qty <= 50) {
                            data.cell.styles.fillColor = [255, 249, 196]; // Light yellow
                        }
                    }
                }
            });
            
            doc.save(`Stok_GIM_${new Date().toISOString().split('T')[0]}.pdf`);
        }

        function toggleTopItems() {
            const content = document.getElementById('topItemsContent');
            const icon = document.getElementById('topItemsToggleIcon');
            
            if (content.classList.contains('hidden')) {
                // Show content
                content.classList.remove('hidden');
                content.classList.add('animate-fadeIn');
                icon.style.transform = 'rotate(180deg)';
            } else {
                // Hide content
                content.classList.add('hidden');
                content.classList.remove('animate-fadeIn');
                icon.style.transform = 'rotate(0deg)';
            }
        }

        function exportToExcel() {
            // Create workbook and worksheet
            const wb = XLSX.utils.book_new();
            
            // Prepare headers
            const headers = [
                'Location ID', 'Item ID', 'Item Name', 'Pallet', 'Pallet No', 
                'Transaction Date', 'Production Date', 'Expiry Date', 'Batch', 
                'Container ID', 'Quantity', 'UOM', 'Base Qty', 'Base UOM'
            ];
            
            // Prepare data with headers
            const wsData = [headers];
            
            // Add filtered data
            filteredData.forEach(row => {
                wsData.push([
                    row[0] || '', row[1] || '', row[2] || '', row[3] || '', row[4] || '',
                    row[5] || '', row[6] || '', row[7] || '', row[8] || '', row[9] || '',
                    parseFloat(row[10]) || 0, row[11] || '', parseFloat(row[12]) || 0, row[13] || ''
                ]);
            });
            
            // Create worksheet
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            
            // Set column widths
            const colWidths = [
                { wch: 12 }, // Location ID
                { wch: 15 }, // Item ID
                { wch: 25 }, // Item Name
                { wch: 12 }, // Pallet
                { wch: 12 }, // Pallet No
                { wch: 15 }, // Transaction Date
                { wch: 15 }, // Production Date
                { wch: 15 }, // Expiry Date
                { wch: 15 }, // Batch
                { wch: 15 }, // Container ID
                { wch: 12 }, // Quantity
                { wch: 8 },  // UOM
                { wch: 12 }, // Base Qty
                { wch: 10 }  // Base UOM
            ];
            ws['!cols'] = colWidths;
            
            // Style the header row
            const headerStyle = {
                font: { bold: true, color: { rgb: "FFFFFF" } },
                fill: { fgColor: { rgb: "4472C4" } },
                alignment: { horizontal: "center", vertical: "center" }
            };
            
            // Apply header styling
            for (let col = 0; col < headers.length; col++) {
                const cellRef = XLSX.utils.encode_cell({ r: 0, c: col });
                if (!ws[cellRef]) ws[cellRef] = {};
                ws[cellRef].s = headerStyle;
            }
            
            // Apply conditional formatting for quantity column (column K, index 10)
            for (let row = 1; row < wsData.length; row++) {
                const qty = parseFloat(wsData[row][10]) || 0;
                const cellRef = XLSX.utils.encode_cell({ r: row, c: 10 });
                
                if (!ws[cellRef]) ws[cellRef] = {};
                
                if (qty < 10) {
                    // Critical - Red background
                    ws[cellRef].s = {
                        fill: { fgColor: { rgb: "FFCDD2" } },
                        font: { bold: true, color: { rgb: "C62828" } }
                    };
                } else if (qty >= 10 && qty <= 50) {
                    // Warning - Yellow background
                    ws[cellRef].s = {
                        fill: { fgColor: { rgb: "FFF9C4" } },
                        font: { bold: true, color: { rgb: "F57F17" } }
                    };
                } else {
                    // Normal - Green background
                    ws[cellRef].s = {
                        fill: { fgColor: { rgb: "C8E6C9" } },
                        font: { bold: true, color: { rgb: "2E7D32" } }
                    };
                }
            }
            
            // Add summary sheet
            const summaryData = [
                ['INVENTORY SUMMARY REPORT'],
                ['Generated on:', new Date().toLocaleString()],
                [''],
                ['METRICS', 'VALUE'],
                ['Total Records', filteredData.length],
                ['Total Inventory Quantity', filteredData.reduce((sum, row) => sum + (parseFloat(row[10]) || 0), 0)],
                ['Total Base Quantity', filteredData.reduce((sum, row) => sum + (parseFloat(row[12]) || 0), 0)],
                ['Unique Items', new Set(filteredData.map(row => row[1])).size],
                ['Low Stock Items (< 10)', filteredData.filter(row => (parseFloat(row[10]) || 0) < 10).length],
                ['Medium Stock Items (10-50)', filteredData.filter(row => {
                    const qty = parseFloat(row[10]) || 0;
                    return qty >= 10 && qty <= 50;
                }).length],
                ['High Stock Items (> 50)', filteredData.filter(row => (parseFloat(row[10]) || 0) > 50).length]
            ];
            
            const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
            
            // Style summary sheet
            summaryWs['!cols'] = [{ wch: 25 }, { wch: 20 }];
            
            // Style title
            if (summaryWs['A1']) {
                summaryWs['A1'].s = {
                    font: { bold: true, size: 16, color: { rgb: "1565C0" } },
                    alignment: { horizontal: "center" }
                };
            }
            
            // Style headers
            if (summaryWs['A4']) {
                summaryWs['A4'].s = headerStyle;
            }
            if (summaryWs['B4']) {
                summaryWs['B4'].s = headerStyle;
            }
            
            // Add worksheets to workbook
            XLSX.utils.book_append_sheet(wb, summaryWs, "Summary");
            XLSX.utils.book_append_sheet(wb, ws, "Inventory Data");
            
            // Generate filename with timestamp
            const timestamp = new Date().toISOString().split('T')[0];
            const filename = `InventoryPro_Report_${timestamp}.xlsx`;
            
            // Save file
            XLSX.writeFile(wb, filename);
        }
