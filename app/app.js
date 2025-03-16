(function () {
    'use strict';

    // Global variables to track conversation
    let currentConversationId = null;
    let latestMessageId = null;

    // Helper functions defined first
    function showStatus(message, isSuccess, type = null) {
        const statusDiv = document.getElementById('status');
        statusDiv.textContent = message;
        statusDiv.className = 'status-message';

        // Clear any existing classes
        statusDiv.classList.remove('status-success', 'status-error', 'status-info', 'status-loading');

        // Set the appropriate class based on message type
        if (isSuccess && type === 'info') {
            statusDiv.classList.add('status-info');
        } else if (isSuccess && type === 'loading') {
            statusDiv.classList.add('status-loading');
        } else if (isSuccess) {
            statusDiv.classList.add('status-success');
        } else {
            statusDiv.classList.add('status-error');
        }

        // Make sure the message is visible
        statusDiv.style.display = 'block';

        // Add a subtle entrance animation by briefly adding and removing a class
        statusDiv.classList.add('status-animate');
        setTimeout(() => {
            statusDiv.classList.remove('status-animate');
        }, 300);

        // Log to console for debugging
        if (isSuccess) {
            console.log('Status:', message);
        } else {
            console.error('Status Error:', message);
        }

        // For loading states, return a function to update the message
        if (type === 'loading') {
            return (newMessage) => {
                if (newMessage) {
                    statusDiv.textContent = newMessage;
                }
            };
        }
    }

    function displayQueryResults(data) {
        const resultsDiv = document.getElementById('query-results');
        resultsDiv.innerHTML = '';

        if (!data || !data.length) {
            resultsDiv.innerHTML = '<p>No results found</p>';
            return;
        }

        // Store data globally for pagination
        window.queryResultData = data;
        window.currentPage = 1;

        // Create a function to render a page that can be reused
        window.renderResultPage = function(page) {
            const rowsPerPage = parseInt(document.getElementById('rows-per-page').value) || 25;
            const totalPages = Math.ceil(window.queryResultData.length / rowsPerPage);

            // Validate page number
            if (page < 1) page = 1;
            if (page > totalPages) page = totalPages;

            window.currentPage = page;

            // Clear the results area
            resultsDiv.innerHTML = '';

            // Create table
            const table = document.createElement('table');
            table.className = 'results-table';

            // Create header row
            const headerRow = document.createElement('tr');
            const headers = Object.keys(window.queryResultData[0]);

            headers.forEach(header => {
                const th = document.createElement('th');
                th.textContent = header;
                headerRow.appendChild(th);
            });

            table.appendChild(headerRow);

            // Calculate slice of data to show
            const startIndex = (page - 1) * rowsPerPage;
            const endIndex = Math.min(startIndex + rowsPerPage, window.queryResultData.length);
            const pageData = window.queryResultData.slice(startIndex, endIndex);

            // Create data rows for current page
            pageData.forEach(row => {
                const tr = document.createElement('tr');

                headers.forEach(header => {
                    const td = document.createElement('td');
                    const value = row[header];

                    if (value === null || value === undefined) {
                        td.textContent = '';
                    } else if (typeof value === 'object' && value instanceof Date) {
                        td.textContent = value.toLocaleDateString();
                        td.title = value.toLocaleDateString(); // Add tooltip for dates
                    } else {
                        td.textContent = value.toString();
                        td.title = value.toString(); // Add tooltip for all other values
                    }

                    tr.appendChild(td);
                });

                table.appendChild(tr);
            });

            resultsDiv.appendChild(table);

            // Add row count
            const rowCount = document.createElement('p');
            rowCount.className = 'row-count';
            rowCount.textContent = `Showing rows ${startIndex + 1}-${endIndex} of ${window.queryResultData.length}`;
            resultsDiv.appendChild(rowCount);

            // Update page info
            document.getElementById('page-info').textContent = `Page ${page} of ${totalPages}`;

            // Update button states
            document.getElementById('prev-page').disabled = page === 1;
            document.getElementById('next-page').disabled = page === totalPages;

            console.log(`Rendered page ${page} of ${totalPages}, showing ${pageData.length} rows (${startIndex + 1}-${endIndex})`);
        };

        // Initial render
        window.renderResultPage(1);

        // Set up event handlers for pagination
        document.getElementById('prev-page').onclick = function() {
            window.renderResultPage(window.currentPage - 1);
        };

        document.getElementById('next-page').onclick = function() {
            window.renderResultPage(window.currentPage + 1);
        };

        document.getElementById('rows-per-page').onchange = function() {
            // Re-render first page with new rows per page
            window.renderResultPage(1);
        };
    }

    async function writeToExcel(data, destination = {}) {
        return new Promise((resolve, reject) => {
            try {
                Excel.run(async (context) => {
                    let sheet;
                    let startCell;

                    if (!data || !data.length) {
                        await context.sync();
                        resolve();
                        return;
                    }

                    // Ensure destination is an object
                    destination = destination || {};

                    // Handle destination options
                    if (destination.newSheet) {
                        // Create new sheet option takes precedence
                        // Ignore append setting if new sheet is selected

                        // Create a new sheet
                        const newSheetName = `Results_${new Date().toISOString().replace(/[:.]/g, '_')}`.substring(0, 31);
                        sheet = context.workbook.worksheets.add(newSheetName);

                        // Activate the new sheet
                        sheet.activate();

                        // Use specified start cell if provided, otherwise default to A1
                        if (destination.startCell && destination.startCell.trim() !== '') {
                            try {
                                // For a new sheet, we only need to care about the cell address, not the sheet name
                                // Extract just the cell part if there's a sheet name included
                                let cellAddress = destination.startCell;
                                if (cellAddress.includes('!')) {
                                    cellAddress = cellAddress.split('!')[1];
                                }
                                startCell = sheet.getRange(cellAddress);
                            } catch (e) {
                                console.error("Invalid cell reference for new sheet:", e);
                                // Fall back to A1
                                startCell = sheet.getRange("A1");
                            }
                        } else {
                            // Default to A1
                            startCell = sheet.getRange("A1");
                        }

                        // Override append setting for new sheets
                        const actualAppend = false; // Never append in a new sheet

                        // Write data as new (not append)
                        await writeDataToSheet(sheet, startCell, data, actualAppend);
                    } else {
                        // Use existing sheet
                        sheet = context.workbook.worksheets.getActiveWorksheet();

                        if (destination.startCell && destination.startCell.trim() !== '') {
                            // Use specified start cell
                            try {
                                startCell = sheet.getRange(destination.startCell);
                            } catch (e) {
                                console.error("Invalid cell reference:", e);
                                // Fall back to A1
                                startCell = sheet.getRange("A1");
                            }
                        } else {
                            // Default to A1
                            startCell = sheet.getRange("A1");
                        }

                        // Use the append setting as provided
                        await writeDataToSheet(sheet, startCell, data, destination.appendData);
                    }

                    await context.sync();
                    resolve();
                });
            } catch (error) {
                reject(error);
            }
        });
    }

    // Helper function to write data to a sheet
    async function writeDataToSheet(sheet, startCell, data, appendData) {
        // Get the address of the start cell to determine row/column
        startCell.load("address");
        await sheet.context.sync();

        // Parse the address to get the starting point (e.g., "Sheet1!B3" -> col=2, row=3)
        const address = startCell.address;
        const match = address.match(/[A-Z]+|\d+/g);
        let startColumn = 0;
        let startRow = 0;

        if (match && match.length >= 2) {
            // Convert column letters to number (A=1, B=2, etc.)
            const colLetters = match[match.length - 2];
            for (let i = 0; i < colLetters.length; i++) {
                startColumn = startColumn * 26 + (colLetters.charCodeAt(i) - 64);
            }

            // Convert row to number
            startRow = parseInt(match[match.length - 1], 10);
        }

        const headers = Object.keys(data[0]);

        if (!appendData) {
            // Clear existing data if not appending
            const clearRange = startCell.getResizedRange(data.length + 1, headers.length - 1);
            clearRange.clear();

            // Write headers if not appending
            const headerRange = startCell.getResizedRange(0, headers.length - 1);
            headerRange.values = [headers];
            headerRange.format.font.bold = true;

            // Prepare data rows
            const rows = data.map(row =>
                headers.map(header => {
                    const value = row[header];
                    if (value === null || value === undefined) return '';
                    return value;
                })
            );

            // Write data rows
            if (rows.length > 0) {
                // Create a range that starts exactly at the row after header
                // and is exactly the size of our data (not with -1 adjustments)
                const dataRange = sheet.getRangeByIndexes(
                    startRow, // Row index (0-based)
                    startColumn - 1, // Column index (0-based)
                    rows.length, // Number of rows (exact)
                    headers.length // Number of columns (exact)
                );

                // Set the values
                dataRange.values = rows;
            }
        } else {
            // For append mode, find the last row with data
            let lastRow = startRow;

            // If we're starting at A1 or equivalent, check for existing data
            if (startRow === 1 && startColumn === 1) {
                const usedRange = sheet.getUsedRange();
                usedRange.load("rowCount");
                await sheet.context.sync();

                // If there's data, get the last row number
                if (usedRange.rowCount > 0) {
                    lastRow = usedRange.rowCount + 1; // +1 to start after the last row
                }
            }

            // If this is a fresh append, we need to add headers first
            const checkHeaderCell = sheet.getCell(startRow - 1, startColumn - 1);
            checkHeaderCell.load("values");
            await sheet.context.sync();

            // If the header cell is empty, add headers
            if (!checkHeaderCell.values[0][0] && lastRow === startRow) {
                const headerRange = startCell.getResizedRange(0, headers.length - 1);
                headerRange.values = [headers];
                headerRange.format.font.bold = true;
                lastRow++; // Move data one row down
            }

            // Prepare data rows
            const rows = data.map(row =>
                headers.map(header => {
                    const value = row[header];
                    if (value === null || value === undefined) return '';
                    return value;
                })
            );

            // Write data rows at the append position
            if (rows.length > 0) {
                const appendStartCell = sheet.getCell(lastRow - 1, startColumn - 1);
                const appendRange = appendStartCell.getResizedRange(rows.length - 1, headers.length - 1);
                appendRange.values = rows;
            }
        }

        // Auto-fit columns
        sheet.getUsedRange().format.autofitColumns();
    }

    async function processSqlQueryWithCellReferences(sqlQuery) {
        return Excel.run(async (context) => {
            // Regular expression to find cell references like ${A1} or ${Sheet1!B2}
            const cellRefRegex = /\${([^}]+)}/g;
            let match;
            let processedQuery = sqlQuery;

            // Find all cell references in the query
            const cellRefs = [];
            while ((match = cellRefRegex.exec(sqlQuery)) !== null) {
                cellRefs.push(match[1]);
            }

            // Process each cell reference
            for (const cellRef of cellRefs) {
                let range;

                // Log the cell reference being processed
                console.log("Processing cell reference:", cellRef);

                try {
                    // Check if it contains a sheet name (like "Sheet1!A1")
                    if (cellRef.includes('!')) {
                        const [sheetName, address] = cellRef.split('!');
                        console.log(`Using sheet: "${sheetName}", address: "${address}"`);

                        // Get the worksheet by name and then the range
                        const worksheet = context.workbook.worksheets.getItem(sheetName);
                        range = worksheet.getRange(address);
                    } else {
                        // No sheet specified, use active worksheet
                        console.log(`Using active worksheet, address: "${cellRef}"`);
                        range = context.workbook.worksheets.getActiveWorksheet().getRange(cellRef);
                    }

                    // Load cell values and number format
                    range.load(["values", "numberFormat"]);

                    // Execute the sync operation with better error handling
                    try {
                        await context.sync();
                        console.log("Cell values:", range.values);
                    } catch (syncError) {
                        console.error("Sync error:", syncError);
                        throw new Error(`Error accessing cell ${cellRef}: ${syncError.message}`);
                    }

                    // Convert cell values to SQL format based on data type and range size
                    const formattedValue = formatRangeValueForSql(range.values, range.numberFormat);
                    console.log(`Formatted value for ${cellRef}:`, formattedValue);

                    // Replace the reference in the query
                    processedQuery = processedQuery.replace(`\${${cellRef}}`, formattedValue);
                } catch (e) {
                    console.error("Error processing cell reference:", e);
                    throw new Error(`Invalid cell reference: ${cellRef} - ${e.message}`);
                }
            }

            console.log("Final processed query:", processedQuery);
            return processedQuery;
        });
    }

    function formatRangeValueForSql(values, numberFormat) {
        // Single cell
        if (values.length === 1 && values[0].length === 1) {
            return formatCellValueForSql(values[0][0], numberFormat[0][0]);
        }

        // Range of cells - create an array
        const formattedValues = [];
        for (let i = 0; i < values.length; i++) {
            for (let j = 0; j < values[i].length; j++) {
                formattedValues.push(formatCellValueForSql(values[i][j], numberFormat[i][j]));
            }
        }

        return "(" + formattedValues.join(", ") + ")";
    }

    function formatCellValueForSql(value, format) {
        if (value === null || value === undefined) {
            return "NULL";
        }

        // Check if it's a date
        if (typeof format === "string" &&
            (format.includes("d") || format.includes("m") || format.includes("y"))) {
            // Format as SQL date string
            return `'${new Date(value).toISOString().split('T')[0]}'`;
        }

        // Number
        if (typeof value === "number") {
            return value.toString();
        }

        // For values that look like numbers but are strings, convert them
        if (typeof value === "string" && !isNaN(value) && value.trim() !== '') {
            return value;
        }

        // Boolean
        if (typeof value === "boolean") {
            return value ? "TRUE" : "FALSE";
        }

        // Default to string with proper escaping
        return `'${value.toString().replace(/'/g, "''")}'`;
    }

    function addMessageToConversation(content, type) {
        const conversationHistory = document.getElementById('conversation-history');

        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${type}-message`;

        const messageContent = document.createElement('div');
        messageContent.className = 'message-content';

        // For Genie messages that might contain HTML
        if (type === 'genie') {
            messageContent.innerHTML = content;
        } else {
            messageContent.textContent = content;
        }

        messageDiv.appendChild(messageContent);
        conversationHistory.appendChild(messageDiv);

        // Scroll to the bottom of the conversation
        conversationHistory.scrollTop = conversationHistory.scrollHeight;
    }

    // Save destination options when they change
    function saveDestinationOptions() {
        const options = {
            startCell: document.getElementById('result-start-cell').value,
            newSheet: document.getElementById('result-new-sheet').checked,
            appendData: document.getElementById('result-append-data').checked
        };
        localStorage.setItem('resultDestinationOptions', JSON.stringify(options));
    }

    // The initialize function must be run each time a new page is loaded
    Office.onReady(function (info) {
        console.log("Office.onReady called", info);
        if (info.host === Office.HostType.Excel) {
            console.log("Excel detected!");

            // Initialize button click handlers
            document.getElementById('run-query').onclick = runSqlQuery;
            document.getElementById('ask-genie').onclick = askGenieQuestion;
            document.getElementById('send-follow-up').onclick = sendFollowUpQuestion;
            document.getElementById('insert-cell-reference').onclick = insertCellReference;
            document.getElementById('select-result-cell').onclick = selectResultCell;
            document.getElementById('genie-select-result-cell').onclick = selectGenieResultCell;
            document.getElementById('prev-page').onclick = () => {}; // Will be overridden in displayQueryResults
            document.getElementById('next-page').onclick = () => {}; // Will be overridden in displayQueryResults
            document.getElementById('rows-per-page').onchange = () => {}; // Will be overridden in displayQueryResults

            // Set default host if saved in localStorage
            const savedHost = localStorage.getItem('databricksHost');
            if (savedHost) {
                document.getElementById('databricks-host').value = savedHost;
            }

            // Set up checkbox logic for results destination
            const newSheetCheckbox = document.getElementById('result-new-sheet');
            const appendDataCheckbox = document.getElementById('result-append-data');
            const startCellInput = document.getElementById('result-start-cell');
            const selectCellButton = document.getElementById('select-result-cell');

            // Set up checkbox logic for Genie results destination (same as in SQL mode but with Genie IDs)
            const genieNewSheetCheckbox = document.getElementById('genie-result-new-sheet');
            const genieAppendDataCheckbox = document.getElementById('genie-result-append-data');
            const genieStartCellInput = document.getElementById('genie-result-start-cell');
            const genieSelectCellButton = document.getElementById('genie-select-result-cell');

            // When "Create new sheet" is checked:
            // 1. Uncheck "Append data" but don't disable it
            // 2. Clear any existing cell location
            // 3. Keep cell input enabled but disable the bullseye selector
            newSheetCheckbox.addEventListener('change', function() {
                if (this.checked) {
                    // Simply uncheck append data but keep it enabled
                    appendDataCheckbox.checked = false;

                    // Clear the start cell input
                    startCellInput.value = '';
                    startCellInput.placeholder = 'A1 (default)';

                    // Keep the input field enabled so user can type a cell reference
                    startCellInput.disabled = false;

                    // But disable the selector button since user can't select from a sheet that doesn't exist yet
                    selectCellButton.disabled = true;
                } else {
                    // Reset placeholder
                    startCellInput.placeholder = 'A1';

                    // Re-enable the selector button when not creating a new sheet
                    selectCellButton.disabled = false;
                }

                // Save the options whenever they change
                saveDestinationOptions();
            });

            // When "Append data" is checked, simply uncheck "Create new sheet" but don't disable it
            appendDataCheckbox.addEventListener('change', function() {
                if (this.checked) {
                    newSheetCheckbox.checked = false;

                    // Enable both the input field and selector button
                    startCellInput.disabled = false;
                    selectCellButton.disabled = false;
                }

                // Save the options whenever they change
                saveDestinationOptions();
            });

            // When "Create new sheet" is checked:
            genieNewSheetCheckbox.addEventListener('change', function() {
                if (this.checked) {
                    // Simply uncheck append data but keep it enabled
                    genieAppendDataCheckbox.checked = false;

                    // Clear the start cell input
                    genieStartCellInput.value = '';
                    genieStartCellInput.placeholder = 'A1 (default)';

                    // Keep the input field enabled so user can type a cell reference
                    genieStartCellInput.disabled = false;

                    // But disable the selector button since user can't select from a sheet that doesn't exist yet
                    genieSelectCellButton.disabled = true;
                } else {
                    // Reset placeholder
                    genieStartCellInput.placeholder = 'A1';

                    // Re-enable the selector button when not creating a new sheet
                    genieSelectCellButton.disabled = false;
                }

                // Save the options whenever they change
                saveGenieDestinationOptions();
            });

            // When "Append data" is checked, simply uncheck "Create new sheet" but don't disable it
            genieAppendDataCheckbox.addEventListener('change', function() {
                if (this.checked) {
                    genieNewSheetCheckbox.checked = false;

                    // Enable both the input field and selector button
                    genieStartCellInput.disabled = false;
                    genieSelectCellButton.disabled = false;
                }

                // Save the options whenever they change
                saveGenieDestinationOptions();
            });

            // Load saved destination options from localStorage
            try {
                const savedOptions = JSON.parse(localStorage.getItem('resultDestinationOptions'));
                if (savedOptions) {
                    startCellInput.value = savedOptions.startCell || '';
                    newSheetCheckbox.checked = savedOptions.newSheet || false;
                    appendDataCheckbox.checked = savedOptions.appendData || false;

                    // Trigger change event to update UI state
                    const event = new Event('change');
                    newSheetCheckbox.dispatchEvent(event);
                }
            } catch (e) {
                console.error('Error loading saved destination options:', e);
            }

            try {
                const savedGenieOptions = JSON.parse(localStorage.getItem('genieResultDestinationOptions'));
                if (savedGenieOptions) {
                    genieStartCellInput.value = savedGenieOptions.startCell || '';
                    genieNewSheetCheckbox.checked = savedGenieOptions.newSheet || false;
                    genieAppendDataCheckbox.checked = savedGenieOptions.appendData || false;

                    // Trigger change event to update UI state
                    const event = new Event('change');
                    genieNewSheetCheckbox.dispatchEvent(event);
                }
            } catch (e) {
                console.error('Error loading saved Genie destination options:', e);
            }

            // Add event listeners to save options
            startCellInput.addEventListener('change', saveDestinationOptions);
            newSheetCheckbox.addEventListener('change', saveDestinationOptions);
            appendDataCheckbox.addEventListener('change', saveDestinationOptions);

            genieStartCellInput.addEventListener('change', saveGenieDestinationOptions);
            genieNewSheetCheckbox.addEventListener('change', saveGenieDestinationOptions);
            genieAppendDataCheckbox.addEventListener('change', saveGenieDestinationOptions);

            document.getElementById('status').textContent = "BrickSheet is ready, let the data games begin!";
            document.getElementById('status').className = "status-message status-success";
        } else {
            console.log("Not running in Excel", info);
            document.getElementById('status').textContent = "Not running in Excel. Host: " + (info.host || "unknown");
            document.getElementById('status').className = "status-message status-error";
        }
    });

    // Traditional SQL query functionality
    async function runSqlQuery() {
        try {
            // Clear previous query results
            const resultsDiv = document.getElementById('query-results');
            resultsDiv.innerHTML = '';

            // Get values from form
            const databricksHost = document.getElementById('databricks-host').value;
            const warehouseId = document.getElementById('warehouse-id').value;
            const accessToken = document.getElementById('access-token').value;
            const sqlQuery = document.getElementById('sql-query').value;

            // Get destination options
            const destination = getDestinationOptions();

            // Validate inputs
            if (!databricksHost || !warehouseId || !accessToken || !sqlQuery) {
                showStatus('Please fill in all fields', false);
                return;
            }

            // Save host to localStorage
            localStorage.setItem('databricksHost', databricksHost);

            // Process any cell references in the SQL query
            const updateStatus = showStatus('Processing Excel cell references...', true, 'loading');
            const processedSqlQuery = await processSqlQueryWithCellReferences(sqlQuery);

            // Update the loading status message
            updateStatus('Running SQL query...');

            // Call the API function with the processed query
            const response = await queryDatabricks(warehouseId, accessToken, processedSqlQuery);

            if (response.error) {
                showStatus(`Error: ${response.error}`, false);
                return;
            }

            // Update status during data processing
            updateStatus('Preparing results display...');

            // Display results in the add-in
            displayQueryResults(response.data);

            // Write the data to Excel using the destination options
            await writeToExcel(response.data, destination);

            showStatus('Data successfully imported to Excel', true);
        } catch (error) {
            showStatus(`Error: ${error.message}`, false);
            console.error('Error:', error);
        }
    }

    // Helper function to get destination options from UI
    function getDestinationOptions(isGenieMode = false) {
        if (isGenieMode) {
            return {
                startCell: document.getElementById('genie-result-start-cell').value,
                newSheet: document.getElementById('genie-result-new-sheet').checked,
                appendData: document.getElementById('genie-result-append-data').checked
            };
        } else {
            return {
                startCell: document.getElementById('result-start-cell').value,
                newSheet: document.getElementById('result-new-sheet').checked,
                appendData: document.getElementById('result-append-data').checked
            };
        }
    }

    // Function to handle Genie result cell selection
    async function selectGenieResultCell() {
        try {
            showStatus('Click a cell in Excel to set as results destination', true, 'info');

            // Get the start cell input
            const cellInput = document.getElementById('genie-result-start-cell');

            // Use Excel API to get the selected cell
            await Excel.run(async (context) => {
                // Get the active worksheet first to get its exact name
                const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
                activeWorksheet.load("name");
                await context.sync();

                const exactSheetName = activeWorksheet.name;
                console.log("Active worksheet name:", exactSheetName);

                // This prompts the user to select a cell in Excel
                context.workbook.getActiveCell().select();

                // Get the selected range
                const selectedRange = context.workbook.getSelectedRange();
                selectedRange.load("address");
                await context.sync();

                // Extract just the address part (without sheet name)
                let cellAddress = selectedRange.address;
                if (cellAddress.includes('!')) {
                    cellAddress = cellAddress.split('!')[1];
                }

                // Update the input field with the selected cell
                cellInput.value = cellAddress;

                // Save the updated options
                saveGenieDestinationOptions();

                showStatus('Results destination cell set', true);
            });
        } catch (error) {
            showStatus(`Error selecting cell: ${error.message}`, false);
            console.error('Error selecting cell:', error);
        }
    }

    // Function to save the Genie destination options
    function saveGenieDestinationOptions() {
        const options = {
            startCell: document.getElementById('genie-result-start-cell').value,
            newSheet: document.getElementById('genie-result-new-sheet').checked,
            appendData: document.getElementById('genie-result-append-data').checked
        };
        localStorage.setItem('genieResultDestinationOptions', JSON.stringify(options));
    }

    async function queryDatabricks(warehouseId, accessToken, sqlQuery) {
        try {
            const databricksHost = document.getElementById('databricks-host').value;
            if (!databricksHost) {
                return { error: 'Databricks host is required' };
            }

            // Ensure the host doesn't end with a slash
            const baseUrl = databricksHost.endsWith('/')
                ? databricksHost.slice(0, -1)
                : databricksHost;

            // Use your existing Node.js server endpoint
            const proxyUrl = 'http://localhost:3000/query-databricks';

            showStatus('Sending request to Databricks via proxy server...', true, 'info');

            // Make the API call through your proxy server
            const response = await fetch(proxyUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    host: baseUrl,
                    warehouseId: warehouseId,
                    accessToken: accessToken,
                    sqlQuery: sqlQuery
                })
            });

            // Parse the response
            const responseData = await response.json();

            if (responseData.error) {
                console.error('API error:', responseData.error);
                return { error: responseData.error };
            }

            // Return the processed data
            return { data: responseData.data };
        } catch (error) {
            console.error('API call failed:', error);
            return { error: `Failed to connect to Databricks SQL Warehouse: ${error.message}` };
        }
    }

    // AI/BI Genie functionality
    async function askGenieQuestion() {
        try {
            // Reset any existing conversation
            currentConversationId = null;
            latestMessageId = null;

            // Clear the conversation history
            document.getElementById('conversation-history').innerHTML = '';
            document.getElementById('follow-up-container').classList.add('hidden');

            // Clear previous query results
            document.getElementById('query-results').innerHTML = '';

            // No longer clearing Excel at the start
            // Instead, we'll let the writeToExcel function handle sheet creation only when needed

            // Get values from form
            const databricksHost = document.getElementById('databricks-host').value;
            const genieSpaceId = document.getElementById('genie-space-id').value;
            const accessToken = document.getElementById('access-token').value;
            const question = document.getElementById('genie-question').value;

            // Validate inputs
            if (!databricksHost || !genieSpaceId || !accessToken || !question) {
                showStatus('Please fill in all fields', false);
                return;
            }

            // Save values to localStorage for convenience
            localStorage.setItem('databricksHost', databricksHost);
            localStorage.setItem('genieSpaceId', genieSpaceId);

            // Add user message to conversation
            addMessageToConversation(question, 'user');

            showStatus('Asking Genie...', true, 'info');

            // Start a new conversation with Genie
            const startResponse = await startGenieConversation(genieSpaceId, accessToken, question);

            if (startResponse.error) {
                showStatus(`Error: ${startResponse.error}`, false);
                return;
            }

            // Save the conversation ID and message ID
            currentConversationId = startResponse.conversation_id;
            latestMessageId = startResponse.message_id;

            // Poll for the message status
            await pollMessageStatus(genieSpaceId, accessToken, currentConversationId, latestMessageId);

            // Show the follow-up container once we have an active conversation
            document.getElementById('follow-up-container').classList.remove('hidden');

        } catch (error) {
            showStatus(`Error: ${error.message}`, false);
            console.error('Error:', error);
        }
    }

    async function sendFollowUpQuestion() {
        try {
            // Get values from form
            const genieSpaceId = document.getElementById('genie-space-id').value;
            const accessToken = document.getElementById('access-token').value;
            const question = document.getElementById('follow-up-question').value;

            // Validate inputs
            if (!genieSpaceId || !accessToken || !question || !currentConversationId) {
                showStatus('Missing required information for follow-up question', false);
                return;
            }

            // Add user message to conversation
            addMessageToConversation(question, 'user');

            // Clear the input field
            document.getElementById('follow-up-question').value = '';

            // Clear previous query results in the UI (not in Excel)
            document.getElementById('query-results').innerHTML = '';

            showStatus('Sending follow-up question...', true, 'info');

            // Send the follow-up message
            const messageResponse = await createGenieMessage(
                genieSpaceId,
                accessToken,
                currentConversationId,
                question
            );

            if (messageResponse.error) {
                showStatus(`Error: ${messageResponse.error}`, false);
                return;
            }

            // Save the message ID
            latestMessageId = messageResponse.message_id;

            // Poll for the message status
            await pollMessageStatus(genieSpaceId, accessToken, currentConversationId, latestMessageId);

        } catch (error) {
            showStatus(`Error: ${error.message}`, false);
            console.error('Error:', error);
        }
    }

    async function insertCellReference() {
        try {
            showStatus('Click a cell in Excel to insert its reference', true, 'info');

            // Get the SQL query textarea
            const textarea = document.getElementById('sql-query');
            const cursorPos = textarea.selectionStart;

            // Use Excel API to get the selected cell
            await Excel.run(async (context) => {
                // Get the active worksheet first to get its exact name
                const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
                activeWorksheet.load("name");
                await context.sync();

                const exactSheetName = activeWorksheet.name;
                console.log("Active worksheet name:", exactSheetName);

                // This prompts the user to select a cell in Excel
                context.workbook.getActiveCell().select();

                // Get the selected range
                const selectedRange = context.workbook.getSelectedRange();
                selectedRange.load("address");
                await context.sync();

                // Extract just the address part (without sheet name)
                let cellAddress = selectedRange.address;
                if (cellAddress.includes('!')) {
                    cellAddress = cellAddress.split('!')[1];
                }

                // Create the reference with exact sheet name
                const reference = `\${${exactSheetName}!${cellAddress}}`;
                console.log("Inserting reference:", reference);

                textarea.value =
                    textarea.value.substring(0, cursorPos) +
                    reference +
                    textarea.value.substring(cursorPos);

                // Update cursor position after the insertion
                textarea.selectionStart = cursorPos + reference.length;
                textarea.selectionEnd = cursorPos + reference.length;
                textarea.focus();

                showStatus('Cell reference inserted', true);
            });
        } catch (error) {
            showStatus(`Error inserting cell reference: ${error.message}`, false);
            console.error('Error inserting cell reference:', error);
        }
    }

    async function startGenieConversation(spaceId, accessToken, content) {
        try {
            const databricksHost = document.getElementById('databricks-host').value;
            if (!databricksHost) {
                return { error: 'Databricks host is required' };
            }

            // Ensure the host doesn't end with a slash
            const baseUrl = databricksHost.endsWith('/')
                ? databricksHost.slice(0, -1)
                : databricksHost;

            // Use the proxy server to handle the API call
            const proxyUrl = 'http://localhost:3000/genie/start-conversation';

            const response = await fetch(proxyUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    host: baseUrl,
                    spaceId: spaceId,
                    accessToken: accessToken,
                    content: content
                })
            });

            const data = await response.json();

            if (data.error) {
                return { error: data.error };
            }

            return {
                conversation_id: data.conversation_id,
                message_id: data.message_id
            };
        } catch (error) {
            console.error('Start conversation failed:', error);
            return { error: `Failed to start Genie conversation: ${error.message}` };
        }
    }

    async function createGenieMessage(spaceId, accessToken, conversationId, content) {
        try {
            const databricksHost = document.getElementById('databricks-host').value;
            if (!databricksHost) {
                return { error: 'Databricks host is required' };
            }

            // Ensure the host doesn't end with a slash
            const baseUrl = databricksHost.endsWith('/')
                ? databricksHost.slice(0, -1)
                : databricksHost;

            // Use the proxy server to handle the API call
            const proxyUrl = 'http://localhost:3000/genie/create-message';

            const response = await fetch(proxyUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    host: baseUrl,
                    spaceId: spaceId,
                    conversationId: conversationId,
                    accessToken: accessToken,
                    content: content
                })
            });

            const data = await response.json();

            if (data.error) {
                return { error: data.error };
            }

            return {
                message_id: data.message_id
            };
        } catch (error) {
            console.error('Create message failed:', error);
            return { error: `Failed to send message to Genie: ${error.message}` };
        }
    }

    async function getGenieMessage(spaceId, accessToken, conversationId, messageId) {
        try {
            const databricksHost = document.getElementById('databricks-host').value;
            if (!databricksHost) {
                return { error: 'Databricks host is required' };
            }

            // Ensure the host doesn't end with a slash
            const baseUrl = databricksHost.endsWith('/')
                ? databricksHost.slice(0, -1)
                : databricksHost;

            // Use the proxy server to handle the API call
            const proxyUrl = 'http://localhost:3000/genie/get-message';

            const response = await fetch(proxyUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    host: baseUrl,
                    spaceId: spaceId,
                    conversationId: conversationId,
                    messageId: messageId,
                    accessToken: accessToken
                })
            });

            const data = await response.json();

            if (data.error) {
                return { error: data.error };
            }

            return data;
        } catch (error) {
            console.error('Get message failed:', error);
            return { error: `Failed to get message from Genie: ${error.message}` };
        }
    }

    async function getQueryResult(spaceId, accessToken, conversationId, messageId, attachmentId) {
        try {
            const databricksHost = document.getElementById('databricks-host').value;
            if (!databricksHost) {
                return { error: 'Databricks host is required' };
            }

            // Ensure the host doesn't end with a slash
            const baseUrl = databricksHost.endsWith('/')
                ? databricksHost.slice(0, -1)
                : databricksHost;

            // Use the proxy server to handle the API call
            const proxyUrl = 'http://localhost:3000/genie/get-query-result';

            const response = await fetch(proxyUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    host: baseUrl,
                    spaceId: spaceId,
                    conversationId: conversationId,
                    messageId: messageId,
                    attachmentId: attachmentId,
                    accessToken: accessToken
                })
            });

            const data = await response.json();

            if (data.error) {
                return { error: data.error };
            }

            return { data: data };
        } catch (error) {
            console.error('Get query result failed:', error);
            return { error: `Failed to get query result from Genie: ${error.message}` };
        }
    }

    async function pollMessageStatus(spaceId, accessToken, conversationId, messageId) {
        const maxAttempts = 60; // 5 minutes at 5-second intervals
        let attempts = 0;

        // Create a loading indicator in the conversation
        const loadingId = 'loading-' + Date.now();
        const loadingHtml = `
            <div id="${loadingId}" class="message genie-message">
                <div class="loading"></div>
                <span class="loader-text">Genie is thinking...</span>
            </div>
        `;
        document.getElementById('conversation-history').insertAdjacentHTML('beforeend', loadingHtml);

        const poll = async () => {
            if (attempts >= maxAttempts) {
                // Remove the loading indicator
                const loadingElement = document.getElementById(loadingId);
                if (loadingElement) {
                    loadingElement.remove();
                }

                showStatus('Timed out waiting for Genie response', false);
                return;
            }

            attempts++;

            try {
                const messageResponse = await getGenieMessage(spaceId, accessToken, conversationId, messageId);

                if (messageResponse.error) {
                    // Remove the loading indicator
                    const loadingElement = document.getElementById(loadingId);
                    if (loadingElement) {
                        loadingElement.remove();
                    }

                    showStatus(`Error: ${messageResponse.error}`, false);
                    return;
                }

                // Check if the message processing is complete
                if (messageResponse.status === 'COMPLETED') {
                    // Remove the loading indicator
                    const loadingElement = document.getElementById(loadingId);
                    if (loadingElement) {
                        loadingElement.remove();
                    }

                    // Handle the completed message
                    await handleCompletedMessage(spaceId, accessToken, conversationId, messageId, messageResponse);

                    return;
                } else if (messageResponse.status === 'EXECUTING_QUERY') {
                    // Update the loading message
                    const loadingElement = document.getElementById(loadingId);
                    if (loadingElement) {
                        loadingElement.innerHTML = '<div class="loading"></div><span class="loader-text">Executing query...</span>';
                    }
                } else if (messageResponse.status === 'ERROR') {
                    // Remove the loading indicator
                    const loadingElement = document.getElementById(loadingId);
                    if (loadingElement) {
                        loadingElement.remove();
                    }

                    const errorMessage = messageResponse.error || 'An unknown error occurred';
                    showStatus(`Genie Error: ${errorMessage}`, false);
                    return;
                }

                // If not complete or error, wait 5 seconds and try again
                setTimeout(poll, 5000);
            } catch (error) {
                // Remove the loading indicator
                const loadingElement = document.getElementById(loadingId);
                if (loadingElement) {
                    loadingElement.remove();
                }

                console.error('Poll error:', error);
                showStatus(`Error polling for message status: ${error.message}`, false);
            }
        };

        // Start polling
        await poll();
    }

    async function handleCompletedMessage(spaceId, accessToken, conversationId, messageId, messageResponse) {
        try {
            // Check if we have a text-only response (no SQL query)
            if (messageResponse.attachments && messageResponse.attachments.length > 0 && messageResponse.attachments[0].text) {
                const textAttachment = messageResponse.attachments[0].text;
                const genieResponse = textAttachment.content || "Genie processed your question.";

                // Add the text response to the conversation
                addMessageToConversation(genieResponse, 'genie');

                // Clear any previous results display
                document.getElementById('query-results').innerHTML = '';

                showStatus('Genie response received', true);
                return;
            }

            // Check if there are SQL query attachments
            if (messageResponse.attachments && messageResponse.attachments.length > 0) {
                const attachment = messageResponse.attachments[0];

                // Prepare response text
                let genieResponse = messageResponse.content || "Genie processed your question.";
                const hasQuery = attachment.query && attachment.query.query;

                if (hasQuery) {
                    // Add SQL query information
                    genieResponse += `
                <div class="query-info">
                    <strong>Generated SQL:</strong>
                    <pre>${attachment.query.query}</pre>
                </div>`;
                }

                // Add the response to the conversation
                addMessageToConversation(genieResponse, 'genie');

                if (hasQuery && attachment.attachment_id) {
                    // Get the query results
                    showStatus('Fetching query results...', true, 'info');

                    const queryResult = await getQueryResult(
                        spaceId,
                        accessToken,
                        conversationId,
                        messageId,
                        attachment.attachment_id
                    );

                    if (queryResult.error) {
                        showStatus(`Error fetching results: ${queryResult.error}`, false);
                        return;
                    }

                    // Get destination options
                    const destination = getDestinationOptions(true);

                    // Check if we have valid data
                    if (Array.isArray(queryResult.data) && queryResult.data.length > 0) {
                        // Display the results and write to Excel using destination options
                        displayQueryResults(queryResult.data);
                        await writeToExcel(queryResult.data, destination);
                        showStatus('Data successfully imported to Excel', true);
                    } else {
                        // No results case
                        addMessageToConversation(
                            "The query executed successfully but returned no results or empty data.",
                            'genie'
                        );
                        showStatus('Query completed with no results', true, 'info');
                    }
                } else {
                    showStatus('Genie response received', true);
                }
            } else {
                // Just a text response with no query
                const genieResponse = messageResponse.content || "Genie processed your question, but no data was returned.";
                addMessageToConversation(genieResponse, 'genie');

                showStatus('Genie response received', true);
            }
        } catch (error) {
            console.error('Handle completed message error:', error);
            showStatus(`Error processing Genie response: ${error.message}`, false);
        }
    }

    // Function to select result cell (similar to insert cell reference)
    async function selectResultCell() {
        try {
            showStatus('Click a cell in Excel to set as results destination', true, 'info');

            // Get the start cell input
            const cellInput = document.getElementById('result-start-cell');

            // Use Excel API to get the selected cell
            await Excel.run(async (context) => {
                // Get the active worksheet first to get its exact name
                const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
                activeWorksheet.load("name");
                await context.sync();

                const exactSheetName = activeWorksheet.name;
                console.log("Active worksheet name:", exactSheetName);

                // This prompts the user to select a cell in Excel
                context.workbook.getActiveCell().select();

                // Get the selected range
                const selectedRange = context.workbook.getSelectedRange();
                selectedRange.load("address");
                await context.sync();

                // Extract just the address part (without sheet name)
                let cellAddress = selectedRange.address;
                if (cellAddress.includes('!')) {
                    cellAddress = cellAddress.split('!')[1];
                }

                // Update the input field with the selected cell
                cellInput.value = cellAddress;

                // Save the updated options
                saveDestinationOptions();

                showStatus('Results destination cell set', true);
            });
        } catch (error) {
            showStatus(`Error selecting cell: ${error.message}`, false);
            console.error('Error selecting cell:', error);
        }
    }
})();
