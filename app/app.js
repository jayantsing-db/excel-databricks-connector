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

    async function writeToExcel(data) {
        return new Promise((resolve, reject) => {
            try {
                Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getActiveWorksheet();

                    // Clear any existing data
                    sheet.getUsedRange().clear();

                    if (!data || !data.length) {
                        await context.sync();
                        resolve();
                        return;
                    }

                    const headers = Object.keys(data[0]);
                    const headerRange = sheet.getRange("A1").getResizedRange(0, headers.length - 1);
                    headerRange.values = [headers];
                    headerRange.format.font.bold = true;

                    // Prepare data for Excel
                    const rows = data.map(row =>
                        headers.map(header => {
                            const value = row[header];
                            if (value === null || value === undefined) return '';
                            return value;
                        })
                    );

                    if (rows.length > 0) {
                        const dataRange = sheet.getRange("A2").getResizedRange(rows.length - 1, headers.length - 1);
                        dataRange.values = rows;
                    }

                    // Auto-fit columns
                    sheet.getUsedRange().format.autofitColumns();

                    await context.sync();
                    resolve();
                });
            } catch (error) {
                reject(error);
            }
        });
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
            document.getElementById('prev-page').onclick = () => {}; // Will be overridden in displayQueryResults
            document.getElementById('next-page').onclick = () => {}; // Will be overridden in displayQueryResults
            document.getElementById('rows-per-page').onchange = () => {}; // Will be overridden in displayQueryResults

            // Set default host if saved in localStorage
            const savedHost = localStorage.getItem('databricksHost');
            if (savedHost) {
                document.getElementById('databricks-host').value = savedHost;
            }

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

            // Write the data to Excel
            await writeToExcel(response.data);

            showStatus('Data successfully imported to Excel', true);
        } catch (error) {
            showStatus(`Error: ${error.message}`, false);
            console.error('Error:', error);
        }
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

            // Clear Excel sheet when starting a new conversation
            try {
                await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getActiveWorksheet();
                    sheet.getUsedRange().clear();
                    await context.sync();
                });
            } catch (excelError) {
                console.error('Excel clear error:', excelError);
                // Continue even if Excel clear fails
            }

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

            // Clear previous query results
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

                    // Check if we have valid data
                    if (Array.isArray(queryResult.data) && queryResult.data.length > 0) {
                        // Display the results and write to Excel
                        displayQueryResults(queryResult.data);
                        await writeToExcel(queryResult.data);
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
})();
