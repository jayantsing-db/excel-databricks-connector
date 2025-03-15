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

        if (isSuccess && type === 'info') {
            statusDiv.classList.add('status-info');
        } else if (isSuccess) {
            statusDiv.classList.add('status-success');
        } else {
            statusDiv.classList.add('status-error');
        }

        // Log to console for debugging
        if (isSuccess) {
            console.log('Status:', message);
        } else {
            console.error('Status Error:', message);
        }
    }

    function displayQueryResults(data) {
        const resultsDiv = document.getElementById('query-results');
        resultsDiv.innerHTML = '';

        if (!data || !data.length) {
            resultsDiv.innerHTML = '<p>No results found</p>';
            return;
        }

        // Create table
        const table = document.createElement('table');
        table.className = 'results-table';

        // Create header row
        const headerRow = document.createElement('tr');
        const headers = Object.keys(data[0]);

        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            headerRow.appendChild(th);
        });

        table.appendChild(headerRow);

        // Create data rows
        data.forEach(row => {
            const tr = document.createElement('tr');

            headers.forEach(header => {
                const td = document.createElement('td');

                // Handle different data types for display
                const value = row[header];
                if (value === null || value === undefined) {
                    td.textContent = '';
                } else if (typeof value === 'object' && value instanceof Date) {
                    td.textContent = value.toLocaleDateString();
                } else {
                    td.textContent = value.toString();
                }

                tr.appendChild(td);
            });

            table.appendChild(tr);
        });

        resultsDiv.appendChild(table);

        // Add row count
        const rowCount = document.createElement('p');
        rowCount.className = 'row-count';
        rowCount.textContent = `Rows: ${data.length}`;
        resultsDiv.appendChild(rowCount);
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

            // Set default host if saved in localStorage
            const savedHost = localStorage.getItem('databricksHost');
            if (savedHost) {
                document.getElementById('databricks-host').value = savedHost;
            }

            document.getElementById('status').textContent = "Databricks Excel Connector is ready!";
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

            // Save host to localStorage for convenience
            localStorage.setItem('databricksHost', databricksHost);

            showStatus('Running SQL query...', true, 'info');

            // Call the API function
            const response = await queryDatabricks(warehouseId, accessToken, sqlQuery);

            if (response.error) {
                showStatus(`Error: ${response.error}`, false);
                return;
            }

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
