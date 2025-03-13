(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.onReady(function (info) {
        console.log("Office.onReady called", info);
        if (info.host === Office.HostType.Excel) {
            console.log("Excel detected!");
            document.getElementById('run-query').onclick = runQuery;
            document.getElementById('status').textContent = "Databricks Add-in is ready!";
            document.getElementById('status').className = "status-message status-success";

            // Set default host if saved in localStorage
            const savedHost = localStorage.getItem('databricksHost');
            if (savedHost) {
                document.getElementById('databricks-host').value = savedHost;
            }
        } else {
            console.log("Not running in Excel", info);
            document.getElementById('status').textContent = "Not running in Excel. Host: " + (info.host || "unknown");
            document.getElementById('status').className = "status-message status-error";
        }
    });

    async function runQuery() {
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

            showStatus('Running query...', true);

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

            showStatus('Sending request to Databricks via proxy server...', true);

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

    function processResponseData(responseData) {
        // Check if we have a valid response
        if (!responseData || !responseData.manifest || !responseData.result) {
            return { error: 'Invalid response format from Databricks' };
        }

        try {
            // Extract schema from the manifest
            const columns = responseData.manifest.schema.columns;

            // Get the data from the result
            const rows = responseData.result.data_array || [];

            // Transform the data into an array of objects
            const transformedData = rows.map(row => {
                const rowObject = {};

                // Map each column to its corresponding value
                columns.forEach((column, index) => {
                    rowObject[column.name] = row[index];
                });

                return rowObject;
            });

            return { data: transformedData };
        } catch (error) {
            console.error('Error processing response data:', error);
            return { error: `Error processing response data: ${error.message}` };
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

    function showStatus(message, isSuccess) {
        const statusDiv = document.getElementById('status');
        statusDiv.textContent = message;
        statusDiv.className = 'status-message';

        if (isSuccess) {
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
})();