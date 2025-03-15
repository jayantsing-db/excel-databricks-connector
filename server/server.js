const express = require('express');
const cors = require('cors');
const axios = require('axios');
const app = express();

app.use(cors());
app.use(express.json());

// Original SQL query endpoint
app.post('/query-databricks', async (req, res) => {
    try {
        const { host, warehouseId, accessToken, sqlQuery } = req.body;

        // Validate inputs
        if (!host || !accessToken || !sqlQuery) {
            return res.status(400).json({ error: 'Missing required parameters' });
        }

        console.log(`Proxying request to: ${host}/api/2.0/sql/statements`);

        // Make request to Databricks SQL Warehouse API
        const response = await axios({
            method: 'POST',
            url: `${host}/api/2.0/sql/statements`,
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            data: {
                warehouse_id: warehouseId,
                statement: sqlQuery,
                format: "JSON"
            }
        });

        // Process the response
        if (response.data && response.data.manifest && response.data.result) {
            try {
                // Extract schema from the manifest
                const columns = response.data.manifest.schema.columns;

                // Get the data from the result
                const rows = response.data.result.data_array || [];

                // Transform the data into an array of objects
                const transformedData = rows.map(row => {
                    const rowObject = {};

                    // Map each column to its corresponding value
                    columns.forEach((column, index) => {
                        rowObject[column.name] = row[index];
                    });

                    return rowObject;
                });

                return res.json({ data: transformedData });
            } catch (error) {
                console.error('Error processing response:', error);
                return res.status(500).json({
                    error: 'Error processing response data',
                    details: error.message
                });
            }
        } else {
            return res.status(500).json({ error: 'Invalid response format from Databricks' });
        }
    } catch (error) {
        console.error('Databricks API error:', error);

        if (error.response) {
            return res.status(error.response.status).json({
                error: `Databricks API error: ${error.response.data.message || error.response.statusText}`
            });
        }

        return res.status(500).json({ error: `Failed to execute query: ${error.message}` });
    }
});

// Genie API endpoints
app.post('/genie/start-conversation', async (req, res) => {
    try {
        const { host, spaceId, accessToken, content } = req.body;

        // Validate inputs
        if (!host || !spaceId || !accessToken || !content) {
            return res.status(400).json({ error: 'Missing required parameters' });
        }

        console.log(`Starting Genie conversation in space: ${spaceId}`);

        // Make request to Databricks Genie API
        const response = await axios({
            method: 'POST',
            url: `${host}/api/2.0/genie/spaces/${spaceId}/start-conversation`,
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            data: {
                content: content
            }
        });

        // Return the conversation_id and message_id
        return res.json({
            conversation_id: response.data.conversation_id,
            message_id: response.data.message_id
        });
    } catch (error) {
        console.error('Genie API error:', error);

        if (error.response) {
            return res.status(error.response.status).json({
                error: `Genie API error: ${error.response.data.message || error.response.statusText}`
            });
        }

        return res.status(500).json({ error: `Failed to start conversation: ${error.message}` });
    }
});

app.post('/genie/create-message', async (req, res) => {
    try {
        const { host, spaceId, conversationId, accessToken, content } = req.body;

        // Validate inputs
        if (!host || !spaceId || !conversationId || !accessToken || !content) {
            return res.status(400).json({ error: 'Missing required parameters' });
        }

        console.log(`Creating message in conversation: ${conversationId}`);

        // Make request to Databricks Genie API
        const response = await axios({
            method: 'POST',
            url: `${host}/api/2.0/genie/spaces/${spaceId}/conversations/${conversationId}/messages`,
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            data: {
                content: content
            }
        });

        // Return the message_id
        return res.json({
            message_id: response.data.message_id
        });
    } catch (error) {
        console.error('Genie API error:', error);

        if (error.response) {
            return res.status(error.response.status).json({
                error: `Genie API error: ${error.response.data.message || error.response.statusText}`
            });
        }

        return res.status(500).json({ error: `Failed to create message: ${error.message}` });
    }
});

app.post('/genie/get-message', async (req, res) => {
    try {
        const { host, spaceId, conversationId, messageId, accessToken } = req.body;

        // Validate inputs
        if (!host || !spaceId || !conversationId || !messageId || !accessToken) {
            return res.status(400).json({ error: 'Missing required parameters' });
        }

        console.log(`Getting message: ${messageId}`);

        // Make request to Databricks Genie API
        const response = await axios({
            method: 'GET',
            url: `${host}/api/2.0/genie/spaces/${spaceId}/conversations/${conversationId}/messages/${messageId}`,
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });

        // Return the message data
        return res.json(response.data);
    } catch (error) {
        console.error('Genie API error:', error);

        if (error.response) {
            return res.status(error.response.status).json({
                error: `Genie API error: ${error.response.data.message || error.response.statusText}`
            });
        }

        return res.status(500).json({ error: `Failed to get message: ${error.message}` });
    }
});

app.post('/genie/get-query-result', async (req, res) => {
    try {
        const { host, spaceId, conversationId, messageId, attachmentId, accessToken } = req.body;

        // Validate inputs
        if (!host || !spaceId || !conversationId || !messageId || !attachmentId || !accessToken) {
            return res.status(400).json({ error: 'Missing required parameters' });
        }

        console.log(`Getting query result for attachment: ${attachmentId}`);

        // Make request to Databricks Genie API
        const response = await axios({
            method: 'GET',
            url: `${host}/api/2.0/genie/spaces/${spaceId}/conversations/${conversationId}/messages/${messageId}/attachments/${attachmentId}/query-result`,
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });

        // Process the specific Genie response format
        if (response.data && response.data.statement_response) {
            const statementResponse = response.data.statement_response;

            // Check if we have results and schema information
            if (statementResponse.result && statementResponse.manifest &&
                statementResponse.manifest.schema && statementResponse.manifest.schema.columns) {

                const columns = statementResponse.manifest.schema.columns;
                const rows = statementResponse.result.data_array || [];

                // Transform the data into an array of objects (similar to SQL endpoint format)
                const transformedData = rows.map(row => {
                    const rowObject = {};

                    // Map each column to its corresponding value
                    columns.forEach((column, index) => {
                        rowObject[column.name] = row[index];
                    });

                    return rowObject;
                });

                return res.json(transformedData);
            } else {
                return res.json([]);
            }
        } else {
            // If not in the expected format, return an empty array
            return res.json([]);
        }
    } catch (error) {
        console.error('Genie API error:', error);

        if (error.response) {
            return res.status(error.response.status).json({
                error: `Genie API error: ${error.response.data.message || error.response.statusText}`
            });
        }

        return res.status(500).json({ error: `Failed to get query result: ${error.message}` });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
