const express = require('express');
const cors = require('cors');
const axios = require('axios');
const app = express();

app.use(cors());
app.use(express.json());

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
                format: "JSON"  // Match the format in your Python example
            }
        });

        // Process the response based on your Python script output format
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
