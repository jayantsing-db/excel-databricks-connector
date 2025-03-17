# Excel-Databricks Connector

A Microsoft Excel add-in that enables seamless integration between Excel and Databricks, allowing users to connect directly to Databricks SQL warehouses, run queries, and import data into Excel spreadsheets with just a few clicks.

![Databricks Logo](./app/assets/icon-80.png)

## Features

- **SQL Query Integration**: Execute SQL queries directly from Excel and view results immediately
- **AI/BI Genie Mode**: Use natural language to query your data with Databricks AI/BI Genie
- **Excel Cell References**: Reference Excel cells directly in your SQL queries with `${A1}` syntax
- **Flexible Data Import**: Import data directly into Excel worksheets with options for:
   - Creating new sheets
   - Specifying start cells
   - Appending to existing data
- **Connection Management**: Save connection details for quick access to your Databricks environment
- **Enhanced Data Viewing**: Results panel with pagination and adjustable rows per page
- **Auto-formatting**: Auto-fit columns for better readability in Excel

## Prerequisites

- Microsoft Excel (Office 365 or Excel 2016+)
- A Databricks workspace with SQL warehouse access
- Node.js and npm for development

## Installation

### For Users

1. Download the latest release from the [Releases page](https://github.com/jayantsing-db/excel-databricks-connector/releases)
2. Open Excel and go to the "Insert" tab
3. Click on "Office Add-ins"
4. Choose "Upload My Add-in" and select the downloaded manifest file

### For Developers

1. Clone the repository:
   ```bash
   git clone https://github.com/jayantsing-db/excel-databricks-connector.git
   cd excel-databricks-connector
   ```

2. Install app dependencies:
   ```bash
   cd app
   npm install
   ```

3. Install server dependencies:
   ```bash
   cd ../server
   npm install
   ```

4. Generate SSL certificates for development:
   ```bash
   cd ../app
   npx office-addin-dev-certs install
   ```

5. Start the development servers:
   ```bash
   # Start the frontend server in the app directory
   npm start
   
   # Start the proxy server in the server directory (in a separate terminal)
   cd ../server
   npm start
   ```

6. Sideload the add-in in Excel:
   - Open Excel
   - Go to Insert > My Add-ins > Upload My Add-in
   - Browse to the project folder and select the manifest.xml file

## Usage

### SQL Mode

1. Launch the add-in in Excel
2. Enter your Databricks host URL (e.g., `https://dbc-xxxxxxxx-xxxx.cloud.databricks.com`)
3. Enter your personal access token for authentication
4. Provide your SQL warehouse ID
5. Write your SQL query in the query text area
   - You can use Excel cell references in your query with `${A1}` or `${Sheet1!B2}` syntax
6. Set your preferred results destination (optional):
   - Choose a specific start cell
   - Create a new sheet for results
   - Append to existing data
7. Click "Run SQL Query" to execute and import the data

### AI/BI Genie Mode

1. Switch to "AI/BI Genie" tab
2. Enter your Databricks host URL and personal access token
3. Enter your Genie Space ID (found in your Genie Space URL)
4. Ask a question about your data in natural language
5. View the generated SQL and results
6. Ask follow-up questions to refine your analysis

## Security Notes

- Your Databricks host URL is stored in localStorage for convenience
- Personal access tokens are never stored and must be re-entered for each session
- All queries are executed via a proxy server to avoid CORS issues
- Keep your personal access tokens secure and do not share them

## Project Structure

```
excel-databricks-connector/
├── app/                  # Main add-in interface files
│   ├── app.css           # Styles for the add-in
│   ├── app.js            # Main application logic
│   ├── assets/           # Static assets
│   ├── function-file.html # Function file for add-in commands
│   ├── index.html        # Main HTML interface
│   └── manifest.xml      # Add-in manifest
├── server/               # Proxy server for Databricks API calls
│   ├── package.json      # Server dependencies
│   └── server.js         # Server implementation
├── package.json          # Project dependencies and scripts
└── README.md             # Project documentation
```

## Developing and Extending

### Adding New Features

1. Frontend changes are made in the `app` directory
2. API proxy and backend logic is in the `server` directory
3. Update the manifest.xml file if you're changing permissions or add-in metadata

### Building for Production

```bash
cd app
npm run build
```

This creates a production-ready build in the `dist` folder.

## Troubleshooting

- **CORS Issues**: If you encounter CORS errors, make sure the proxy server is running
- **Certificate Errors**: For development, ensure you've installed the dev certificates
- **Authentication Errors**: Verify your personal access token has the correct permissions

## License

[MIT License](LICENSE)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request or open an Issue.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Acknowledgments

- This project is not officially affiliated with Databricks
- Inspired by the need for easier data access between Excel and Databricks

## Contact

Jayant Singh - [@jayantsing_db](https://github.com/jayantsing-db)

Project Link: [https://github.com/jayantsing-db/excel-databricks-connector](https://github.com/jayantsing-db/excel-databricks-connector)
