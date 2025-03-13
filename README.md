# Excel-Databricks Connector

A Microsoft Excel add-in that allows users to connect directly to Databricks SQL warehouses, run queries, and import data into Excel spreadsheets.

## Features

- Connect to Databricks SQL warehouses using personal access tokens
- Execute SQL queries directly from Excel
- View query results in the add-in interface
- Import data directly into Excel worksheets
- Save connection details for quick access
- Auto-fit columns for better readability

## Prerequisites

- Microsoft Excel (Office 365 or Excel 2016+)
- Node.js and npm for development
- A Databricks workspace with SQL warehouse access

## Installation

### For Users

1. Download the latest release from the [Releases page](https://github.com/jayantsing-db/excel-databricks-connector/releases).
2. Open Excel and go to the "Insert" tab.
3. Click on "Office Add-ins".
4. Choose "Upload My Add-in" and select the downloaded manifest file.

### For Developers

1. Clone the repository:
   ```bash
   git clone https://github.com/jayantsing-db/excel-databricks-connector.git
   cd excel-databricks-connector
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Start the development server:
   ```bash
   npm start
   ```

4. Sideload the add-in in Excel:
    - Open Excel
    - Go to Insert > My Add-ins > Upload My Add-in
    - Browse to the project folder and select the manifest.xml file

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
├── certs/                # SSL certificates for development
├── server/               # Proxy server for Databricks API calls
│   ├── package.json      # Server dependencies
│   └── server.js         # Server implementation
├── package.json          # Project dependencies and scripts
└── README.md             # Project documentation
```

## Usage

1. Launch the add-in in Excel.
2. Enter your Databricks host URL (e.g., `https://dbc-xxxxxxxx-xxxx.cloud.databricks.com`).
3. Input your SQL warehouse ID.
4. Provide your personal access token for authentication.
5. Write your SQL query in the query text area.
6. Click "Run Query" to execute and import the data.

## Security Notes

- The add-in stores your Databricks host URL in localStorage for convenience, but does not store your access token.
- All queries are executed via a proxy server to avoid CORS issues.
- Make sure to keep your personal access tokens secure and do not share them.

## Development

### Starting the Development Server

```bash
npm start
```

This will start both the add-in server and the proxy server for Databricks API calls.

### Building for Production

```bash
npm run build
```

This creates a production-ready build in the `dist` folder.

## License

[MIT License](LICENSE)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Contact

Jayant Singh - [@jayantsing_db](https://github.com/jayantsing-db)

Project Link: [https://github.com/jayantsing-db/excel-databricks-connector](https://github.com/jayantsing-db/excel-databricks-connector)
