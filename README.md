# Florastyle Stock Update Web Application

A simple Python web application that fetches stock data from the IQ Retail API and displays it in a table format with Excel export functionality.

## Features

- 📊 Display stock data in a clean, modern web interface
- 📥 Export data to Excel (.xlsx) format
- 🔄 Refresh data on demand
- 🎨 Beautiful, responsive UI design

## Required Fields Displayed

Per product, the application displays:
- **Description**
- **Supplier_Item_Code**
- **Onhand_Available**

## Setup Instructions

### 1. Install Python Dependencies

```bash
pip install -r requirements.txt
```

### 2. Run the Application

```bash
python app.py
```

The application will start on `http://localhost:5000`

### 3. Access the Web Interface

Open your web browser and navigate to:
```
http://localhost:5000
```

## Usage

1. **View Stock Data**: The main page automatically fetches and displays stock data from the API
2. **Refresh Data**: Click the "Refresh Data" button to reload the latest stock information
3. **Export to Excel**: Click the "Export to Excel" button to download the data as an .xlsx file

## API Configuration

The API configuration is set in `app.py`:
- **Endpoint**: `http://192.168.5.90:8090/IQRetailRestAPI/v1/IQ_API_Request_Stock_Attributes`
- **Method**: POST
- **Format**: XML

Note: Make sure the API server is accessible from your network before running the application.

## File Structure

```
Florastyle_Stock_Update/
├── app.py              # Main Flask application
├── requirements.txt    # Python dependencies
├── templates/
│   └── index.html     # Web interface template
└── README.md          # This file
```

## Requirements

- Python 3.7+
- Flask 3.0.0
- requests 2.31.0
- pandas 2.1.3
- openpyxl 3.1.2

## Working?