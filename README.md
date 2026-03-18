# Demand Forecasting Deep Dive

A Streamlit-based application for analyzing and visualizing demand forecasting data with interactive charts and timeline analysis.

## Features

- Interactive data visualization using Altair
- Quarter-based timeline analysis
- Demand forecasting analysis
- Wide layout for comprehensive data exploration

## Requirements

- Python 3.7+
- See `requirements.txt` for dependencies

## Installation

1. Clone or download this project
2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Running the Project

### Using Streamlit (Recommended)
```bash
streamlit run analysis.py
```

The application will open in your default browser at `http://localhost:8501`

### Command-line Options

Run with specific settings:
```bash
# Run with custom port
streamlit run analysis.py --server.port 8502

# Run in development mode
streamlit run analysis.py --logger.level=debug

# Run without browser auto-open
streamlit run analysis.py --client.showErrorDetails=true --logger.level=debug
```

## Project Structure

```
├── analysis.py              # Main Streamlit application
├── requirements.txt         # Python dependencies
├── README.md               # This file
└── CFL_External Data Pack_Phase1.xlsx  # Data source
```

## Usage

1. Start the Streamlit server using the command above
2. The application will load the data from the Excel file
3. Navigate through the interactive visualizations
4. Use the sidebar and controls to filter and analyze data

## Data Source

The application reads from: `D:\projects and files\CFL\CFL_External Data Pack_Phase1.xlsx`

## Troubleshooting

- **Port already in use**: Use the `--server.port` flag to specify a different port
- **File not found**: Ensure the Excel data file exists at the specified path
- **Module not found**: Run `pip install -r requirements.txt` again to ensure all dependencies are installed
