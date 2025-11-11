# Dynamic API Test Generator

This project dynamically generates Postman collections, runs Newman, and creates Excel reports for multiple services using:

- Controller-based YAML test definitions
- Swagger API documentation
- Dynamic paths per service

---

## Project Structure

```
.
├── src/                           # Source code modules
│   ├── __init__.py               # Package initialization
│   ├── main.py                   # Main entry point
│   ├── swagger_fetcher.py        # Fetch Swagger documentation
│   ├── test_data_loader.py       # Load test data from APIs/files
│   ├── postman_generator.py      # Generate Postman collections
│   ├── excel_generator.py        # Generate Excel reports
│   └── newman_runner.py          # Run Newman CLI tool
├── run.py                        # Root-level entry point (recommended)
├── generate_tests.py             # Legacy monolithic entry point (kept for reference)
├── services_config.yaml          # Service configuration
├── requirements.txt              # Python dependencies
└── README.md                     # This file
```

## Setup

1. Install Python dependencies:

```bash
pip install -r requirements.txt
```

2. Configure services in `services_config.yaml`

## Usage

### Option 1: Run from root (Recommended)

```bash
python run.py
```

### Option 2: Run from src directory

```bash
python src/main.py
```

### Option 3: Run as module

```bash
python -m src.main
```

## Modules

- **swagger_fetcher.py**: Fetches Swagger/OpenAPI documentation from API endpoints
- **test_data_loader.py**: Loads test data from API endpoints or local files
- **postman_generator.py**: Generates Postman collection JSON from Swagger and test data
- **excel_generator.py**: Creates Excel reports with test results and summaries
- **newman_runner.py**: Executes Newman CLI to run Postman collections
- **main.py**: Orchestrates the entire workflow

## Output

The tool generates:
- `{service_name}_postman_collection.json` - Postman collection file
- `{service_name}_swagger_testdata_summary.xlsx` - API endpoints summary
- `newman_report.html` - Newman test execution report
