# BrowserStack Testathon - Automated Testing Workflow

An end-to-end, production-ready, and high-performance automated testing framework integrating **Python**, **Selenium WebDriver**, **BrowserStack Automate**, and **PostgreSQL**.

The framework supports both public web application testing and local intranet/API endpoints via BrowserStack Local tunneling, featuring automated synchronization with PostgreSQL and rich dual reporting (**Excel (.xlsx)** with openpyxl styling and **PDF** with ReportLab).

---

## 1. Directory Structure

```text
browserstacktestthon/
├── .env.example              # Template for environment variables and secrets
├── .gitignore                # Excludes virtual env, credentials, and temp report outputs
├── requirements.txt          # Production dependencies
├── setup.bat                 # Automated Windows setup script (creates env, installs pkgs)
├── test_runner.py            # Complete modular test runner
├── reports/                  # Generated reports directory (Excel & PDF)
│   ├── .gitkeep
│   ├── execution_report_<timestamp>.xlsx
│   └── execution_report_<timestamp>.pdf
└── README.md                 # Complete documentation & usage guide
```

---

## 2. Technical Stack & Dependencies

- **Language & Runtime**: Python 3.10+ (tested on Python 3.14)
- **Virtual Environment**: Isolated project environment `./env`
- **Testing & Automation**:
  - `selenium>=4.25.0`: Remote WebDriver targeting BrowserStack Cloud Hub with W3C `bstack:options`.
  - `requests>=2.31.0`: High-performance HTTP client for API / Endpoint verification.
- **Database Synchronization**:
  - `psycopg2-binary>=2.9.9`: PostgreSQL driver with connection pooling and parameterized transactions.
- **Reporting**:
  - `openpyxl>=3.1.2` & `pandas>=2.2.0`: Styled Excel reports with custom cell fills, borders, and auto-adjusted column dimensions.
  - `reportlab>=4.2.0`: PDF generation with executive KPI metrics cards, execution tables, and payload verification snippets.
- **Security & Configuration**:
  - `python-dotenv>=1.0.1`: Zero hardcoded secrets, strictly environment-driven.

---

## 3. Quick Start & Setup

### Automated Setup (Windows)

Run the included `setup.bat` script. It automatically:
1. Validates Python in your system PATH.
2. Creates the `./env` isolated virtual environment via `python -m venv env`.
3. Activates the environment.
4. Upgrades `pip` and installs all dependencies from `requirements.txt`.
5. Initializes `.env` from `.env.example` if not already present.

```cmd
setup.bat
```

### Manual Setup (CLI / PowerShell)

```powershell
# 1. Create isolated virtual environment
python -m venv env

# 2. Activate environment
.\env\Scripts\activate

# 3. Upgrade pip and install dependencies
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

# 4. Create your .env file
copy .env.example .env
```

---

## 4. Configuration (`.env`)

Configure your `.env` file with your credentials:

```ini
# BrowserStack Automate Credentials
BROWSERSTACK_USERNAME=your_browserstack_username
BROWSERSTACK_ACCESS_KEY=your_browserstack_access_key

# Optional: BrowserStack Local Tunnel Identifier
BROWSERSTACK_LOCAL_IDENTIFIER=local_testathon_tunnel

# PostgreSQL Database
DB_HOST=localhost
DB_PORT=5432
DB_NAME=testathon_db
DB_USER=postgres
DB_PASSWORD=your_db_password

# API Endpoint Testing (Scenario 2)
API_TEST_ENDPOINT=https://httpbin.org/post
API_TEST_USERNAME=testathon_user
API_TEST_PASSWORD=SecurePassword123!
API_TEST_ROLL_NUMBER=BST-2026-9042
```

---

## 5. Execution Guide

Run the suite using the virtual environment:

```powershell
.\env\Scripts\python.exe test_runner.py
```

### Advanced CLI Flags

| Flag | Description |
| :--- | :--- |
| `--all` | Runs the full comprehensive matrix: Public UI, Local Tunnel UI, Public API, and Local Intranet API. |
| `--local` | Enables BrowserStack Local tunneling capability (`bstack:options -> local: true`). |
| `--api-url <URL>` | Overrides the target API endpoint for Scenario 2. |
| `--output-dir <DIR>` | Customizes the report output directory (defaults to `./reports`). |
| `--skip-ui` | Runs API endpoint scenario only, skipping UI web navigation. |
| `--skip-api` | Runs UI scenario only, skipping API endpoint testing. |
| `--skip-db` | Skips PostgreSQL synchronization. |
| `--dry-run` | Executes UI tests in simulated mode (ideal for CI or offline verification without cloud credentials). |

#### Examples

```powershell
# Run with BrowserStack Local tunnel enabled
.\env\Scripts\python.exe test_runner.py --local

# Test a custom API endpoint
.\env\Scripts\python.exe test_runner.py --api-url https://api.mycompany.internal/verify

# Run in dry-run mode (tests pipeline, API, DB fallback, and Excel/PDF generation)
.\env\Scripts\python.exe test_runner.py --dry-run
```

---

## 6. Architecture & Core Modules

`test_runner.py` is structured into 4 clean, production-grade modular components:

### 1. `DatabaseManager`
- Manages PostgreSQL connection pooling (`SimpleConnectionPool`).
- Automatically verifies and provisions the `test_executions` table schema:
  ```sql
  CREATE TABLE IF NOT EXISTS test_executions (
      id SERIAL PRIMARY KEY,
      test_name VARCHAR(255) NOT NULL,
      target_url TEXT NOT NULL,
      execution_status VARCHAR(50) NOT NULL,
      latency_ms FLOAT NOT NULL,
      response_summary JSONB,
      executed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
  );
  ```
- Implements safe parameterized single & bulk `INSERT` statements wrapped in `try ... finally` blocks.
- Seamlessly falls back to in-memory logging if PostgreSQL is unreachable, allowing test runs and report generation to complete without crashing.

### 2. `BrowserStackDriverFactory`
- Instantiates W3C-compliant Remote Selenium WebDriver targeting `https://<USERNAME>:<ACCESS_KEY>@hub-cloud.browserstack.com/wd/hub`.
- Configures `bstack:options` capabilities:
  - OS: Windows 11
  - Browser: Chrome latest
  - Session Name: "BrowserStack Testathon Suite"
  - Build Name: "Build_v1.0"
  - Dynamic Local Tunneling: `local = True/False` + `localIdentifier`.
- Updates BrowserStack session status via JavaScript executor:
  `browserstack_executor: {"action": "setSessionStatus", "arguments": {"status": "passed"/"failed", "reason": ...}}`.
- Guarantees clean `driver.quit()` teardown.

### 3. `TestSuite`
- **Scenario 1 (UI Web Navigation)**: Navigates to `https://www.google.com`, measures page load latency, asserts page title retrieval and `document.readyState == "complete"`, and updates remote session status.
- **Scenario 2 (API Endpoint Verification)**: Dispatches HTTP POST/GET passing `username`, `password`, and `roll_number`, measures millisecond roundtrip latency, and validates HTTP status codes and JSON schema integrity.

### 4. `ReportGenerator`
- **Excel Report (`.xlsx`)**:
  - Styled header with navy blue fill (`#1F4E78`) and white bold text.
  - Soft green / red status cell highlights (`#D4EDDA` / `#F8D7DA`).
  - Auto-adjusted column widths and numeric latency formatting.
- **PDF Executive Report (`.pdf`)**:
  - Executive KPI summary card block (Total, Passed, Failed, Pass Rate %, Avg Latency).
  - Clean execution table with status badges.
  - Formatted JSON response snippets for full auditability.

---

## 7. License & Attribution

Built for the BrowserStack Testathon. Zero placeholders, 100% production-ready.
