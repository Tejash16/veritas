# Veritas AI Auditor

Enterprise-grade AI presentation validation system that automatically validates Business PDF Presentations against source Excel spreadsheets.

## Features

- **AI-Powered Data Extraction**: Uses Google Gemini 2.5 Pro to extract and understand data from PDFs and Excel files
- **Intelligent Mapping**: Automatically suggests mappings between presentation values and source data
- **Interactive Validation**: User-friendly interface for confirming and editing data mappings
- **Comprehensive Auditing**: Context-aware validation with detailed discrepancy analysis
- **Multi-Format Reporting**: Interactive dashboards, annotated Excel files, and executive summaries
- **Enterprise Security**: JWT authentication, secure file storage, input validation
- **Usage Analytics**: Cost tracking, performance monitoring, error logging

## Technology Stack

**Backend:**
- FastAPI (Python)
- SQLAlchemy + SQLite/PostgreSQL
- Google Gemini 2.5 Pro for AI processing
- JWT authentication
- Prometheus metrics

**Frontend:**
- React.js 18
- Tailwind CSS
- Chart.js for visualizations
- Axios for API communication
- React Router for navigation

## Quick Start

### Prerequisites

- Python 3.9+
- Node.js 16+
- Google AI Studio API key

### Installation

1. Clone the repository
2. Run setup script: `scripts\setup.bat`
3. Configure environment variables
4. Start development servers: `scripts\run-dev.bat`

## Project Structure
## API Documentation

API documentation is available at `http://localhost:8000/docs` when running the development server.

## Security Features

- JWT-based authentication
- File upload validation and sanitization
- SQL injection prevention
- CORS configuration
- Rate limiting
- Secure file storage

## Monitoring & Metrics

- Prometheus metrics on port 8001
- Request latency tracking
- AI token usage monitoring
- Cost tracking per operation
- Error rate monitoring

## License

Enterprise License - Contact for commercial use
```
Veritas-Agent
├─ backend
│  ├─ app
│  │  ├─ api
│  │  │  ├─ audit.py
│  │  │  ├─ auth.py
│  │  │  ├─ reports.py
│  │  │  ├─ upload.py
│  │  │  └─ __init__.py
│  │  ├─ config.py
│  │  ├─ database
│  │  │  ├─ database.py
│  │  │  └─ __init__.py
│  │  ├─ database.py
│  │  ├─ main.py
│  │  ├─ models
│  │  │  ├─ audit.py
│  │  │  ├─ document.py
│  │  │  ├─ user.py
│  │  │  └─ __init__.py
│  │  ├─ services
│  │  │  ├─ ai_service.py
│  │  │  ├─ audit_service.py
│  │  │  ├─ enhanced_ai_service.py
│  │  │  ├─ enhanced_pdf_service.py
│  │  │  ├─ excel_service.py
│  │  │  ├─ mapping_service.py
│  │  │  ├─ pdf_service.py
│  │  │  ├─ report_service.py
│  │  │  └─ __init__.py
│  │  ├─ utils
│  │  │  ├─ file_handler.py
│  │  │  ├─ metrics.py
│  │  │  ├─ security.py
│  │  │  └─ __init__.py
│  │  └─ __init__.py
│  ├─ Dockerfile
│  ├─ requirements.txt
│  ├─ temp_reports
│  └─ test_env.py
├─ docker-compose.yml
├─ docs
│  ├─ API_DOCUMENTATION.md
│  ├─ SYSTEM_OVERVIEW.md
│  ├─ TROUBLESHOOTING.md
│  └─ USER_GUIDE.md
├─ frontend
│  ├─ package-lock.json
│  ├─ package.json
│  ├─ postcss.config.js
│  ├─ src
│  │  ├─ App.css
│  │  ├─ App.jsx
│  │  ├─ components
│  │  │  ├─ Common
│  │  │  │  ├─ ErrorBoundary.jsx
│  │  │  │  ├─ Header.jsx
│  │  │  │  └─ LoadingSpinner.jsx
│  │  │  ├─ Dashboard
│  │  │  │  ├─ AuditDashboard.jsx
│  │  │  │  ├─ DiscrepancyList.jsx
│  │  │  │  └─ MetricsPanel.jsx
│  │  │  ├─ Preview
│  │  │  │  └─ EnhancedDocumentPreview.jsx
│  │  │  ├─ Reports
│  │  │  │  ├─ ReportGenerator.jsx
│  │  │  │  └─ ReportViewer.jsx
│  │  │  ├─ Upload
│  │  │  │  ├─ EnhancedMappingValidation.jsx
│  │  │  │  ├─ FileUpload.jsx
│  │  │  │  └─ MappingConfirmation.jsx
│  │  │  └─ Viewer
│  │  │     ├─ DocumentViewer.jsx
│  │  │     ├─ ExcelViewer.jsx
│  │  │     └─ PDFViewer.jsx
│  │  ├─ hooks
│  │  │  ├─ useAudit.js
│  │  │  └─ useAuth.js
│  │  ├─ index.css
│  │  ├─ index.js
│  │  ├─ services
│  │  │  ├─ api.js
│  │  │  ├─ auth.js
│  │  │  └─ fileService.js
│  │  └─ utils
│  │     ├─ constants.js
│  │     └─ helpers.js
│  └─ tailwind.config.js
├─ README.md
└─ scripts
   ├─ build.bat
   ├─ run-dev.bat
   └─ setup.bat

```