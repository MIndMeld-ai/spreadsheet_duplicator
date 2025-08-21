Prototype backend: simple server to accept a template and mapping JSON and return generated workbooks using openpyxl.

This prototype explains how to run the Python server locally.

Requirements
- Python 3.8+
- pip

Install

```bash
python -m venv .venv
source .venv/bin/activate
pip install fastapi uvicorn python-multipart openpyxl
```

Run server

```bash
uvicorn server:app --reload --port 8000
```

Server endpoints
- POST /generate : multipart form-data. Fields:
  - template: file (.xlsx)
  - mapping: JSON string describing rows and target mappings
  - options: optional JSON string for naming/patterns

Response: application/zip or application/vnd.openxmlformats-officedocument.spreadsheetml.sheet

Notes
- This is a local prototype only. To host publicly, deploy to any server that supports Python (e.g. DigitalOcean, AWS EC2, Heroku-like, Railway). See below for quick deployment ideas.
