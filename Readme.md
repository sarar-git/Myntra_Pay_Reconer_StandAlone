# Myntra Payment Register

Generate a Payment Register from the Myntra Revamped Payment Report.

## Features

- Upload Myntra Payment Report
- Merge Prepaid & Postpaid settlements
- Group by Settlement Date + UTR
- Sum payment amounts
- Download Payment Register in Excel format

## Installation

```bash
pip install -r requirements.txt
```

Run

```bash
uvicorn app:app --reload
```

Open

```
http://localhost:8000/docs
```
