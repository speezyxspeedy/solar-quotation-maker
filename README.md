# SHREE GANESH SOLAR ENERGY - Solar Quotation Generator

A complete 6-page solar quotation generator built with HTML, CSS, JavaScript, and FastAPI.

## Features

- Six dashboard tabs: Quotation Details, System Specification, Price Details, Terms & Conditions, Preview, and History
- Auto quotation numbers in `SGSE-2026-0001` style using browser `localStorage`
- Simple form fields for customer name, location, capacity, date, final price, panel brand, and inverter brand
- Editable specification rows, price rows, introduction, and terms
- Browser `localStorage` quotation history with View, Edit, Delete, PDF, Word, Excel, and WhatsApp actions
- FastAPI export endpoints for a fixed 6-page PDF, Word document, and Excel workbook
- No database and no login

## Run

Install dependencies:

```powershell
pip install -r requirements.txt
```

Start the backend:

```powershell
uvicorn main:app --reload
```

Open:

```text
http://127.0.0.1:8000/
```

## Backend Routes

- `GET /`
- `POST /generate-pdf`
- `POST /generate-word`
- `POST /generate-excel`
- `POST /save-quotation`
- `GET /history`

The browser is the V1 storage source. The `/save-quotation` and `/history` routes keep only an in-memory runtime copy for API compatibility.
# solar-quotation-maker
