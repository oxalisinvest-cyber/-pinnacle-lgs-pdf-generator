# Pinnacle LGS PDF & Excel Generator

Web service that generates proforma invoices (PDF) and corresponding spreadsheets (XLSX) for Pinnacle LGS.

## Endpoints

- `GET /` — service info
- `GET /health` — health check
- `POST /generate-pdf` — returns PDF as base64
- `POST /generate-excel` — returns Excel as base64
- `POST /generate-both` — returns both in one call

## Deployment

Deployed on Render.com as a Python web service.
