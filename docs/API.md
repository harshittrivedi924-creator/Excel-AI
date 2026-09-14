# REST API Documentation

Base URL: `http://localhost:5000`

## Endpoints

### GET /api/workbook

Get information about the currently loaded workbook.

**Response:**
```json
{
  "loaded": true,
  "sheets": ["Sheet1", "Sales"],
  "active_sheet": "Sheet1",
  "rows": [
    ["Month", "Revenue", "Expense"],
    ["January", "10000", "8000"],
    ["February", "12000", "9000"]
  ],
  "file": "sales.xlsx"
}
```

**Not loaded:**
```json
{
  "loaded": false,
  "sheets": [],
  "rows": [],
  "file": null
}
```

### POST /api/upload

Upload an Excel workbook.

**Request:**
- Content-Type: `multipart/form-data`
- Body: `file` field with .xlsx/.xlsm/.xls file

**Response:** Same as `GET /api/workbook`

**Error:**
```json
{
  "loaded": false,
  "error": "Couldn't open file: ..."
}
```

### POST /api/command

Execute a natural language command.

**Request:**
```json
{
  "command": "Column D ka sum karo",
  "sheet": "Sheet1",
  "allow_overwrite": false
}
```

**Response (success):**
```json
{
  "success": true,
  "message": "Done. The sum of D2:D20 is 54000.",
  "result": 54000,
  "overwrite_needed": false
}
```

**Response (overwrite needed):**
```json
{
  "success": false,
  "message": "D21 already contains data. Replace it?",
  "result": null,
  "overwrite_needed": true
}
```

**Response (error):**
```json
{
  "success": false,
  "message": "I couldn't find Column X in this sheet.",
  "result": null,
  "overwrite_needed": false
}
```

### GET /api/history

Get the command history.

**Response:**
```json
[
  {
    "operation": "SUM",
    "result": 54000,
    "destination": "D21"
  },
  {
    "operation": "AVERAGE",
    "result": 3000,
    "destination": null
  }
]
```

## Error Handling

All endpoints return appropriate HTTP status codes:
- `200` - Success
- `400` - Bad request (missing file, invalid command)
- `500` - Server error

## CORS

For development, CORS is not configured. The API expects to be served from the same origin as the frontend.
