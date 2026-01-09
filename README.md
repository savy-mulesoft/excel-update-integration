# RBC Excel API Documentation

> ⚠️ **Disclaimer**
>
> This API and the accompanying documentation are **not production-grade**.  
> They are provided **solely for development guidance and proof-of-concept (PoC) purposes**.
>
> Security, scalability, performance hardening, error handling, and operational controls
> have **not** been fully implemented and **must be addressed before any production use**.


## Overview

The RBC Excel API is a RESTful service designed to facilitate the modification of Excel files (`.xlsm` format). It allows users to update specific cells in an Excel file with new data, either by using a server-side template or by uploading their own Excel file.

### Key Features

* **Cell Updates:** Update specific cells in an Excel file using JSON payloads.
* **Flexible Output:** Save the updated file on the server or download it directly as a response.
* **Custom Uploads:** Support for uploading custom Excel files for modification.
* **Input Formats:** Supports JSON and multipart form data.

---

## Environments & Base URLs

| Environment | URL |
|:------------|:----|
| **Development** | 
| **Local** | `http://localhost:8081` |

---

## Quick Start

To test the API immediately using the development environment, run the following cURL command to update a file and download the result:

```bash
curl --location 'https://rbc-excel-proxy-app-fomag7.5sc6y6-1.usa-e2.cloudhub.io/update-download' \
--header 'Content-Type: application/json' \
--data '[
    {
        "key": "A1",
        "value": "Test Value"
    },
    {
        "key": "B2",
        "value": 123
    }
]' --output updated_file.xlsm
```

---

## Endpoint Reference

### 1. Update & Save

**POST** `/update-save`

Updates cells in the server's template Excel file and saves the result to the server's output directory.

- **Use Case:** Storing the updated file on the server for later access.
- **Server Storage:** Yes.

#### Request Body

A JSON array of objects containing key (cell reference) and value.

```json
[
  { "key": "A1", "value": "Hello World" },
  { "key": "B2", "value": 12345 }
]
```

#### Response (200 OK)

```json
{
  "status": "success",
  "message": "Excel file updated successfully",
  "requestId": "f8195c9e-7060-43c8-aeee-4ce25f2d9cde",
  "fileName": "updated-excel-20251204-203059-f8195c9e.xlsm",
  "updatedCells": 3,
  "timestamp": "20251204-203059"
}
```

### 2. Update & Download

**POST** `/update-download`

Updates cells in the server's template Excel file and returns the modified file as a direct download.

- **Use Case:** Immediate retrieval without server-side storage.
- **Server Storage:** No (In-memory processing).

#### Request Body

Same JSON array format as `/update-save`.

#### Response (200 OK)

- **Body:** Binary Excel file content (.xlsm).
- **Headers:**
  - `Content-Disposition`: Attachment with filename.
  - `X-Updated-Cells`: Number of cells updated.

### 3. Upload & Update

**POST** `/upload-update`

Uploads a custom Excel file, updates specified cells, and returns the modified file.

- **Use Case:** Updating a user-provided Excel file rather than the server template.
- **Content-Type:** `multipart/form-data`

#### Form Data Parameters

| Parameter | Type | Description |
|:----------|:-----|:------------|
| `excelFile` | Binary | The .xlsm file to update |
| `cellUpdates` | String | JSON string array of cell updates |

#### Example Request

```bash
curl -X POST 'http://localhost:8081/upload-update' \
  --form 'excelFile=@"/path/to/your/file.xlsm"' \
  --form 'cellUpdates="[{\"key\": \"A1\", \"value\": \"Hello\"}, {\"key\": \"B2\", \"value\": 123}]"' \
  -o result.xlsm
```

---

## Error Handling

The API uses standard HTTP status codes.

| Code | Description |
|:-----|:------------|
| 200  | Request was successful |
| 500  | Internal server error (check message and error fields) |

### Error Response Schema

```json
{
  "status": "error",
  "message": "Failed to update Excel file",
  "error": "Detailed error description",
  "requestId": "unique-request-id"
}
