# Excel Update Integration

A MuleSoft integration flow that listens to HTTP POST requests on the `/update/` endpoint and updates Excel files with key-value pairs from the request payload.

## Features

- **HTTP Listener**: Accepts POST requests on `/update/` endpoint
- **Excel File Processing**: Creates unique copies of template Excel files
- **Dynamic Cell Updates**: Updates Excel cells based on key-value pairs in the payload
- **Error Handling**: Comprehensive error handling with detailed logging
- **Unique File Generation**: Creates timestamped files with request IDs for traceability

## Project Structure

```
excel-update-integration/
├── pom.xml                                    # Maven configuration
├── mule-artifact.json                         # Mule artifact configuration
├── src/main/
│   ├── mule/
│   │   └── excel-update-flow.xml             # Main flow implementation
│   └── resources/
│       ├── application.properties            # Configuration properties
│       ├── template.xlsm                     # Template Excel file
│       └── output/                           # Directory for updated files
└── README.md                                 # This file
```

## API Usage

### Endpoint
```
POST http://localhost:8081/update
```

### Request Payload
Send an array of key-value objects where:
- `key`: Excel cell address (e.g., "A1", "B2", "C3")
- `value`: Value to set in the cell

```json
[
  {
    "key": "A1",
    "value": "Customer Name"
  },
  {
    "key": "B1", 
    "value": "John Doe"
  },
  {
    "key": "A2",
    "value": "Amount"
  },
  {
    "key": "B2",
    "value": 1500.50
  }
]
```

### Response
```json
{
  "status": "success",
  "message": "Excel file updated successfully",
  "requestId": "550e8400-e29b-41d4-a716-446655440000",
  "fileName": "updated-excel-20241130-213000-550e8400-e29b-41d4-a716-446655440000.xlsm",
  "updatedCells": 4,
  "timestamp": "20241130-213000"
}
```

## Configuration

### Application Properties
- `http.host`: HTTP listener host (default: 0.0.0.0)
- `http.port`: HTTP listener port (default: 8081)
- `excel.template.path`: Path to template Excel file
- `excel.output.directory`: Directory for updated Excel files

### Dependencies
- **HTTP Connector**: For REST API endpoints
- **File Connector**: For file operations
- **Microsoft Excel Connector**: For Excel file manipulation
- **APIKit Module**: For API management

## Running the Application

1. **Prerequisites**:
   - Java 17+
   - Maven 3.6+
   - MuleSoft Runtime 4.10.0+

2. **Build the project**:
   ```bash
   mvn clean compile
   ```

3. **Run locally**:
   ```bash
   mvn mule:run
   ```

4. **Test the endpoint**:
   ```bash
   curl -X POST http://localhost:8081/update \
     -H "Content-Type: application/json" \
     -d '[{"key":"A1","value":"Test Value"}]'
   ```

## How It Works

1. **Request Reception**: HTTP listener receives POST request on `/update/` endpoint
2. **Payload Validation**: Validates that payload is an array of key-value objects
3. **File Operations**: 
   - Creates unique filename with timestamp and request ID
   - Copies template Excel file to output directory with unique name
4. **Excel Updates**: Iterates through payload and updates each specified cell
5. **Response**: Returns success response with file details and update count

## Error Handling

The flow includes comprehensive error handling:
- Invalid payload structure validation
- File operation error handling
- Excel update error handling
- Detailed error logging with correlation IDs

## Repository

GitHub: https://github.com/savy-mulesoft/excel-update-integration

## License

This project is part of MuleSoft integration examples.
