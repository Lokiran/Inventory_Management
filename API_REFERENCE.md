# Employee Asset Management System - API Reference

## Base URL
```
https://your-api.com/api
```

## Authentication
All API endpoints require Bearer token authentication:

```
Authorization: Bearer {jwt_token}
```

## Response Format

All responses follow this format:

**Success (2xx):**
```json
{
  "data": { ... },
  "message": "Success",
  "statusCode": 200
}
```

**Error (4xx, 5xx):**
```json
{
  "error": "Error message",
  "statusCode": 400
}
```

---

## Endpoints

### Dashboard

#### Get Dashboard Statistics
```
GET /dashboard/stats?email={email}
```

**Query Parameters:**
- `email` (string, required): Employee email address

**Response:**
```json
{
  "totalRequests": 10,
  "pendingRequests": 2,
  "approvedRequests": 5,
  "resolvedIncidents": 8,
  "openIncidents": 1
}
```

**Status Codes:**
- `200`: Success
- `404`: Employee not found
- `500`: Server error

---

### Asset Requests

#### Create Asset Request
```
POST /asset-requests
Content-Type: application/json
```

**Request Body:**
```json
{
  "employeeEmail": "user@example.com",
  "assetType": "Laptop",
  "assetName": "Dell XPS 13",
  "quantity": 1,
  "priority": "High",
  "reasonDescription": "Replacement for damaged device",
  "requiredDate": "2024-01-15T00:00:00Z"
}
```

**Response:**
```json
{
  "requestId": 1,
  "employeeId": 1,
  "assetType": "Laptop",
  "assetName": "Dell XPS 13",
  "status": "Pending",
  "requestDate": "2024-01-01T10:00:00Z"
}
```

**Status Codes:**
- `201`: Created
- `400`: Bad request
- `404`: Employee not found
- `500`: Server error

---

#### Get Employee Asset Requests
```
GET /asset-requests/employee?email={email}
```

**Query Parameters:**
- `email` (string, required): Employee email address

**Response:**
```json
[
  {
    "id": "1",
    "assetType": "Laptop",
    "assetName": "Dell XPS 13",
    "quantity": 1,
    "priority": "High",
    "status": "Pending",
    "requestDate": "2024-01-01T10:00:00Z",
    "requiredDate": "2024-01-15T00:00:00Z",
    "description": "Replacement for damaged device"
  }
]
```

**Status Codes:**
- `200`: Success
- `404`: Employee not found
- `500`: Server error

---

#### Get Single Asset Request
```
GET /asset-requests/{id}
```

**Path Parameters:**
- `id` (integer, required): Request ID

**Response:**
```json
{
  "requestId": 1,
  "employeeId": 1,
  "assetType": "Laptop",
  "quantity": 1,
  "status": "Pending",
  "requestDate": "2024-01-01T10:00:00Z"
}
```

**Status Codes:**
- `200`: Success
- `404`: Request not found
- `500`: Server error

---

#### Cancel Asset Request
```
PUT /asset-requests/{id}/cancel
```

**Path Parameters:**
- `id` (integer, required): Request ID

**Response:**
```json
{
  "message": "Request cancelled successfully"
}
```

**Status Codes:**
- `200`: Success
- `400`: Bad request (not pending)
- `404`: Request not found
- `500`: Server error

---

#### Approve Asset Request (Admin/Manager)
```
PUT /asset-requests/{id}/approve
Content-Type: application/json
```

**Path Parameters:**
- `id` (integer, required): Request ID

**Request Body:**
```json
{
  "approvedBy": 3,
  "comment": "Approved for replacement"
}
```

**Response:**
```json
{
  "message": "Request approved successfully"
}
```

**Status Codes:**
- `200`: Success
- `404`: Request not found
- `500`: Server error

---

### Incidents

#### Create Incident
```
POST /incidents
Content-Type: multipart/form-data
```

**Form Data:**
- `employeeEmail` (string, required): Employee email
- `assetId` (string, required): Asset serial number
- `issueType` (string, required): Type of issue
- `issueDescription` (string, required): Detailed description
- `priority` (string, required): Priority level
- `reportedDate` (datetime, required): Report date
- `attachment` (file, optional): Supporting document

**Response:**
```json
{
  "incidentId": 1,
  "incidentNumber": "INC-20240101100000",
  "status": "Open",
  "reportedDate": "2024-01-01T10:00:00Z"
}
```

**Status Codes:**
- `201`: Created
- `400`: Bad request
- `404`: Employee or asset not found
- `500`: Server error

---

#### Get Employee Incidents
```
GET /incidents/employee?email={email}
```

**Query Parameters:**
- `email` (string, required): Employee email address

**Response:**
```json
[
  {
    "id": "1",
    "incidentId": "INC-20240101100000",
    "assetId": "DELL-001",
    "assetName": "Dell XPS 13",
    "issueType": "Hardware Damage",
    "issueDescription": "Screen is cracked",
    "priority": "High",
    "status": "Open",
    "reportedDate": "2024-01-01T10:00:00Z",
    "assignedTo": "John Support"
  }
]
```

**Status Codes:**
- `200`: Success
- `404`: Employee not found
- `500`: Server error

---

#### Get Single Incident
```
GET /incidents/{id}
```

**Path Parameters:**
- `id` (integer, required): Incident ID

**Response:**
```json
{
  "incidentId": 1,
  "incidentNumber": "INC-20240101100000",
  "assetId": 1,
  "issueDescription": "Screen is cracked",
  "status": "Open"
}
```

**Status Codes:**
- `200`: Success
- `404`: Incident not found
- `500`: Server error

---

#### Update Incident Status
```
PUT /incidents/{id}/status
Content-Type: application/json
```

**Path Parameters:**
- `id` (integer, required): Incident ID

**Request Body:**
```json
{
  "status": "Resolved",
  "resolutionNotes": "Screen replaced with new one"
}
```

**Response:**
```json
{
  "message": "Incident updated successfully"
}
```

**Status Codes:**
- `200`: Success
- `404`: Incident not found
- `500`: Server error

---

### Assets

#### Get All Assets
```
GET /assets
```

**Response:**
```json
[
  {
    "id": 1,
    "assetType": "Laptop",
    "assetName": "Dell XPS 13",
    "serialNumber": "DELL-001",
    "status": "Available",
    "location": "IT Store"
  }
]
```

**Status Codes:**
- `200`: Success
- `500`: Server error

---

#### Get Available Assets
```
GET /assets/available
```

**Response:**
```json
[
  {
    "id": 1,
    "assetType": "Laptop",
    "assetName": "Dell XPS 13",
    "serialNumber": "DELL-001",
    "status": "Available",
    "location": "IT Store"
  }
]
```

**Status Codes:**
- `200`: Success
- `500`: Server error

---

#### Get Assigned Assets
```
GET /assets/assigned?email={email}
```

**Query Parameters:**
- `email` (string, required): Employee email address

**Response:**
```json
[
  {
    "id": "1",
    "assetType": "Laptop",
    "assetName": "Dell XPS 13",
    "serialNumber": "DELL-001",
    "assignmentDate": "2023-12-01T00:00:00Z",
    "status": "Assigned",
    "condition": "Good",
    "location": "Desk 101"
  }
]
```

**Status Codes:**
- `200`: Success
- `404`: Employee not found
- `500`: Server error

---

#### Get Single Asset
```
GET /assets/{id}
```

**Path Parameters:**
- `id` (integer, required): Asset ID

**Response:**
```json
{
  "id": 1,
  "assetType": "Laptop",
  "assetName": "Dell XPS 13",
  "serialNumber": "DELL-001",
  "status": "Available",
  "location": "IT Store"
}
```

**Status Codes:**
- `200`: Success
- `404`: Asset not found
- `500`: Server error

---

## Enumerations

### Asset Types
- `Laptop`
- `Mouse`
- `Keyboard`
- `Monitor`
- `Headset`
- `Mobile`

### Priority Levels
- `Low`
- `Medium`
- `High`
- `Urgent` (for requests)
- `Critical` (for incidents)

### Request Status
- `Pending`
- `Approved`
- `Rejected`
- `Issued`
- `Cancelled`

### Incident Status
- `Open`
- `In Progress`
- `Resolved`
- `Closed`

### Asset Status
- `Available`
- `Assigned`
- `Damaged`
- `Retired`

### Condition Levels
- `Good`
- `Fair`
- `Poor`

### Issue Types
- `Hardware Damage`
- `Software Issue`
- `Not Working`
- `Missing Parts`
- `Performance Issue`
- `Other`

### User Roles
- `Employee`
- `Manager`
- `Admin`

---

## Error Codes

| Code | Message | Description |
|------|---------|-------------|
| 400 | Bad Request | Invalid request parameters |
| 401 | Unauthorized | Missing or invalid authentication |
| 403 | Forbidden | Insufficient permissions |
| 404 | Not Found | Resource not found |
| 409 | Conflict | Resource already exists |
| 500 | Internal Server Error | Server error |
| 503 | Service Unavailable | Database connection error |

---

## Pagination (Future Implementation)

```
GET /asset-requests/employee?email={email}&page=1&pageSize=10&sortBy=requestDate&sortOrder=desc
```

**Query Parameters:**
- `page` (integer): Page number (default: 1)
- `pageSize` (integer): Items per page (default: 10)
- `sortBy` (string): Sort column
- `sortOrder` (string): asc or desc

---

## Rate Limiting (Future Implementation)

- API Rate Limit: 100 requests per minute per user
- Headers returned:
  - `X-RateLimit-Limit`
  - `X-RateLimit-Remaining`
  - `X-RateLimit-Reset`

---

## Testing with Postman

### Import Collection

1. Create new Postman collection
2. Add requests for each endpoint
3. Configure authentication (Bearer token)
4. Set up environment variables

**Environment Variables:**
```json
{
  "base_url": "https://your-api.com/api",
  "email": "user@example.com",
  "token": "your-jwt-token"
}
```

---

## Webhooks (Future Implementation)

Subscribe to events:
- `asset_request.created`
- `asset_request.approved`
- `incident.created`
- `incident.resolved`

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0.0 | 2024-01-01 | Initial API release |

---

For more information, visit the Swagger documentation at:
```
https://your-api.com/swagger
```

Last Updated: 2024
