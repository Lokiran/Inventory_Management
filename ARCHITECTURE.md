# Architecture Documentation

## System Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                  SharePoint Online                          │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌──────────────────────────────────────────────────────┐  │
│  │         Employee Management Web Part (SPFx)         │  │
│  │                     (React 17)                       │  │
│  ├──────────────────────────────────────────────────────┤  │
│  │ ┌────────────┐ ┌────────────┐ ┌────────────┐        │  │
│  │ │ Dashboard  │ │  Requests  │ │ Incidents  │ ...    │  │
│  │ └────────────┘ └────────────┘ └────────────┘        │  │
│  │                                                      │  │
│  │ ┌──────────────────────────────────────────────┐   │  │
│  │ │      InventoryService (API Client)           │   │  │
│  │ │  - REST API calls                            │   │  │
│  │ │  - Error handling                            │   │  │
│  │ │  - Token management                          │   │  │
│  │ └──────────────────────────────────────────────┘   │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
└─────────────────────────────────────────────────────────────┘
                           │
                  (HTTPS / REST API)
                           │
        ┌──────────────────┴──────────────────┐
        │                                     │
┌───────▼──────────────────────┐   ┌────────▼────────────────┐
│     API Gateway / Load       │   │   Azure Key Vault       │
│     Balancer (Optional)      │   │  (Secrets Management)   │
└───────┬──────────────────────┘   └────────────────────────┘
        │
┌───────▼──────────────────────────────────────────────────┐
│          ASP.NET Core 8 API Server                       │
├────────────────────────────────────────────────────────┤
│                                                          │
│  ┌────────────────────────────────────────────────┐   │
│  │           Swagger / OpenAPI UI                │   │
│  └────────────────────────────────────────────────┘   │
│                                                          │
│  ┌────────────────────────────────────────────────┐   │
│  │           Controllers                          │   │
│  │  ┌─────────────┬─────────────┬──────────────┐ │   │
│  │  │  Asset Req  │  Incidents  │  Dashboard   │ │   │
│  │  │  Controller │ Controller  │ Controller   │ │   │
│  │  └─────────────┴─────────────┴──────────────┘ │   │
│  └────────────────────────────────────────────────┘   │
│                                                          │
│  ┌────────────────────────────────────────────────┐   │
│  │        Business Logic / Services               │   │
│  │  ┌──────────────┐                             │   │
│  │  │ Asset Svc    │                             │   │
│  │  │ Request Svc  │                             │   │
│  │  │ Incident Svc │                             │   │
│  │  └──────────────┘                             │   │
│  └────────────────────────────────────────────────┘   │
│                                                          │
│  ┌────────────────────────────────────────────────┐   │
│  │      Data Access Layer (EF Core)               │   │
│  │  ┌──────────────────────────────────────────┐ │   │
│  │  │     InventoryContext (DbContext)         │ │   │
│  │  │  - Entity mappings                       │ │   │
│  │  │  - Relationships                         │ │   │
│  │  │  - Migrations                            │ │   │
│  │  └──────────────────────────────────────────┘ │   │
│  └────────────────────────────────────────────────┘   │
│                                                          │
└───────┬──────────────────────────────────────────────────┘
        │
        │ (SQL Connection)
        │
┌───────▼──────────────────────────────────────────────────┐
│           SQL Server 2019 / Azure SQL                   │
├────────────────────────────────────────────────────────┤
│                                                          │
│  ┌──────────────────────────────────────────────┐     │
│  │        Database: EmployeeAssetManagement    │     │
│  │                                              │     │
│  │  ┌───────────┐  ┌───────────┐  ┌──────────┐ │     │
│  │  │ Employees │  │  Assets   │  │ Requests │ │     │
│  │  └───────────┘  └───────────┘  └──────────┘ │     │
│  │                                              │     │
│  │  ┌───────────┐  ┌──────────────┐  ┌───────┐ │     │
│  │  │ Incidents │  │ Assignments  │  │ Logs  │ │     │
│  │  └───────────┘  └──────────────┘  └───────┘ │     │
│  │                                              │     │
│  │  ┌──────────────────────────────────────┐  │     │
│  │  │  Indexes & Stored Procedures         │  │     │
│  │  │  - Performance optimizations         │  │     │
│  │  │  - Complex queries                   │  │     │
│  │  └──────────────────────────────────────┘  │     │
│  └──────────────────────────────────────────────┘     │
│                                                          │
└──────────────────────────────────────────────────────────┘
```

## Component Architecture

### Frontend Components

```
EmployeeManagementWebPart
│
├── EmployeeManagementPanel (Main Container)
│   ├── Navigation Sidebar
│   │   └── Nav Links
│   │
│   └── Content Area
│       ├── Dashboard
│       │   ├── Welcome Section
│       │   ├── Quick Action Cards
│       │   └── Charts (Pie, Line)
│       │
│       ├── AssetRequestModule
│       │   ├── Employee Info Section
│       │   ├── Asset Info Section
│       │   └── Request Details Section
│       │
│       ├── IncidentRequestModule
│       │   ├── Asset Info Section
│       │   ├── Issue Info Section
│       │   └── Attachment Upload
│       │
│       ├── MyRequests
│       │   ├── Search & Filter
│       │   └── DetailsList
│       │
│       ├── MyAssignedAssets
│       │   ├── Search
│       │   └── DetailsList
│       │
│       └── IncidentHistory
│           ├── Search & Filter
│           └── DetailsList

InventoryService (Shared)
├── API Methods
├── Error Handling
└── Token Management
```

### Backend Architecture

```
ASP.NET Core Application
│
├── Controllers
│   ├── AssetRequestController
│   ├── IncidentController
│   ├── DashboardController
│   └── AssetsController
│
├── Models
│   ├── Employee
│   ├── Asset
│   ├── AssetRequest
│   ├── Incident
│   ├── AssetAssignment
│   └── AuditLog
│
├── Services (Optional Layer)
│   ├── IAssetService
│   ├── AssetService
│   ├── IRequestService
│   └── RequestService
│
├── Data Access
│   ├── InventoryContext
│   └── Migrations
│
├── Middleware
│   ├── Authentication
│   ├── Error Handling
│   └── Logging
│
└── Configuration
    ├── Startup
    ├── Dependency Injection
    └── CORS Policy
```

## Data Flow

### Create Asset Request Flow

```
┌─────────────────────────────────────────────────┐
│ Employee fills form in AssetRequestModule       │
└────────────┬────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ Submit form data                                │
│ {employeeId, assetType, quantity, ...}         │
└────────────┬────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ InventoryService.createAssetRequest()           │
│ POST /api/asset-requests                        │
└────────────┬────────────────────────────────────┘
             │
             ▼ (HTTPS)
┌─────────────────────────────────────────────────┐
│ AssetRequestController.CreateAssetRequest()    │
└────────────┬────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ Validate input                                  │
│ Check employee exists                           │
└────────────┬────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ Create AssetRequest entity                      │
│ Set Status = "Pending"                          │
└────────────┬────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ InventoryContext.SaveChangesAsync()             │
│ (Entity Framework creates database record)      │
└────────────┬────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ Create AuditLog entry                           │
└────────────┬────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ Return 201 Created response                     │
│ With request details                            │
└────────────┬────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────────┐
│ Show success message to user                    │
│ Clear form                                      │
└─────────────────────────────────────────────────┘
```

## Database Schema Relationships

```
Employees (1) ─── (Many) AssetRequests
    │
    ├─── (Many) Incidents
    │
    ├─── (Many) AssetAssignments
    │
    └─── (Many) AuditLogs

Assets (1) ─── (Many) AssetRequests
   │
   ├─── (Many) Incidents
   │
   └─── (Many) AssetAssignments
```

## Authentication & Authorization Flow

```
┌──────────────────────────────┐
│  User Login (Azure AD)       │
└──────┬───────────────────────┘
       │
       ▼
┌──────────────────────────────┐
│  JWT Token Generated         │
│  - Claims (name, email, role)│
│  - Expiration time           │
└──────┬───────────────────────┘
       │
       ▼
┌──────────────────────────────┐
│  Token stored in client      │
└──────┬───────────────────────┘
       │
       ▼
┌──────────────────────────────┐
│  API request with token      │
│  Authorization: Bearer {token}
└──────┬───────────────────────┘
       │
       ▼
┌──────────────────────────────┐
│  API validates token         │
│  - Signature verification    │
│  - Expiration check          │
│  - Role verification         │
└──────┬───────────────────────┘
       │
       ├─── Valid ───┐
       │              │
       │              ▼
       │         ┌─────────────┐
       │         │ Process req │
       │         │ Return data │
       │         └─────────────┘
       │
       └─── Invalid ─┐
                      │
                      ▼
                  ┌────────────┐
                  │ 401 / 403  │
                  │ Error      │
                  └────────────┘
```

## Deployment Architecture

```
┌────────────────────────────────────┐
│  CI/CD Pipeline (Azure DevOps)    │
│  - Build                           │
│  - Test                            │
│  - Package                         │
└────────────┬───────────────────────┘
             │
    ┌────────┴────────┐
    │                 │
    ▼                 ▼
┌─────────────────┐ ┌──────────────┐
│  SPFx Package   │ │ API Package  │
│  (.sppkg)       │ │ (.zip)       │
└────────┬────────┘ └──────┬───────┘
         │                 │
         ▼                 ▼
┌──────────────────────────────────┐
│ Release Management               │
│ - Staging environment            │
│ - Production environment         │
└────────┬─────────────────────────┘
         │
    ┌────┴────┐
    │          │
    ▼          ▼
┌──────────┐ ┌──────────────────┐
│SharePoint│ │ App Service / IIS│
│App Cat   │ │ + Database       │
└──────────┘ └──────────────────┘
```

## Security Architecture

```
┌─────────────────────────────────────────┐
│       Azure AD / OAuth 2.0              │
│  (Authentication & Authorization)       │
└──────────────┬──────────────────────────┘
               │
┌──────────────▼──────────────────────────┐
│       JWT Token                         │
│  (Claims: user, roles, permissions)    │
└──────────────┬──────────────────────────┘
               │
┌──────────────▼──────────────────────────┐
│    API Request with Bearer Token        │
│    + HTTPS Encryption                   │
│    + CORS Validation                    │
└──────────────┬──────────────────────────┘
               │
┌──────────────▼──────────────────────────┐
│    Token Validation                     │
│    + Signature check                    │
│    + Expiration check                   │
│    + Role-based authorization           │
└──────────────┬──────────────────────────┘
               │
┌──────────────▼──────────────────────────┐
│    Parameterized Queries                │
│    (SQL Injection Prevention)            │
└──────────────┬──────────────────────────┘
               │
┌──────────────▼──────────────────────────┐
│    Audit Logging                        │
│    (Compliance & Tracking)              │
└─────────────────────────────────────────┘
```

## Performance Optimization Strategies

1. **Frontend**
   - Code splitting
   - Lazy loading components
   - Minification
   - GZIP compression

2. **Backend**
   - Async/await for I/O
   - Connection pooling
   - Query optimization
   - Caching strategies

3. **Database**
   - Indexes on primary keys and foreign keys
   - Partitioning for large tables
   - Stored procedures for complex queries
   - Query execution plans

4. **API**
   - Pagination for large datasets
   - Response caching headers
   - GZIP compression
   - CDN for static content

## Scalability Considerations

- **Horizontal scaling**: Multiple API instances behind load balancer
- **Database scaling**: Read replicas for reporting
- **Caching layer**: Redis for session/token caching
- **CDN**: For static web part assets
- **Message queue**: For async operations (future)

---

Last Updated: 2024
Version: 1.0.0
