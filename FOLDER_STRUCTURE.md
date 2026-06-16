# Employee Asset Management System - Folder Structure

## Directory Layout

```
INVENTORY/
├── src/
│   └── webparts/
│       └── employeeManagement/           # React Web Part
│           ├── EmployeeManagementWebPart.ts
│           ├── EmployeeManagementWebPart.manifest.json
│           ├── components/
│           │   ├── EmployeeManagementPanel.tsx          # Main container
│           │   ├── EmployeeManagementPanel.module.scss
│           │   ├── IEmployeeManagementProps.ts          # Interfaces
│           │   ├── Dashboard/
│           │   │   ├── Dashboard.tsx                    # Dashboard component
│           │   │   └── Dashboard.module.scss
│           │   ├── AssetRequest/
│           │   │   ├── AssetRequestModule.tsx           # Asset request form
│           │   │   └── AssetRequestModule.module.scss
│           │   ├── IncidentRequest/
│           │   │   ├── IncidentRequestModule.tsx        # Incident form
│           │   │   └── IncidentRequestModule.module.scss
│           │   ├── MyRequests/
│           │   │   ├── MyRequests.tsx                   # Request list
│           │   │   └── MyRequests.module.scss
│           │   ├── MyAssignedAssets/
│           │   │   ├── MyAssignedAssets.tsx            # Asset list
│           │   │   └── MyAssignedAssets.module.scss
│           │   └── IncidentHistory/
│           │       ├── IncidentHistory.tsx              # Incident history
│           │       └── IncidentHistory.module.scss
│           ├── services/
│           │   └── InventoryService.ts                  # API service
│           └── loc/
│               ├── mystrings.d.ts
│               └── en-us.json
├── lib/                                  # Compiled output
├── lib-commonjs/                         # CommonJS output
├── sharepoint/
│   └── solution/                         # SPFx package
│       └── spfx-project.sppkg
├── config/                               # SPFx config files
├── Backend/
│   └── EmployeeAssetManagement.Api/      # .NET Core API
│       ├── Program.cs                    # Startup configuration
│       ├── appsettings.json
│       ├── appsettings.Production.json
│       ├── EmployeeAssetManagement.Api.csproj
│       ├── Models/
│       │   └── DataModels.cs            # Entity models
│       ├── Data/
│       │   └── InventoryContext.cs      # Entity Framework DbContext
│       ├── Controllers/
│       │   ├── AssetRequestController.cs
│       │   ├── IncidentController.cs
│       │   ├── DashboardController.cs
│       │   └── AssetsController.cs
│       ├── Services/                     # Business logic (optional)
│       │   ├── IAssetService.cs
│       │   └── AssetService.cs
│       └── Middleware/                   # Custom middleware
│           ├── AuthenticationMiddleware.cs
│           └── ErrorHandlingMiddleware.cs
├── Backend/Database/
│   ├── DatabaseSchema.sql                # SQL Server schema
│   ├── StoredProcedures.sql             # Stored procedures
│   └── SampleData.sql                    # Sample data
├── Documentation/
│   ├── SETUP_GUIDE.md                   # This file (setup guide)
│   ├── ARCHITECTURE.md                  # Architecture documentation
│   ├── API_REFERENCE.md                 # API endpoints reference
│   └── DEPLOYMENT.md                    # Deployment guide
├── package.json                          # NPM dependencies
├── tsconfig.json                         # TypeScript config
└── README.md                             # Project README

```

## Folder Descriptions

### Frontend (src/webparts/employeeManagement/)

#### Components
- **EmployeeManagementPanel.tsx**: Main navigation container with sidebar
- **Dashboard.tsx**: Dashboard with statistics and charts
- **AssetRequestModule.tsx**: Form for requesting assets
- **IncidentRequestModule.tsx**: Form for reporting incidents
- **MyRequests.tsx**: Table showing employee's requests
- **MyAssignedAssets.tsx**: Table showing assigned assets
- **IncidentHistory.tsx**: Table showing incident history

#### Services
- **InventoryService.ts**: API client for all HTTP calls
  - Methods for CRUD operations
  - Error handling
  - Token management

#### Localization
- **en-us.json**: English localization strings
- **mystrings.d.ts**: TypeScript interface for strings

### Backend (Backend/EmployeeAssetManagement.Api/)

#### Models
- **DataModels.cs**: Entity Framework models
  - Employee
  - Asset
  - AssetRequest
  - Incident
  - AssetAssignment
  - AuditLog

#### Data
- **InventoryContext.cs**: DbContext configuration
  - Entity mappings
  - Relationships
  - Seed data (optional)

#### Controllers
- **AssetRequestController.cs**: Asset request endpoints
- **IncidentController.cs**: Incident endpoints
- **DashboardController.cs**: Dashboard statistics
- **AssetsController.cs**: Asset management endpoints

#### Services (Optional)
- Business logic layer
- Separate from controllers
- Reusable across controllers

### Database (Backend/Database/)

- **DatabaseSchema.sql**: Complete database schema
  - Tables
  - Indexes
  - Relationships
  - Sample data
  - Stored procedures

## Key Technologies

### Frontend
- React 17
- TypeScript
- Fluent UI (Office UI Fabric)
- Chart.js for data visualization
- SCSS for styling

### Backend
- ASP.NET Core 8
- Entity Framework Core
- Swagger/OpenAPI documentation
- JWT authentication (ready to implement)

### Database
- SQL Server 2019+
- Stored procedures for complex queries
- Audit logging

## File Naming Conventions

### React Components
- PascalCase for components: `Dashboard.tsx`
- Module styles: `Dashboard.module.scss`

### Controllers
- PascalCase: `AssetRequestController.cs`
- Suffix with "Controller"

### Models
- PascalCase: `Employee.cs`

### Services
- Interface: `IAssetService.cs`
- Implementation: `AssetService.cs`

### Database
- Tables: PascalCase
- Stored Procedures: `sp_GetAssets`
- Indexes: `idx_TableName_Column`

## Build Output

### Frontend Build
- **lib/**: Contains compiled JavaScript
- **lib-commonjs/**: CommonJS format
- **sharepoint/solution/**: SPFx package (.sppkg)

### Backend Build
- **bin/**: Compiled binaries
- **obj/**: Temporary build files
- **publish/**: Production-ready output

## Dependencies Management

### Frontend (package.json)
```json
{
  "@fluentui/react": "^8.106.4",
  "react": "17.0.1",
  "react-chartjs-2": "^5.3.1",
  "chart.js": "^4.5.1"
}
```

### Backend (.csproj)
- Microsoft.EntityFrameworkCore
- Microsoft.EntityFrameworkCore.SqlServer
- Swashbuckle.AspNetCore (Swagger)
- System.IdentityModel.Tokens.Jwt

## Configuration Files

### SPFx Configuration
- **tsconfig.json**: TypeScript compiler options
- **config/package-solution.json**: SPFx package settings
- **config/serve.json**: Development server settings

### API Configuration
- **appsettings.json**: Development settings
- **appsettings.Production.json**: Production settings

## Development Workflow

1. **Local Development**
   - Run SPFx dev server: `npm start`
   - Run API: `dotnet run`
   - Database: Local SQL Server

2. **Build & Test**
   - Build frontend: `npm run build`
   - Build backend: `dotnet build`
   - Run tests: `npm test`, `dotnet test`

3. **Packaging**
   - Package SPFx: `npm run package-solution`
   - Publish backend: `dotnet publish`

4. **Deployment**
   - Upload .sppkg to App Catalog
   - Deploy API to App Service/IIS
   - Update database schema

## Performance Considerations

- Frontend: Code splitting, lazy loading
- Backend: Async operations, caching
- Database: Indexes on frequently queried columns
- API: GZIP compression, response caching

## Security Considerations

- JWT authentication for API
- Role-based authorization
- CORS configuration
- Input validation
- SQL injection prevention (EF Core parameterized queries)
- Audit logging for compliance

---

For more details, see:
- [SETUP_GUIDE.md](SETUP_GUIDE.md)
- [API_REFERENCE.md](API_REFERENCE.md)
- [DEPLOYMENT.md](DEPLOYMENT.md)
