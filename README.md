# Employee Asset Management System

A comprehensive web-based asset management solution for enterprises using SharePoint Framework (SPFx), React, ASP.NET Core, and SQL Server.

## 🎯 Overview

The Employee Asset Management System enables organizations to:

✅ Manage employee asset requests (laptops, peripherals, etc.)
✅ Track and report incidents on assigned assets
✅ Monitor asset assignment lifecycle
✅ View request history and incident tracking
✅ Generate analytics and reports
✅ Maintain audit logs for compliance
✅ Send notifications for request updates

## ✨ Features

### For Employees
- 📊 Dashboard with request and incident statistics
- 📝 Request new assets with priority levels
- 🔧 Report incidents and issues with attachments
- 👁️ Track request approval status
- 🏷️ View assigned assets with conditions
- 📥 Download request and incident history
- 📢 Receive notifications for approvals and updates

### For Managers
- ✅ Approve or reject asset requests
- 📋 Review pending requests
- 🔍 Track team's assets and incidents
- 📊 View department analytics

### For IT Administrators
- 🛠️ Manage asset inventory
- 👤 Manage user roles and permissions
- 📝 Generate compliance reports
- 🔐 View and manage audit logs

## 🛠️ Technology Stack

### Frontend
- **Framework**: React 17
- **UI Library**: Fluent UI (Office UI Fabric)
- **Build Tool**: SharePoint Framework (SPFx) v1.22.2
- **Charting**: Chart.js & react-chartjs-2
- **Language**: TypeScript

### Backend
- **Runtime**: ASP.NET Core 8
- **ORM**: Entity Framework Core
- **API**: RESTful with Swagger/OpenAPI
- **Authentication**: JWT (ready to implement)
- **Language**: C#

### Database
- **Engine**: SQL Server 2019+
- **Schema**: Normalized relational design
- **Stored Procedures**: Performance-optimized queries

## 🚀 Quick Start

### Prerequisites
- Node.js 22.14.0+
- .NET 8 SDK
- SQL Server 2019+
- Visual Studio Code or Visual Studio 2022

### Frontend Setup
```bash
cd INVENTORY
npm install
npm start
```

### Backend Setup
```bash
cd Backend/EmployeeAssetManagement.Api
dotnet restore
dotnet build
dotnet run
```

API will be available at: `https://localhost:5001/swagger`

### Database Setup
```bash
sqlcmd -S YOUR_SERVER -U sa -P YOUR_PASSWORD -i Backend/Database/DatabaseSchema.sql
```

For detailed setup instructions, see [SETUP_GUIDE.md](SETUP_GUIDE.md).

## 📁 Project Structure

```
INVENTORY/
├── src/webparts/employeeManagement/    # React Web Part
├── Backend/
│   ├── EmployeeAssetManagement.Api/    # ASP.NET Core API
│   └── Database/                       # SQL scripts
├── SETUP_GUIDE.md                      # Complete setup guide
├── FOLDER_STRUCTURE.md                 # Folder descriptions
├── API_REFERENCE.md                    # API endpoints
├── ARCHITECTURE.md                     # System architecture
└── README.md                           # This file
```

## 📚 Documentation

- **[SETUP_GUIDE.md](SETUP_GUIDE.md)** - Complete installation and configuration guide
- **[FOLDER_STRUCTURE.md](FOLDER_STRUCTURE.md)** - Project organization and file descriptions
- **[API_REFERENCE.md](API_REFERENCE.md)** - API endpoints and specifications
- **[ARCHITECTURE.md](ARCHITECTURE.md)** - System design and architecture diagrams

## 🔗 API Endpoints

### Dashboard
- `GET /api/dashboard/stats` - Get dashboard statistics

### Asset Requests
- `POST /api/asset-requests` - Create request
- `GET /api/asset-requests/employee` - Get employee requests
- `PUT /api/asset-requests/{id}/cancel` - Cancel request
- `PUT /api/asset-requests/{id}/approve` - Approve request

### Incidents
- `POST /api/incidents` - Report incident
- `GET /api/incidents/employee` - Get employee incidents
- `PUT /api/incidents/{id}/status` - Update incident status

### Assets
- `GET /api/assets` - Get all assets
- `GET /api/assets/available` - Get available assets
- `GET /api/assets/assigned` - Get assigned assets

For complete API documentation, see [API_REFERENCE.md](API_REFERENCE.md).

## 🗄️ Database Schema

The system uses 6 main tables:
- **Employees** - User information and roles
- **Assets** - Inventory of available assets
- **AssetRequests** - Employee asset requests
- **Incidents** - Asset issue tracking
- **AssetAssignments** - Asset-to-employee mappings
- **AuditLogs** - Compliance and audit trail

## 🚢 Deployment

### Development
```bash
# Frontend
npm run build
npm run package-solution

# Backend
dotnet build
dotnet run
```

### Production
```bash
# Frontend
npm run build -- --production
npm run package-solution -- --production

# Backend
dotnet publish -c Release -o ./publish
```

For detailed deployment instructions, see [SETUP_GUIDE.md#deployment](SETUP_GUIDE.md#deployment).

## ⚙️ Configuration

### Frontend Configuration
Edit `src/webparts/employeeManagement/EmployeeManagementWebPart.ts` and set the API base URL.

### Backend Configuration
Edit `appsettings.json` and provide your SQL Server connection string.

## 🔐 Security

- ✅ JWT authentication ready
- ✅ CORS configuration
- ✅ Input validation
- ✅ SQL injection prevention (EF Core)
- ✅ Audit logging
- ✅ Role-based authorization

## 📊 Key Components

### Frontend Components
- Dashboard - Statistics and analytics
- AssetRequestModule - Request form
- IncidentRequestModule - Incident reporting
- MyRequests - Request tracking
- MyAssignedAssets - Asset inventory
- IncidentHistory - Incident tracking

### Backend Controllers
- AssetRequestController - Request management
- IncidentController - Incident management
- DashboardController - Statistics
- AssetsController - Asset inventory

## 🧪 Testing

### Frontend
```bash
npm test
```

### Backend
```bash
dotnet test
```

## 🐛 Troubleshooting

### Frontend Issues
- **Web part not showing?** Clear cache (Ctrl+F5)
- **CORS errors?** Check API URL configuration
- **Build errors?** Run `npm cache clean --force && npm install`

### Backend Issues
- **Database connection error?** Verify connection string
- **Port in use?** Change port in `launchSettings.json`

For more help, see [SETUP_GUIDE.md#troubleshooting](SETUP_GUIDE.md#troubleshooting).

## 🛣️ Roadmap

### Future Features
- [ ] Advanced reporting and analytics
- [ ] Asset lifecycle management
- [ ] Mobile app support
- [ ] Automated approval workflows
- [ ] Integration with external systems
- [ ] Multi-tenancy support

## 📝 Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0.0 | 2024-01-01 | Initial release |

## 📞 Support

For issues and questions, please refer to the documentation or contact the development team.

## 📄 License

Proprietary - All rights reserved

## 🔗 Resources

- [SharePoint Framework Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/)
- [React Documentation](https://reactjs.org)
- [ASP.NET Core Documentation](https://docs.microsoft.com/en-us/dotnet/core/aspnet/)
- [Fluent UI Components](https://developer.microsoft.com/en-us/fluentui)

---

**Last Updated**: 2024  
**Version**: 1.0.0  
**Status**: Production Ready ✅

> Include any additional steps as needed.

Other build commands can be listed using `heft --help`.

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Heft Documentation](https://heft.rushstack.io/)