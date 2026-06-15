# Quick Reference Guide

## 🚀 Quick Start Commands

### Frontend
```bash
# Install dependencies
npm install

# Start development server
npm start

# Build for production
npm run build -- --production

# Package for SharePoint
npm run package-solution -- --production

# Run tests
npm test

# Clean build
npm run clean
```

### Backend
```bash
# Restore dependencies
dotnet restore

# Build project
dotnet build

# Run application
dotnet run

# Create migration
dotnet ef migrations add MigrationName

# Apply migrations
dotnet ef database update

# Create database from scratch
dotnet ef database drop
dotnet ef database update

# Publish for production
dotnet publish -c Release -o ./publish

# Run tests
dotnet test
```

### Database
```bash
# Create database from script
sqlcmd -S YOUR_SERVER -U sa -P YOUR_PASSWORD -i Backend/Database/DatabaseSchema.sql

# Backup database
BACKUP DATABASE EmployeeAssetManagement TO DISK = 'path\backup.bak';

# Restore database
RESTORE DATABASE EmployeeAssetManagement FROM DISK = 'path\backup.bak';
```

---

## 📁 Key Files & Locations

### Frontend Components
| File | Purpose |
|------|---------|
| `EmployeeManagementWebPart.ts` | Main web part |
| `Dashboard.tsx` | Dashboard component |
| `AssetRequestModule.tsx` | Asset request form |
| `IncidentRequestModule.tsx` | Incident form |
| `MyRequests.tsx` | Request list |
| `MyAssignedAssets.tsx` | Asset list |
| `IncidentHistory.tsx` | Incident list |
| `InventoryService.ts` | API client |

### Backend Controllers
| File | Handles |
|------|---------|
| `AssetRequestController.cs` | Asset requests |
| `IncidentController.cs` | Incidents |
| `DashboardController.cs` | Dashboard stats |
| `AssetsController.cs` | Assets |

### Configuration
| File | Purpose |
|------|---------|
| `appsettings.json` | Dev configuration |
| `appsettings.Production.json` | Prod config |
| `tsconfig.json` | TypeScript config |
| `package.json` | NPM dependencies |

### Documentation
| File | Content |
|------|---------|
| `README.md` | Overview |
| `SETUP_GUIDE.md` | Installation |
| `API_REFERENCE.md` | API docs |
| `DEPLOYMENT.md` | Deployment |
| `ARCHITECTURE.md` | Design |
| `FOLDER_STRUCTURE.md` | Organization |

---

## 🔌 API Quick Reference

### Base URL
```
https://your-api.com/api
```

### Dashboard
```
GET /dashboard/stats?email={email}
```

### Asset Requests
```
POST /asset-requests
GET /asset-requests/employee?email={email}
PUT /asset-requests/{id}/cancel
PUT /asset-requests/{id}/approve
```

### Incidents
```
POST /incidents (multipart/form-data)
GET /incidents/employee?email={email}
PUT /incidents/{id}/status
```

### Assets
```
GET /assets
GET /assets/available
GET /assets/assigned?email={email}
```

---

## 🗄️ Database Schema Quick View

### Employees
```sql
SELECT * FROM dbo.Employees;
-- EmployeeId, Name, Email, Department, Role, Phone, CreatedDate
```

### Assets
```sql
SELECT * FROM dbo.Assets;
-- AssetId, AssetType, AssetName, SerialNumber, Status, Location
```

### AssetRequests
```sql
SELECT * FROM dbo.AssetRequests;
-- RequestId, EmployeeId, AssetType, Quantity, Priority, Status
```

### Incidents
```sql
SELECT * FROM dbo.Incidents;
-- IncidentId, EmployeeId, AssetId, IssueType, Priority, Status
```

### AssetAssignments
```sql
SELECT * FROM dbo.AssetAssignments;
-- AssignmentId, EmployeeId, AssetId, Status, Condition
```

### AuditLogs
```sql
SELECT * FROM dbo.AuditLogs;
-- LogId, EmployeeId, Action, TableName, CreatedDate
```

---

## 🔑 Common Values

### Priority Levels
- `Low`
- `Medium`
- `High`
- `Urgent` (Requests)
- `Critical` (Incidents)

### Asset Types
- `Laptop`
- `Mouse`
- `Keyboard`
- `Monitor`
- `Headset`
- `Mobile`

### Status Values
| Status | Usage |
|--------|-------|
| `Pending` | Request waiting approval |
| `Approved` | Request approved |
| `Rejected` | Request rejected |
| `Issued` | Asset issued to employee |
| `Open` | Incident just reported |
| `In Progress` | Incident being worked on |
| `Resolved` | Incident resolved |
| `Closed` | Incident closed |

### Roles
- `Employee` - Regular employee
- `Manager` - Can approve requests
- `Admin` - Full system access

---

## 🛠️ Troubleshooting Quick Tips

### Frontend Issues
| Issue | Solution |
|-------|----------|
| Web part not showing | Clear cache (Ctrl+F5) |
| CORS error | Check API URL in config |
| Build errors | `npm cache clean --force` |
| Port 3000 in use | Change port in serve.json |

### Backend Issues
| Issue | Solution |
|-------|----------|
| DB connection error | Verify connection string |
| Port 5001 in use | Use `--urls` parameter |
| Migration failure | `dotnet ef database drop` |
| Cannot find package | `dotnet restore` |

### Database Issues
| Issue | Solution |
|-------|----------|
| Login failed | Verify credentials |
| Table exists | `DROP DATABASE` and recreate |
| Connection timeout | Check firewall/network |

---

## 📊 Important URLs

### Development
```
Frontend: http://localhost:3000
Backend: https://localhost:5001
Swagger: https://localhost:5001/swagger
```

### SharePoint
```
App Catalog: https://[tenant]-admin.sharepoint.com/sites/appcatalog
```

### Azure
```
Portal: https://portal.azure.com
```

---

## 🔐 Authentication Setup

### Local Testing
Use Bearer token header:
```
Authorization: Bearer your-test-token
```

### Production Setup
1. Configure Azure AD
2. Update Program.cs with Azure AD config
3. Generate JWT tokens
4. Store secrets in Key Vault

---

## 📦 Deployment Checklist

### Before Deployment
- [ ] Code reviewed
- [ ] Tests passing
- [ ] No console errors
- [ ] Backup created
- [ ] Configuration updated

### Deployment Steps
- [ ] Build frontend
- [ ] Package solution
- [ ] Upload to App Catalog
- [ ] Deploy backend
- [ ] Update database
- [ ] Verify functionality

### After Deployment
- [ ] Test all features
- [ ] Check logs
- [ ] Monitor performance
- [ ] Verify backups

---

## 💾 Backup & Restore

### Database Backup
```sql
BACKUP DATABASE EmployeeAssetManagement 
TO DISK = 'C:\Backups\EmployeeAssetManagement.bak' 
WITH INIT;
```

### Database Restore
```sql
RESTORE DATABASE EmployeeAssetManagement 
FROM DISK = 'C:\Backups\EmployeeAssetManagement.bak';
```

### Application Backup
```powershell
# Backup app folder
Compress-Archive -Path C:\apps\EmployeeAssetManagement -DestinationPath backup.zip
```

---

## 🔗 Important Links

### Microsoft Documentation
- [SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/)
- [React Documentation](https://reactjs.org)
- [ASP.NET Core](https://docs.microsoft.com/en-us/dotnet/core/aspnet/)
- [Entity Framework Core](https://docs.microsoft.com/en-us/ef/core/)
- [SQL Server](https://docs.microsoft.com/en-us/sql/sql-server/)

### Tools
- [Postman](https://www.postman.com/) - API testing
- [SQL Server Management Studio](https://docs.microsoft.com/en-us/sql/ssms/) - DB management
- [Visual Studio Code](https://code.visualstudio.com/) - Code editor
- [Visual Studio](https://visualstudio.microsoft.com/) - IDE

---

## 📱 Configuration Variables

### Environment Variables (.env)
```
ASPNETCORE_ENVIRONMENT=Production
DATABASE_CONNECTION_STRING=Server=...
JWT_SECRET=your-secret-key
API_BASE_URL=https://your-api.com/api
```

### appsettings.json
```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=...;Database=...;"
  },
  "Logging": {
    "LogLevel": {
      "Default": "Information"
    }
  }
}
```

---

## 🎯 Performance Tips

### Frontend
- Enable code splitting
- Use React.memo for components
- Lazy load non-critical components
- Minify assets

### Backend
- Use async/await
- Enable query caching
- Add database indexes
- Implement pagination

### Database
- Keep indexes optimized
- Regular maintenance
- Monitor query performance
- Archive old data

---

## 🆘 Getting Help

### Documentation
- README.md - Overview
- SETUP_GUIDE.md - Setup help
- API_REFERENCE.md - API details
- ARCHITECTURE.md - Design info
- DEPLOYMENT.md - Deploy help

### Common Questions
**Q: How do I change the API URL?**
A: Edit `EmployeeManagementWebPart.ts` line 25

**Q: How do I add a new field?**
A: Add to model → Create migration → Update UI

**Q: How do I deploy to production?**
A: See DEPLOYMENT.md for complete instructions

**Q: Where are the logs?**
A: Frontend: Browser console | Backend: Output console

---

## 📈 Monitoring

### Performance Metrics
- API response time
- Database query time
- Web part load time
- User count

### Health Checks
```bash
# API health
curl https://your-api.com/api/health

# Database connection
SELECT 1

# Web part status
Check browser console
```

---

## 🔄 Update Process

### Frontend Update
```bash
npm update
npm run build -- --production
npm run package-solution -- --production
# Upload new .sppkg to App Catalog
```

### Backend Update
```bash
dotnet build
dotnet publish -c Release
# Deploy new binaries
```

### Database Update
```bash
dotnet ef migrations add NewMigration
dotnet ef database update
```

---

**Last Updated**: 2024
**Version**: 1.0.0
**Status**: Quick Reference Complete ✅
