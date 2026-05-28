# Employee Asset Management System - Setup Guide

## Table of Contents
1. [Overview](#overview)
2. [Prerequisites](#prerequisites)
3. [Frontend Setup (SharePoint Framework)](#frontend-setup-sharepoint-framework)
4. [Backend Setup (.NET Core)](#backend-setup-net-core)
5. [Database Setup (SQL Server)](#database-setup-sql-server)
6. [Deployment](#deployment)
7. [Configuration](#configuration)
8. [API Documentation](#api-documentation)
9. [Troubleshooting](#troubleshooting)

## Overview

The Employee Asset Management System is a comprehensive solution for managing employee asset requests, incidents, and assignments. It consists of:

- **Frontend**: React-based SharePoint Framework (SPFx) Web Part
- **Backend**: ASP.NET Core REST API
- **Database**: SQL Server

### Features

✅ Employee Dashboard with statistics
✅ Asset Request Management
✅ Incident Reporting and Tracking
✅ Assigned Assets Management
✅ Request History and Downloads
✅ Role-based Access Control
✅ Email Notifications
✅ Audit Logging
✅ Dashboard Analytics

## Prerequisites

### Frontend Development
- Node.js (v22.14.0 or higher)
- npm or yarn
- Visual Studio Code or Visual Studio 2022
- SharePoint Online Tenant
- SPFx CLI tools

### Backend Development
- .NET 8 SDK or higher
- Visual Studio 2022 or Visual Studio Code
- SQL Server 2019 or higher (or Azure SQL Database)
- Postman (for API testing)

### General
- Git
- Admin access to SharePoint tenant
- Access to Azure DevOps or similar for CI/CD (optional)

## Frontend Setup (SharePoint Framework)

### Step 1: Install SPFx Tools

```bash
# Install Yeoman and SPFx generator globally
npm install -g yo @microsoft/generator-sharepoint

# Or if already installed, update
npm update -g yo @microsoft/generator-sharepoint
```

### Step 2: Create Web Part Project

```bash
# Navigate to your project directory
cd /path/to/INVENTORY

# Create the web part (if not already done)
yo @microsoft/sharepoint

# When prompted:
# - SharePoint Framework version: 1.22.2
# - Which type of client-side component?: Web part
# - What is your Web part name?: EmployeeManagement
# - Which framework would you like to use?: React
```

### Step 3: Install Dependencies

```bash
# In the project root directory
npm install

# Install additional required packages
npm install chart.js react-chartjs-2
npm install @pnp/sp @pnp/logging
```

### Step 4: Build and Package

```bash
# Build the project
npm run build

# Bundle and package for SharePoint
npm run package-solution -- --production
```

The SharePoint package will be generated at:
```
sharepoint/solution/spfx-project.sppkg
```

### Step 5: Deploy Web Part

1. Go to your SharePoint App Catalog:
   ```
   https://[tenant]-admin.sharepoint.com/sites/appcatalog
   ```

2. Upload the `.sppkg` file from `sharepoint/solution/`

3. Trust the app when prompted

4. Make the app available to your organization or specific site collections

### Step 6: Add Web Part to Page

1. Go to your target SharePoint site
2. Create or edit a page
3. Add the web part:
   - Click "+"
   - Search for "Employee Management"
   - Add to page

4. Configure the web part:
   - Enter API Base URL: `https://your-api.com/api`
   - Save

## Backend Setup (.NET Core)

### Step 1: Create API Project

```bash
# Navigate to Backend folder
cd Backend

# Create new web API project (if not already done)
dotnet new webapi -n EmployeeAssetManagement.Api

# Or open existing project in Visual Studio
```

### Step 2: Install NuGet Packages

```bash
# Navigate to API project directory
cd EmployeeAssetManagement.Api

# Install required packages
dotnet add package Microsoft.EntityFrameworkCore
dotnet add package Microsoft.EntityFrameworkCore.SqlServer
dotnet add package Swashbuckle.AspNetCore
dotnet add package System.IdentityModel.Tokens.Jwt
dotnet add package Microsoft.AspNetCore.Authentication.JwtBearer

# Restore packages
dotnet restore
```

### Step 3: Configure Database Connection

Edit `appsettings.json`:

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=YOUR_SERVER;Database=EmployeeAssetManagement;User Id=sa;Password=YOUR_PASSWORD;Encrypt=false;TrustServerCertificate=true;"
  }
}
```

Replace:
- `YOUR_SERVER`: Your SQL Server instance (e.g., `localhost` or `server.database.windows.net`)
- `YOUR_PASSWORD`: Your SQL Server password

### Step 4: Build and Run

```bash
# Build the project
dotnet build

# Run migrations (creates database)
dotnet ef database update

# Run the API
dotnet run

# API will be available at: https://localhost:5001/swagger
```

### Step 5: Test API Endpoints

Open Swagger UI at:
```
https://localhost:5001/swagger/index.html
```

Test endpoints:
- GET `/api/dashboard/stats?email=user@example.com`
- POST `/api/asset-requests`
- GET `/api/asset-requests/employee?email=user@example.com`
- POST `/api/incidents`
- GET `/api/assets/assigned?email=user@example.com`

## Database Setup (SQL Server)

### Step 1: Create Database

Option A: Using SQL Server Management Studio

1. Open SQL Server Management Studio
2. Connect to your SQL Server instance
3. Right-click "Databases" → "New Database"
4. Name: `EmployeeAssetManagement`
5. Click "OK"

Option B: Using T-SQL Script

```bash
# Using sqlcmd
sqlcmd -S YOUR_SERVER -U sa -P YOUR_PASSWORD -i Database\DatabaseSchema.sql
```

### Step 2: Run Database Schema Script

1. Open SQL Server Management Studio
2. Connect to your server
3. Open `Backend\Database\DatabaseSchema.sql`
4. Execute the script

This will create:
- All required tables
- Indexes for performance
- Sample data
- Stored procedures

### Step 3: Verify Database

```sql
-- Run this query to verify
USE EmployeeAssetManagement;
SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES;
```

You should see:
- Employees
- Assets
- AssetRequests
- Incidents
- AssetAssignments
- AuditLogs

## Deployment

### Frontend Deployment to Production

```bash
# Build for production
npm run build -- --production

# Package solution
npm run package-solution -- --production

# Upload to App Catalog and trust
```

### Backend Deployment to Azure App Service

```bash
# Publish to Azure
dotnet publish -c Release -o ./publish

# Using Azure CLI
az webapp deployment source config-zip --resource-group YOUR_RG --name YOUR_APP_NAME --src ./publish.zip
```

### Or Deploy to IIS (On-Premises)

1. Publish the project:
```bash
dotnet publish -c Release -o ./publish
```

2. Copy `publish` folder to your IIS server

3. Create new Application in IIS:
   - App Name: EmployeeAssetManagement
   - Physical Path: Path to `publish` folder
   - Application Pool: Create new (.NET Core)

4. Configure binding (HTTPS recommended)

5. Create `web.config` for IIS

## Configuration

### Environment Variables

Create `.env` file in API root:

```
ASPNETCORE_ENVIRONMENT=Production
DatabaseConnectionString=Server=YOUR_SERVER;Database=EmployeeAssetManagement;User Id=sa;Password=YOUR_PASSWORD
JWT_SECRET=your-secret-key-here
JWT_ISSUER=your-issuer
JWT_AUDIENCE=your-audience
```

### CORS Configuration

In `Program.cs`, update CORS policy:

```csharp
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowSharePoint", builder =>
        builder.WithOrigins("https://[tenant].sharepoint.com")
               .AllowAnyMethod()
               .AllowAnyHeader());
});
```

### Email Notifications

Configure email service in `appsettings.json`:

```json
{
  "EmailSettings": {
    "SmtpServer": "smtp.company.com",
    "SmtpPort": 587,
    "FromAddress": "noreply@company.com",
    "Password": "your-password",
    "EnableSSL": true
  }
}
```

## API Documentation

### Base URL
```
https://your-api.com/api
```

### Authentication
All requests require Bearer token in Authorization header:
```
Authorization: Bearer {token}
```

### Endpoints

#### Dashboard
```
GET /dashboard/stats?email=user@example.com
Response: {
  "totalRequests": 10,
  "pendingRequests": 2,
  "approvedRequests": 5,
  "resolvedIncidents": 8,
  "openIncidents": 1
}
```

#### Asset Requests
```
POST /asset-requests
Body: {
  "employeeEmail": "user@example.com",
  "assetType": "Laptop",
  "assetName": "Dell XPS 13",
  "quantity": 1,
  "priority": "High",
  "reasonDescription": "Replacement for damaged device",
  "requiredDate": "2024-01-15"
}

GET /asset-requests/employee?email=user@example.com
GET /asset-requests/{id}
PUT /asset-requests/{id}/cancel
PUT /asset-requests/{id}/approve
```

#### Incidents
```
POST /incidents (multipart/form-data)
FormData: {
  "employeeEmail": "user@example.com",
  "assetId": "DELL-001",
  "issueType": "Hardware Damage",
  "issueDescription": "Screen is cracked",
  "priority": "High",
  "attachment": [file]
}

GET /incidents/employee?email=user@example.com
GET /incidents/{id}
PUT /incidents/{id}/status
```

#### Assets
```
GET /assets
GET /assets/available
GET /assets/assigned?email=user@example.com
GET /assets/{id}
```

## Troubleshooting

### Frontend Issues

**Problem**: Web part not showing in SharePoint
- Solution: Clear browser cache, hard refresh (Ctrl+F5)
- Check if app is deployed to correct site collection
- Verify API URL is correctly configured

**Problem**: API calls failing with CORS error
- Solution: Check CORS configuration in backend
- Verify API URL matches domain in CORS policy

**Problem**: Build errors
```bash
# Clear node_modules and reinstall
rm -rf node_modules
npm install

# Clear cache
npm cache clean --force
npm run clean
```

### Backend Issues

**Problem**: Database connection error
- Verify connection string in `appsettings.json`
- Check SQL Server is running and accessible
- Verify firewall rules allow connection

**Problem**: Migrations not running
```bash
# Remove and reapply migrations
dotnet ef database drop
dotnet ef database update

# Or manually run schema script
sqlcmd -S YOUR_SERVER -U sa -P YOUR_PASSWORD -i Database\DatabaseSchema.sql
```

**Problem**: Port 5001 already in use
```bash
# Change port in launchSettings.json
# Or use different port:
dotnet run --urls "https://localhost:5002"
```

### Database Issues

**Problem**: Login failed for user 'sa'
- Verify SQL Server authentication is enabled
- Check user password is correct
- For Azure SQL: Check firewall rules allow your IP

**Problem**: Table already exists
- Drop database and recreate:
```sql
DROP DATABASE EmployeeAssetManagement;
-- Then run DatabaseSchema.sql again
```

## Security Best Practices

1. **Authentication**: Implement JWT with short expiration times
2. **Authorization**: Enforce role-based access control
3. **HTTPS**: Always use HTTPS in production
4. **Secrets**: Store secrets in Azure Key Vault or similar
5. **Database**: Use strong passwords, limit user permissions
6. **Input Validation**: Validate all user inputs
7. **Logging**: Enable audit logging for compliance
8. **Updates**: Keep frameworks and packages updated

## Performance Optimization

- Add database query caching
- Implement pagination for large datasets
- Use async/await for all I/O operations
- Configure database connection pooling
- Enable GZIP compression in IIS
- Use CDN for static assets

## Support & Maintenance

For support, contact:
- **Frontend**: SharePoint Team
- **Backend**: Development Team
- **Database**: Database Administrator

Regular maintenance:
- Monitor API performance
- Review audit logs weekly
- Backup database daily
- Update dependencies monthly
- Test disaster recovery quarterly

---

Last Updated: 2024
Version: 1.0.0
