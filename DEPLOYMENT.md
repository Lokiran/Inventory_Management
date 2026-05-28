# Deployment Guide

## Table of Contents
1. [Pre-Deployment Checklist](#pre-deployment-checklist)
2. [Frontend Deployment](#frontend-deployment)
3. [Backend Deployment](#backend-deployment)
4. [Database Deployment](#database-deployment)
5. [Post-Deployment Verification](#post-deployment-verification)
6. [Rollback Procedures](#rollback-procedures)
7. [Monitoring & Maintenance](#monitoring--maintenance)

## Pre-Deployment Checklist

### Code Quality
- [ ] All tests passing
- [ ] Code reviewed and approved
- [ ] No console errors/warnings
- [ ] Security scan completed
- [ ] Performance benchmarks acceptable

### Configuration
- [ ] API URLs configured correctly
- [ ] Database connection strings verified
- [ ] Secrets stored in Key Vault/secure location
- [ ] CORS policies configured
- [ ] Email settings configured

### Documentation
- [ ] README updated
- [ ] API documentation current
- [ ] Setup guide accurate
- [ ] Known issues documented

### Backup
- [ ] Database backed up
- [ ] Previous version packaged
- [ ] Rollback plan documented

---

## Frontend Deployment

### Step 1: Build for Production

```bash
# Navigate to project root
cd INVENTORY

# Install latest dependencies
npm install

# Run production build
npm run build -- --production

# Generate optimized package
npm run package-solution -- --production
```

This creates:
- `sharepoint/solution/spfx-project.sppkg` - Production package

### Step 2: Upload to App Catalog

1. **Navigate to App Catalog**
   ```
   https://[tenant]-admin.sharepoint.com/sites/appcatalog
   ```

2. **Upload the package**
   - Click "New" → "Files"
   - Upload `spfx-project.sppkg`

3. **Trust the app**
   - Click on the uploaded file
   - In the ribbon, click "Files" → "Details"
   - Check "Make this solution available to all sites in the organization"
   - Click "Save"

4. **Verify app availability**
   - Wait 5-10 minutes for propagation
   - Go to a SharePoint site
   - Try adding the web part

### Step 3: Add Web Part to Target Sites

**Option A: Manual - Via Site Pages**
1. Go to target SharePoint site
2. Create or edit a page
3. Click "+"
4. Search for "Employee Management"
5. Add web part to page
6. Configure API URL in web part settings

**Option B: Automated - Via PowerShell**
```powershell
# Connect to SharePoint
Connect-SPOService -Url https://[tenant]-admin.sharepoint.com

# Add web part to site
Add-SPOAppInstance -App $AppInstance -Site https://[tenant].sharepoint.com/sites/[site]

# Configure web part settings
Set-SPOClientSidePage -Identity "page.aspx" -LayoutType Home
```

### Step 4: Verify Frontend Deployment

1. Navigate to the page with the web part
2. Verify web part loads without errors
3. Test navigation and basic functionality
4. Check browser console for errors (F12)
5. Test on different browsers (Chrome, Edge, Safari)

---

## Backend Deployment

### Option A: Deploy to Azure App Service

#### Step 1: Create App Service (if not exists)

```powershell
# Using Azure CLI
az group create --name myResourceGroup --location eastus

az appservice plan create `
  --name myAppServicePlan `
  --resource-group myResourceGroup `
  --sku B2

az webapp create `
  --resource-group myResourceGroup `
  --plan myAppServicePlan `
  --name EmployeeAssetManagementApi
```

#### Step 2: Publish to Azure

```bash
# Publish the application
dotnet publish -c Release -o ./publish

# Create deployment package
Compress-Archive -Path ./publish/* -DestinationPath api.zip

# Deploy using Azure CLI
az webapp deployment source config-zip `
  --resource-group myResourceGroup `
  --name EmployeeAssetManagementApi `
  --src api.zip
```

#### Step 3: Configure App Settings in Azure

```powershell
az webapp config appsettings set `
  --name EmployeeAssetManagementApi `
  --resource-group myResourceGroup `
  --settings `
  ConnectionStrings__DefaultConnection="Server=your-sql-server.database.windows.net;Database=EmployeeAssetManagement;User Id=sa;Password=your-password;" `
  ASPNETCORE_ENVIRONMENT="Production"
```

#### Step 4: Enable HTTPS

1. Go to Azure Portal
2. Navigate to your App Service
3. Under "Settings" → "SSL settings"
4. Add SSL certificate (or use App Service Managed Certificate)
5. Enable "HTTPS Only"

### Option B: Deploy to IIS (On-Premises)

#### Step 1: Prepare Server

```powershell
# Install .NET Hosting Bundle
# Download from: https://dotnet.microsoft.com/download/dotnet/

# Install IIS (if not present)
Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServer, IIS-WebServerRole
```

#### Step 2: Publish Application

```bash
# Publish to folder
dotnet publish -c Release -o C:\apps\EmployeeAssetManagement

# Set permissions
icacls "C:\apps\EmployeeAssetManagement" /grant:r "IIS_IUSRS:(OI)(CI)(F)"
```

#### Step 3: Create IIS Application

1. Open **IIS Manager**
2. Expand server, then "Sites"
3. Right-click "Default Web Site" → "Add Application"
4. Configure:
   - **Alias**: `api`
   - **Physical path**: `C:\apps\EmployeeAssetManagement`
   - **Application pool**: Create new (No Managed Code)

5. Set Application Pool settings:
   - .NET CLR version: "No Managed Code"
   - Managed pipeline mode: "Integrated"

#### Step 4: Configure SSL/HTTPS

1. Right-click site → "Edit Bindings"
2. Add HTTPS binding (port 443)
3. Select SSL certificate
4. Set HTTP to redirect to HTTPS

#### Step 5: Create web.config

```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <system.webServer>
    <handlers>
      <add name="aspNetCore" path="*" verb="*" modules="AspNetCoreModule" resourceType="Unspecified" />
    </handlers>
    <aspNetCore processPath="dotnet" arguments=".\EmployeeAssetManagement.Api.dll" stdoutLogEnabled="false" stdoutLogFile=".\logs\stdout" hostingModel="inprocess" />
  </system.webServer>
</configuration>
```

### Step 6: Verify Backend Deployment

```bash
# Test API
curl -X GET https://your-api.com/api/assets -H "Authorization: Bearer {token}"

# Check Swagger UI
https://your-api.com/swagger

# View application logs
# Windows: Event Viewer → Windows Logs → Application
# Azure: App Service → Log stream
```

---

## Database Deployment

### Step 1: Create Database

**Option A: Azure SQL Database**

```powershell
# Create SQL Server
az sql server create `
  --name my-sql-server `
  --resource-group myResourceGroup `
  --admin-user sqladmin `
  --admin-password MyPassword123!

# Create database
az sql db create `
  --resource-group myResourceGroup `
  --server my-sql-server `
  --name EmployeeAssetManagement `
  --service-objective S0
```

**Option B: On-Premises SQL Server**

```sql
-- Create database
CREATE DATABASE EmployeeAssetManagement;
GO

-- Create login
CREATE LOGIN appuser WITH PASSWORD = 'YourPassword123!';
GO

-- Create user
USE EmployeeAssetManagement;
CREATE USER appuser FOR LOGIN appuser;
GRANT CONTROL ON DATABASE::EmployeeAssetManagement TO appuser;
GO
```

### Step 2: Run Database Schema Script

```bash
# Using sqlcmd
sqlcmd -S your-server -U sa -P your-password -i Backend/Database/DatabaseSchema.sql

# Or via Entity Framework (from API project)
dotnet ef database update
```

### Step 3: Verify Database

```sql
-- Connect to database
USE EmployeeAssetManagement;
GO

-- Check tables
SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES;

-- Verify indexes
SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID('dbo.Employees');

-- Check sample data
SELECT COUNT(*) FROM dbo.Employees;
SELECT COUNT(*) FROM dbo.Assets;
```

### Step 4: Configure Backup (Optional)

**Azure SQL Database (Automatic)**
- Automatic backups included in Azure SQL
- Check backup retention in Portal

**On-Premises SQL Server**
```sql
-- Create backup job
BACKUP DATABASE EmployeeAssetManagement
TO DISK = 'C:\Backups\EmployeeAssetManagement.bak'
WITH INIT, COMPRESSION;
GO

-- Schedule daily backup
-- Use SQL Server Agent or Windows Task Scheduler
```

---

## Post-Deployment Verification

### Frontend Verification Checklist

```powershell
# ✅ Web part loads
# ✅ Navigation works
# ✅ Dashboard displays stats
# ✅ Asset request form functional
# ✅ Incident form works
# ✅ File uploads work
# ✅ Download functionality works
# ✅ No JavaScript errors
# ✅ Mobile responsive
# ✅ Charts render correctly
```

### Backend Verification Checklist

```powershell
# ✅ API responds to requests
# ✅ Swagger UI accessible
# ✅ Authentication working
# ✅ Database operations successful
# ✅ Logs being recorded
# ✅ Audit logs created
# ✅ Error handling working
# ✅ CORS configured correctly
# ✅ Rate limiting (if enabled)
```

### Database Verification Checklist

```powershell
# ✅ Database accessible
# ✅ All tables present
# ✅ Indexes created
# ✅ Sample data loaded
# ✅ Stored procedures functional
# ✅ Backup completed
# ✅ Query performance acceptable
```

### Load Testing

```bash
# Using Apache Bench
ab -n 1000 -c 10 https://your-api.com/api/assets

# Using curl loop
for i in {1..100}; do
  curl -X GET https://your-api.com/api/dashboard/stats?email=user@example.com
done
```

---

## Rollback Procedures

### Frontend Rollback

```powershell
# Revert to previous version in App Catalog
1. Go to App Catalog
2. Find the app
3. Click "Details" → "Version History"
4. Select previous version
5. Click "Restore"
```

### Backend Rollback

**Azure App Service:**
```powershell
# Use Deployment Slots
az webapp deployment slot create `
  --resource-group myResourceGroup `
  --name EmployeeAssetManagementApi `
  --slot staging

# Test in staging, then swap
az webapp deployment slot swap `
  --resource-group myResourceGroup `
  --name EmployeeAssetManagementApi `
  --slot staging

# Rollback by swapping again
az webapp deployment slot swap `
  --resource-group myResourceGroup `
  --name EmployeeAssetManagementApi `
  --slot staging
```

**IIS Server:**
```powershell
# Stop application
Stop-WebAppPool -Name "EmployeeAssetManagement"

# Restore previous version
Copy-Item -Path "C:\apps\EmployeeAssetManagement_Backup\*" `
          -Destination "C:\apps\EmployeeAssetManagement" `
          -Recurse -Force

# Start application
Start-WebAppPool -Name "EmployeeAssetManagement"
```

### Database Rollback

```sql
-- Restore from backup
RESTORE DATABASE EmployeeAssetManagement
FROM DISK = 'C:\Backups\EmployeeAssetManagement_2024_01_01.bak'
WITH REPLACE;
GO
```

---

## Monitoring & Maintenance

### Application Health Monitoring

```powershell
# Azure Monitor
az monitor metrics list `
  --resource /subscriptions/{id}/resourceGroups/myResourceGroup/providers/Microsoft.Web/sites/EmployeeAssetManagementApi `
  --metric "HealthCheckStatus" `
  --start-time 2024-01-01T00:00:00Z

# IIS Performance Monitor
perfmon.exe

# Application Insights
az monitor app-insights component create `
  --app MyApplicationInsights `
  --location eastus `
  --kind web `
  --resource-group myResourceGroup
```

### Log Analysis

```bash
# View application logs
# Azure
az webapp log tail --name EmployeeAssetManagementApi --resource-group myResourceGroup

# IIS
Get-EventLog -LogName Application -Source "EmployeeAssetManagement"

# Docker
docker logs container-name
```

### Performance Optimization

1. **Database Optimization**
   - Review slow queries
   - Add missing indexes
   - Update statistics

2. **API Optimization**
   - Enable response caching
   - Implement pagination
   - Use compression

3. **Frontend Optimization**
   - Enable code splitting
   - Implement lazy loading
   - Minify assets

### Scheduled Maintenance

**Daily:**
- Check error logs
- Monitor disk space
- Verify backups

**Weekly:**
- Review performance metrics
- Analyze audit logs
- Update dependencies (if needed)

**Monthly:**
- Security patch assessment
- Performance optimization
- Capacity planning

**Quarterly:**
- Disaster recovery testing
- Security audit
- Infrastructure review

---

## Emergency Procedures

### API is Down

1. Check service status
   ```bash
   curl -I https://your-api.com/api/health
   ```

2. Review logs
   ```powershell
   # Azure
   az webapp log tail --name EmployeeAssetManagementApi
   
   # IIS
   Get-EventLog -LogName Application -Newest 100
   ```

3. Restart service
   ```powershell
   # Azure
   az webapp restart --name EmployeeAssetManagementApi
   
   # IIS
   Restart-WebAppPool -Name "EmployeeAssetManagement"
   ```

4. If issue persists, rollback

### Database Connection Issues

1. Verify connection string
2. Check database status
3. Test connectivity
   ```bash
   sqlcmd -S your-server -U sa -P your-password -Q "SELECT @@VERSION"
   ```
4. Restart SQL Server if necessary

### High Disk Space Usage

1. Check what's using space
   ```powershell
   Get-ChildItem -Path "C:\" -Recurse -Force | 
   Group-Object -Property Directory | 
   Sort-Object -Property @{Expression={$_.Group | Measure-Object -Property Length -Sum | Select-Object -ExpandProperty Sum}} -Descending
   ```

2. Clean up old logs/backups
3. Expand disk if needed

---

## Deployment Automation (CI/CD)

### Azure DevOps Pipeline Example

```yaml
trigger:
  - main

pool:
  vmImage: 'windows-latest'

variables:
  buildConfiguration: 'Release'

stages:
- stage: Build
  jobs:
  - job: BuildJob
    steps:
    - task: UseDotNet@2
      inputs:
        packageType: 'sdk'
        version: '8.x'

    - task: DotNetCoreCLI@2
      inputs:
        command: 'build'
        arguments: '--configuration $(buildConfiguration)'

    - task: DotNetCoreCLI@2
      inputs:
        command: 'publish'
        publishWebProjects: true
        arguments: '--configuration $(buildConfiguration) --output $(Build.ArtifactStagingDirectory)'

    - task: PublishBuildArtifacts@1

- stage: Deploy
  dependsOn: Build
  condition: succeeded()
  jobs:
  - deployment: DeployJob
    environment: 'Production'
    strategy:
      runOnce:
        deploy:
          steps:
          - task: AzureWebApp@1
            inputs:
              azureSubscription: 'MyAzureSubscription'
              appType: 'webAppLinux'
              appName: 'EmployeeAssetManagementApi'
              package: '$(Pipeline.Workspace)/drop/api.zip'
```

---

Last Updated: 2024
Version: 1.0.0
