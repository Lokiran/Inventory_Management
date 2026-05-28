-- ===================================================================
-- Employee Asset Management Database Schema
-- ===================================================================

-- Drop existing database if needed (for fresh setup)
-- DROP DATABASE IF EXISTS EmployeeAssetManagement;

-- Create Database
CREATE DATABASE IF NOT EXISTS EmployeeAssetManagement;
GO

USE EmployeeAssetManagement;
GO

-- ===================================================================
-- Create Tables
-- ===================================================================

-- Employees Table
CREATE TABLE dbo.Employees (
    EmployeeId INT PRIMARY KEY IDENTITY(1,1),
    Name NVARCHAR(MAX) NOT NULL,
    Email NVARCHAR(MAX) NOT NULL UNIQUE,
    Department NVARCHAR(MAX) NOT NULL,
    Role NVARCHAR(MAX) NOT NULL, -- Employee, Manager, Admin
    Phone NVARCHAR(20) NULL,
    CreatedDate DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    ModifiedDate DATETIME2 NULL
);

-- Assets Table
CREATE TABLE dbo.Assets (
    AssetId INT PRIMARY KEY IDENTITY(1,1),
    AssetType NVARCHAR(MAX) NOT NULL, -- Laptop, Mouse, Keyboard, Monitor, Headset, Mobile
    AssetName NVARCHAR(MAX) NOT NULL,
    Description NVARCHAR(MAX) NULL,
    SerialNumber NVARCHAR(MAX) NOT NULL UNIQUE,
    Status NVARCHAR(MAX) NOT NULL, -- Available, Assigned, Damaged, Retired
    Location NVARCHAR(MAX) NULL,
    PurchasePrice DECIMAL(10, 2) NULL,
    PurchaseDate DATETIME2 NULL,
    CreatedDate DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    ModifiedDate DATETIME2 NULL
);

-- AssetRequests Table
CREATE TABLE dbo.AssetRequests (
    RequestId INT PRIMARY KEY IDENTITY(1,1),
    EmployeeId INT NOT NULL,
    AssetId INT NULL,
    AssetType NVARCHAR(MAX) NOT NULL,
    AssetName NVARCHAR(MAX) NOT NULL,
    Quantity INT NOT NULL,
    Priority NVARCHAR(MAX) NOT NULL, -- Low, Medium, High, Urgent
    Status NVARCHAR(MAX) NOT NULL, -- Pending, Approved, Rejected, Issued, Cancelled
    ReasonDescription NVARCHAR(MAX) NULL,
    RequiredDate DATETIME2 NOT NULL,
    RequestDate DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    ApprovedDate DATETIME2 NULL,
    ApprovalComment NVARCHAR(MAX) NULL,
    ApprovedBy INT NULL,
    ModifiedDate DATETIME2 NULL,
    FOREIGN KEY (EmployeeId) REFERENCES dbo.Employees(EmployeeId),
    FOREIGN KEY (AssetId) REFERENCES dbo.Assets(AssetId),
    FOREIGN KEY (ApprovedBy) REFERENCES dbo.Employees(EmployeeId)
);

-- Incidents Table
CREATE TABLE dbo.Incidents (
    IncidentId INT PRIMARY KEY IDENTITY(1,1),
    IncidentNumber NVARCHAR(MAX) NOT NULL,
    EmployeeId INT NOT NULL,
    AssetId INT NOT NULL,
    IssueType NVARCHAR(MAX) NOT NULL, -- Hardware Damage, Software Issue, Not Working, Missing Parts, Performance, Other
    IssueDescription NVARCHAR(MAX) NOT NULL,
    Priority NVARCHAR(MAX) NOT NULL, -- Low, Medium, High, Critical
    Status NVARCHAR(MAX) NOT NULL, -- Open, In Progress, Resolved, Closed
    AttachmentUrl NVARCHAR(MAX) NULL,
    ReportedDate DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    ResolvedDate DATETIME2 NULL,
    ResolutionNotes NVARCHAR(MAX) NULL,
    AssignedTo INT NULL,
    CreatedDate DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    ModifiedDate DATETIME2 NULL,
    FOREIGN KEY (EmployeeId) REFERENCES dbo.Employees(EmployeeId),
    FOREIGN KEY (AssetId) REFERENCES dbo.Assets(AssetId),
    FOREIGN KEY (AssignedTo) REFERENCES dbo.Employees(EmployeeId)
);

-- AssetAssignments Table
CREATE TABLE dbo.AssetAssignments (
    AssignmentId INT PRIMARY KEY IDENTITY(1,1),
    EmployeeId INT NOT NULL,
    AssetId INT NOT NULL,
    Condition NVARCHAR(MAX) NOT NULL, -- Good, Fair, Poor
    Notes NVARCHAR(MAX) NULL,
    AssignmentDate DATETIME2 NOT NULL,
    ReturnDate DATETIME2 NULL,
    Status NVARCHAR(MAX) NOT NULL, -- Active, Returned
    CreatedDate DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    ModifiedDate DATETIME2 NULL,
    FOREIGN KEY (EmployeeId) REFERENCES dbo.Employees(EmployeeId),
    FOREIGN KEY (AssetId) REFERENCES dbo.Assets(AssetId)
);

-- AuditLogs Table
CREATE TABLE dbo.AuditLogs (
    LogId INT PRIMARY KEY IDENTITY(1,1),
    EmployeeId INT NOT NULL,
    Action NVARCHAR(MAX) NOT NULL,
    TableName NVARCHAR(MAX) NOT NULL,
    OldValues NVARCHAR(MAX) NULL,
    NewValues NVARCHAR(MAX) NULL,
    CreatedDate DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    IPAddress NVARCHAR(50) NULL,
    FOREIGN KEY (EmployeeId) REFERENCES dbo.Employees(EmployeeId)
);

-- ===================================================================
-- Create Indexes
-- ===================================================================

CREATE INDEX idx_Employees_Email ON dbo.Employees(Email);
CREATE INDEX idx_Assets_SerialNumber ON dbo.Assets(SerialNumber);
CREATE INDEX idx_Assets_Status ON dbo.Assets(Status);
CREATE INDEX idx_AssetRequests_EmployeeId ON dbo.AssetRequests(EmployeeId);
CREATE INDEX idx_AssetRequests_Status ON dbo.AssetRequests(Status);
CREATE INDEX idx_AssetRequests_RequestDate ON dbo.AssetRequests(RequestDate);
CREATE INDEX idx_Incidents_EmployeeId ON dbo.Incidents(EmployeeId);
CREATE INDEX idx_Incidents_Status ON dbo.Incidents(Status);
CREATE INDEX idx_Incidents_ReportedDate ON dbo.Incidents(ReportedDate);
CREATE INDEX idx_AssetAssignments_EmployeeId ON dbo.AssetAssignments(EmployeeId);
CREATE INDEX idx_AssetAssignments_Status ON dbo.AssetAssignments(Status);
CREATE INDEX idx_AuditLogs_CreatedDate ON dbo.AuditLogs(CreatedDate);

-- ===================================================================
-- Insert Sample Data (Optional)
-- ===================================================================

-- Insert sample employees
INSERT INTO dbo.Employees (Name, Email, Department, Role, CreatedDate)
VALUES 
    ('John Doe', 'john.doe@company.com', 'IT', 'Employee', GETUTCDATE()),
    ('Jane Smith', 'jane.smith@company.com', 'HR', 'Manager', GETUTCDATE()),
    ('Admin User', 'admin@company.com', 'IT', 'Admin', GETUTCDATE());

-- Insert sample assets
INSERT INTO dbo.Assets (AssetType, AssetName, SerialNumber, Status, Location, CreatedDate)
VALUES 
    ('Laptop', 'Dell XPS 13', 'DELL-001', 'Available', 'IT Store', GETUTCDATE()),
    ('Mouse', 'Logitech MX Master', 'LOG-001', 'Available', 'IT Store', GETUTCDATE()),
    ('Keyboard', 'Mechanical RGB', 'KEY-001', 'Assigned', 'Desk 101', GETUTCDATE()),
    ('Monitor', 'LG UltraWide', 'LG-001', 'Available', 'IT Store', GETUTCDATE()),
    ('Headset', 'Sony WH-1000XM5', 'SONY-001', 'Available', 'IT Store', GETUTCDATE());

GO

-- ===================================================================
-- Create Stored Procedures (Optional)
-- ===================================================================

-- Get Employee Dashboard Summary
CREATE PROCEDURE sp_GetEmployeeDashboardSummary
    @EmployeeEmail NVARCHAR(MAX)
AS
BEGIN
    SELECT 
        COUNT(CASE WHEN ar.Status = 'Pending' THEN 1 END) AS PendingRequests,
        COUNT(CASE WHEN ar.Status = 'Approved' THEN 1 END) AS ApprovedRequests,
        COUNT(CASE WHEN i.Status IN ('Open', 'In Progress') THEN 1 END) AS OpenIncidents,
        COUNT(CASE WHEN i.Status IN ('Resolved', 'Closed') THEN 1 END) AS ResolvedIncidents,
        COUNT(DISTINCT aa.AssetId) AS AssignedAssets
    FROM dbo.Employees e
    LEFT JOIN dbo.AssetRequests ar ON e.EmployeeId = ar.EmployeeId
    LEFT JOIN dbo.Incidents i ON e.EmployeeId = i.EmployeeId
    LEFT JOIN dbo.AssetAssignments aa ON e.EmployeeId = aa.EmployeeId AND aa.Status = 'Active'
    WHERE e.Email = @EmployeeEmail;
END;

GO

-- Get Pending Approval Requests (for Managers)
CREATE PROCEDURE sp_GetPendingApprovals
    @ManagerEmail NVARCHAR(MAX)
AS
BEGIN
    SELECT 
        ar.RequestId,
        ar.AssetType,
        ar.AssetName,
        ar.Quantity,
        ar.Priority,
        ar.RequestDate,
        ar.RequiredDate,
        ar.ReasonDescription,
        e.Name AS EmployeeName,
        e.Email AS EmployeeEmail,
        e.Department
    FROM dbo.AssetRequests ar
    JOIN dbo.Employees e ON ar.EmployeeId = e.EmployeeId
    WHERE ar.Status = 'Pending'
    ORDER BY ar.RequestDate ASC;
END;

GO

-- Get Asset Utilization Report
CREATE PROCEDURE sp_GetAssetUtilizationReport
AS
BEGIN
    SELECT 
        a.AssetType,
        COUNT(a.AssetId) AS TotalAssets,
        SUM(CASE WHEN a.Status = 'Available' THEN 1 ELSE 0 END) AS AvailableCount,
        SUM(CASE WHEN a.Status = 'Assigned' THEN 1 ELSE 0 END) AS AssignedCount,
        SUM(CASE WHEN a.Status = 'Damaged' THEN 1 ELSE 0 END) AS DamagedCount,
        CAST(SUM(CASE WHEN a.Status = 'Assigned' THEN 1 ELSE 0 END) AS FLOAT) / 
        NULLIF(COUNT(a.AssetId), 0) * 100 AS UtilizationPercentage
    FROM dbo.Assets a
    GROUP BY a.AssetType
    ORDER BY a.AssetType;
END;

GO

PRINT 'Database schema created successfully!'
