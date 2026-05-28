namespace EmployeeAssetManagement.Api.Models
{
    public class Employee
    {
        public int EmployeeId { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string Department { get; set; }
        public string Role { get; set; } // Employee, Manager, Admin
        public string? Phone { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }

        public virtual ICollection<AssetRequest> AssetRequests { get; set; }
        public virtual ICollection<Incident> Incidents { get; set; }
        public virtual ICollection<AssetAssignment> AssetAssignments { get; set; }
    }

    public class Asset
    {
        public int AssetId { get; set; }
        public string AssetType { get; set; } // Laptop, Mouse, Keyboard, Monitor, Headset, Mobile
        public string AssetName { get; set; }
        public string? Description { get; set; }
        public string SerialNumber { get; set; }
        public string Status { get; set; } // Available, Assigned, Damaged, Retired
        public string? Location { get; set; }
        public decimal? PurchasePrice { get; set; }
        public DateTime? PurchaseDate { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }

        public virtual ICollection<AssetRequest> AssetRequests { get; set; }
        public virtual ICollection<AssetAssignment> AssetAssignments { get; set; }
        public virtual ICollection<Incident> Incidents { get; set; }
    }

    public class AssetRequest
    {
        public int RequestId { get; set; }
        public int EmployeeId { get; set; }
        public int? AssetId { get; set; }
        public string AssetType { get; set; }
        public string AssetName { get; set; }
        public int Quantity { get; set; }
        public string Priority { get; set; } // Low, Medium, High, Urgent
        public string Status { get; set; } // Pending, Approved, Rejected, Issued
        public string? ReasonDescription { get; set; }
        public DateTime RequiredDate { get; set; }
        public DateTime RequestDate { get; set; }
        public DateTime? ApprovedDate { get; set; }
        public string? ApprovalComment { get; set; }
        public int? ApprovedBy { get; set; }
        public DateTime? ModifiedDate { get; set; }

        public virtual Employee Employee { get; set; }
        public virtual Asset? Asset { get; set; }
        public virtual Employee? ApprovedByEmployee { get; set; }
    }

    public class Incident
    {
        public int IncidentId { get; set; }
        public string IncidentNumber { get; set; } // For user reference
        public int EmployeeId { get; set; }
        public int AssetId { get; set; }
        public string IssueType { get; set; } // Hardware Damage, Software Issue, Not Working, etc.
        public string IssueDescription { get; set; }
        public string Priority { get; set; } // Low, Medium, High, Critical
        public string Status { get; set; } // Open, In Progress, Resolved, Closed
        public string? AttachmentUrl { get; set; }
        public DateTime ReportedDate { get; set; }
        public DateTime? ResolvedDate { get; set; }
        public string? ResolutionNotes { get; set; }
        public int? AssignedTo { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }

        public virtual Employee Employee { get; set; }
        public virtual Asset Asset { get; set; }
        public virtual Employee? AssignedToEmployee { get; set; }
    }

    public class AssetAssignment
    {
        public int AssignmentId { get; set; }
        public int EmployeeId { get; set; }
        public int AssetId { get; set; }
        public string Condition { get; set; } // Good, Fair, Poor
        public string? Notes { get; set; }
        public DateTime AssignmentDate { get; set; }
        public DateTime? ReturnDate { get; set; }
        public string Status { get; set; } // Active, Returned
        public DateTime CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }

        public virtual Employee Employee { get; set; }
        public virtual Asset Asset { get; set; }
    }

    public class AuditLog
    {
        public int LogId { get; set; }
        public int EmployeeId { get; set; }
        public string Action { get; set; }
        public string TableName { get; set; }
        public string? OldValues { get; set; }
        public string? NewValues { get; set; }
        public DateTime CreatedDate { get; set; }
        public string? IPAddress { get; set; }

        public virtual Employee Employee { get; set; }
    }
}
