using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using EmployeeAssetManagement.Api.Data;
using EmployeeAssetManagement.Api.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace EmployeeAssetManagement.Api.Controllers
{
    [ApiController]
    [Route("api/incidents")]
    public class IncidentController : ControllerBase
    {
        private readonly InventoryContext _context;
        private readonly ILogger<IncidentController> _logger;

        public IncidentController(InventoryContext context, ILogger<IncidentController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Create a new incident
        /// </summary>
        [HttpPost]
        public async Task<ActionResult<Incident>> CreateIncident([FromForm] CreateIncidentDto dto)
        {
            try
            {
                var employee = await _context.Employees.FirstOrDefaultAsync(e => e.Email == dto.EmployeeEmail);
                if (employee == null)
                {
                    return NotFound("Employee not found");
                }

                var asset = await _context.Assets.FirstOrDefaultAsync(a => a.SerialNumber == dto.AssetId);
                if (asset == null)
                {
                    return NotFound("Asset not found");
                }

                string? attachmentUrl = null;
                if (dto.Attachment != null)
                {
                    attachmentUrl = await SaveAttachment(dto.Attachment);
                }

                var incident = new Incident
                {
                    IncidentNumber = GenerateIncidentNumber(),
                    EmployeeId = employee.EmployeeId,
                    AssetId = asset.AssetId,
                    IssueType = dto.IssueType,
                    IssueDescription = dto.IssueDescription,
                    Priority = dto.Priority,
                    Status = "Open",
                    AttachmentUrl = attachmentUrl,
                    ReportedDate = DateTime.UtcNow,
                    CreatedDate = DateTime.UtcNow
                };

                _context.Incidents.Add(incident);
                await _context.SaveChangesAsync();

                await LogAudit(employee.EmployeeId, "Created Incident", "Incidents", null, incident.IncidentNumber);

                return CreatedAtAction(nameof(GetIncident), new { id = incident.IncidentId }, incident);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating incident");
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Get incidents for an employee
        /// </summary>
        [HttpGet("employee")]
        public async Task<ActionResult<IEnumerable<IncidentDto>>> GetEmployeeIncidents([FromQuery] string email)
        {
            try
            {
                var incidents = await _context.Incidents
                    .Include(i => i.Employee)
                    .Include(i => i.Asset)
                    .Where(i => i.Employee.Email == email)
                    .OrderByDescending(i => i.ReportedDate)
                    .ToListAsync();

                var dtos = incidents.Select(i => new IncidentDto
                {
                    Id = i.IncidentId.ToString(),
                    IncidentId = i.IncidentNumber,
                    AssetId = i.Asset.SerialNumber,
                    AssetName = i.Asset.AssetName,
                    IssueType = i.IssueType,
                    IssueDescription = i.IssueDescription,
                    Priority = i.Priority,
                    Status = i.Status,
                    ReportedDate = i.ReportedDate,
                    ResolvedDate = i.ResolvedDate,
                    AssignedTo = i.AssignedToEmployee?.Name,
                    Resolution = i.ResolutionNotes
                }).ToList();

                return Ok(dtos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching employee incidents");
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Get a specific incident
        /// </summary>
        [HttpGet("{id}")]
        public async Task<ActionResult<Incident>> GetIncident(int id)
        {
            var incident = await _context.Incidents
                .Include(i => i.Employee)
                .Include(i => i.Asset)
                .FirstOrDefaultAsync(i => i.IncidentId == id);

            if (incident == null)
            {
                return NotFound();
            }

            return Ok(incident);
        }

        /// <summary>
        /// Update incident status
        /// </summary>
        [HttpPut("{id}/status")]
        public async Task<IActionResult> UpdateIncidentStatus(int id, [FromBody] UpdateIncidentStatusDto dto)
        {
            try
            {
                var incident = await _context.Incidents.FindAsync(id);
                if (incident == null)
                {
                    return NotFound();
                }

                incident.Status = dto.Status;
                if (dto.Status == "Resolved" || dto.Status == "Closed")
                {
                    incident.ResolvedDate = DateTime.UtcNow;
                }
                incident.ResolutionNotes = dto.ResolutionNotes;
                incident.ModifiedDate = DateTime.UtcNow;

                _context.Incidents.Update(incident);
                await _context.SaveChangesAsync();

                await LogAudit(incident.EmployeeId, "Updated Incident Status", "Incidents", null, incident.IncidentNumber);

                return Ok("Incident updated successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating incident");
                return StatusCode(500, "Internal server error");
            }
        }

        private string GenerateIncidentNumber()
        {
            var timestamp = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
            return $"INC-{timestamp}";
        }

        private async Task<string> SaveAttachment(IFormFile file)
        {
            try
            {
                var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "uploads");
                if (!Directory.Exists(uploadsFolder))
                {
                    Directory.CreateDirectory(uploadsFolder);
                }

                var fileName = $"{Guid.NewGuid()}_{file.FileName}";
                var filePath = Path.Combine(uploadsFolder, fileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                return $"/uploads/{fileName}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving attachment");
                return null;
            }
        }

        private async Task LogAudit(int employeeId, string action, string tableName, string? oldValues, string? newValues)
        {
            var auditLog = new AuditLog
            {
                EmployeeId = employeeId,
                Action = action,
                TableName = tableName,
                OldValues = oldValues,
                NewValues = newValues,
                CreatedDate = DateTime.UtcNow,
                IPAddress = HttpContext.Connection.RemoteIpAddress?.ToString()
            };

            _context.AuditLogs.Add(auditLog);
            await _context.SaveChangesAsync();
        }
    }

    // DTOs
    public class CreateIncidentDto
    {
        public string EmployeeEmail { get; set; }
        public string AssetId { get; set; }
        public string IssueType { get; set; }
        public string IssueDescription { get; set; }
        public string Priority { get; set; }
        public IFormFile? Attachment { get; set; }
        public DateTime ReportedDate { get; set; }
    }

    public class IncidentDto
    {
        public string Id { get; set; }
        public string IncidentId { get; set; }
        public string AssetId { get; set; }
        public string AssetName { get; set; }
        public string IssueType { get; set; }
        public string IssueDescription { get; set; }
        public string Priority { get; set; }
        public string Status { get; set; }
        public DateTime ReportedDate { get; set; }
        public DateTime? ResolvedDate { get; set; }
        public string? AssignedTo { get; set; }
        public string? Resolution { get; set; }
    }

    public class UpdateIncidentStatusDto
    {
        public string Status { get; set; }
        public string? ResolutionNotes { get; set; }
    }
}
