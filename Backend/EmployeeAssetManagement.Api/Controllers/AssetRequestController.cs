using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using EmployeeAssetManagement.Api.Data;
using EmployeeAssetManagement.Api.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EmployeeAssetManagement.Api.Controllers
{
    [ApiController]
    [Route("api/asset-requests")]
    public class AssetRequestController : ControllerBase
    {
        private readonly InventoryContext _context;
        private readonly ILogger<AssetRequestController> _logger;

        public AssetRequestController(InventoryContext context, ILogger<AssetRequestController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Create a new asset request
        /// </summary>
        [HttpPost]
        public async Task<ActionResult<AssetRequest>> CreateAssetRequest([FromBody] CreateAssetRequestDto dto)
        {
            try
            {
                var employee = await _context.Employees.FirstOrDefaultAsync(e => e.Email == dto.EmployeeEmail);
                if (employee == null)
                {
                    return NotFound("Employee not found");
                }

                var assetRequest = new AssetRequest
                {
                    EmployeeId = employee.EmployeeId,
                    AssetType = dto.AssetType,
                    AssetName = dto.AssetName,
                    Quantity = dto.Quantity,
                    Priority = dto.Priority,
                    ReasonDescription = dto.ReasonDescription,
                    RequiredDate = dto.RequiredDate,
                    RequestDate = DateTime.UtcNow,
                    Status = "Pending"
                };

                _context.AssetRequests.Add(assetRequest);
                await _context.SaveChangesAsync();

                await LogAudit(employee.EmployeeId, "Created Asset Request", "AssetRequests", null, assetRequest.RequestId.ToString());

                return CreatedAtAction(nameof(GetAssetRequest), new { id = assetRequest.RequestId }, assetRequest);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating asset request");
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Get asset requests for an employee
        /// </summary>
        [HttpGet("employee")]
        public async Task<ActionResult<IEnumerable<AssetRequestDto>>> GetEmployeeAssetRequests([FromQuery] string email)
        {
            try
            {
                var requests = await _context.AssetRequests
                    .Include(ar => ar.Employee)
                    .Where(ar => ar.Employee.Email == email)
                    .OrderByDescending(ar => ar.RequestDate)
                    .ToListAsync();

                var dtos = requests.Select(r => new AssetRequestDto
                {
                    Id = r.RequestId.ToString(),
                    AssetType = r.AssetType,
                    AssetName = r.AssetName,
                    Quantity = r.Quantity,
                    Priority = r.Priority,
                    Status = r.Status,
                    RequestDate = r.RequestDate,
                    RequiredDate = r.RequiredDate,
                    Description = r.ReasonDescription
                }).ToList();

                return Ok(dtos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching employee asset requests");
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Get a specific asset request
        /// </summary>
        [HttpGet("{id}")]
        public async Task<ActionResult<AssetRequest>> GetAssetRequest(int id)
        {
            var assetRequest = await _context.AssetRequests
                .Include(ar => ar.Employee)
                .FirstOrDefaultAsync(ar => ar.RequestId == id);

            if (assetRequest == null)
            {
                return NotFound();
            }

            return Ok(assetRequest);
        }

        /// <summary>
        /// Cancel an asset request
        /// </summary>
        [HttpPut("{id}/cancel")]
        public async Task<IActionResult> CancelAssetRequest(int id)
        {
            try
            {
                var assetRequest = await _context.AssetRequests.FindAsync(id);
                if (assetRequest == null)
                {
                    return NotFound();
                }

                if (assetRequest.Status != "Pending")
                {
                    return BadRequest("Only pending requests can be cancelled");
                }

                assetRequest.Status = "Cancelled";
                assetRequest.ModifiedDate = DateTime.UtcNow;

                _context.AssetRequests.Update(assetRequest);
                await _context.SaveChangesAsync();

                await LogAudit(assetRequest.EmployeeId, "Cancelled Asset Request", "AssetRequests", null, id.ToString());

                return Ok("Request cancelled successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error cancelling asset request");
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Approve an asset request (Admin/Manager only)
        /// </summary>
        [HttpPut("{id}/approve")]
        public async Task<IActionResult> ApproveAssetRequest(int id, [FromBody] ApproveAssetRequestDto dto)
        {
            try
            {
                var assetRequest = await _context.AssetRequests.FindAsync(id);
                if (assetRequest == null)
                {
                    return NotFound();
                }

                assetRequest.Status = "Approved";
                assetRequest.ApprovedDate = DateTime.UtcNow;
                assetRequest.ApprovalComment = dto.Comment;
                assetRequest.ApprovedBy = dto.ApprovedBy;
                assetRequest.ModifiedDate = DateTime.UtcNow;

                _context.AssetRequests.Update(assetRequest);
                await _context.SaveChangesAsync();

                await LogAudit(dto.ApprovedBy, "Approved Asset Request", "AssetRequests", null, id.ToString());

                return Ok("Request approved successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error approving asset request");
                return StatusCode(500, "Internal server error");
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
    public class CreateAssetRequestDto
    {
        public string EmployeeEmail { get; set; }
        public string AssetType { get; set; }
        public string AssetName { get; set; }
        public int Quantity { get; set; }
        public string Priority { get; set; }
        public string? ReasonDescription { get; set; }
        public DateTime RequiredDate { get; set; }
    }

    public class AssetRequestDto
    {
        public string Id { get; set; }
        public string AssetType { get; set; }
        public string AssetName { get; set; }
        public int Quantity { get; set; }
        public string Priority { get; set; }
        public string Status { get; set; }
        public DateTime RequestDate { get; set; }
        public DateTime RequiredDate { get; set; }
        public string? Description { get; set; }
    }

    public class ApproveAssetRequestDto
    {
        public int ApprovedBy { get; set; }
        public string? Comment { get; set; }
    }
}
