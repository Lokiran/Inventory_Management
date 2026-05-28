using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using EmployeeAssetManagement.Api.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EmployeeAssetManagement.Api.Controllers
{
    [ApiController]
    [Route("api/dashboard")]
    public class DashboardController : ControllerBase
    {
        private readonly InventoryContext _context;
        private readonly ILogger<DashboardController> _logger;

        public DashboardController(InventoryContext context, ILogger<DashboardController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Get dashboard statistics for an employee
        /// </summary>
        [HttpGet("stats")]
        public async Task<ActionResult<DashboardStatsDto>> GetDashboardStats([FromQuery] string email)
        {
            try
            {
                var employee = await _context.Employees.FirstOrDefaultAsync(e => e.Email == email);
                if (employee == null)
                {
                    return NotFound("Employee not found");
                }

                var totalRequests = await _context.AssetRequests
                    .Where(ar => ar.EmployeeId == employee.EmployeeId && ar.Status != "Cancelled")
                    .CountAsync();

                var pendingRequests = await _context.AssetRequests
                    .Where(ar => ar.EmployeeId == employee.EmployeeId && ar.Status == "Pending")
                    .CountAsync();

                var approvedRequests = await _context.AssetRequests
                    .Where(ar => ar.EmployeeId == employee.EmployeeId && ar.Status == "Approved")
                    .CountAsync();

                var openIncidents = await _context.Incidents
                    .Where(i => i.EmployeeId == employee.EmployeeId && (i.Status == "Open" || i.Status == "In Progress"))
                    .CountAsync();

                var resolvedIncidents = await _context.Incidents
                    .Where(i => i.EmployeeId == employee.EmployeeId && (i.Status == "Resolved" || i.Status == "Closed"))
                    .CountAsync();

                var stats = new DashboardStatsDto
                {
                    TotalRequests = totalRequests,
                    PendingRequests = pendingRequests,
                    ApprovedRequests = approvedRequests,
                    ResolvedIncidents = resolvedIncidents,
                    OpenIncidents = openIncidents
                };

                return Ok(stats);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching dashboard stats");
                return StatusCode(500, "Internal server error");
            }
        }
    }

    public class DashboardStatsDto
    {
        public int TotalRequests { get; set; }
        public int PendingRequests { get; set; }
        public int ApprovedRequests { get; set; }
        public int ResolvedIncidents { get; set; }
        public int OpenIncidents { get; set; }
    }
}
