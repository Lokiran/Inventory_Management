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
    [Route("api/assets")]
    public class AssetsController : ControllerBase
    {
        private readonly InventoryContext _context;
        private readonly ILogger<AssetsController> _logger;

        public AssetsController(InventoryContext context, ILogger<AssetsController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Get assigned assets for an employee
        /// </summary>
        [HttpGet("assigned")]
        public async Task<ActionResult<IEnumerable<AssetAssignmentDto>>> GetAssignedAssets([FromQuery] string email)
        {
            try
            {
                var employee = await _context.Employees.FirstOrDefaultAsync(e => e.Email == email);
                if (employee == null)
                {
                    return NotFound("Employee not found");
                }

                var assignments = await _context.AssetAssignments
                    .Include(aa => aa.Asset)
                    .Where(aa => aa.EmployeeId == employee.EmployeeId && aa.Status == "Active")
                    .ToListAsync();

                var dtos = assignments.Select(a => new AssetAssignmentDto
                {
                    Id = a.AssignmentId.ToString(),
                    AssetType = a.Asset.AssetType,
                    AssetName = a.Asset.AssetName,
                    SerialNumber = a.Asset.SerialNumber,
                    AssignmentDate = a.AssignmentDate,
                    Status = a.Asset.Status,
                    Condition = a.Condition,
                    Location = a.Asset.Location
                }).ToList();

                return Ok(dtos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching assigned assets");
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Get available assets
        /// </summary>
        [HttpGet("available")]
        public async Task<ActionResult<IEnumerable<AssetDto>>> GetAvailableAssets()
        {
            try
            {
                var assets = await _context.Assets
                    .Where(a => a.Status == "Available")
                    .ToListAsync();

                var dtos = assets.Select(a => new AssetDto
                {
                    Id = a.AssetId,
                    AssetType = a.AssetType,
                    AssetName = a.AssetName,
                    SerialNumber = a.SerialNumber,
                    Status = a.Status,
                    Location = a.Location
                }).ToList();

                return Ok(dtos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching available assets");
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Get all assets
        /// </summary>
        [HttpGet]
        public async Task<ActionResult<IEnumerable<AssetDto>>> GetAllAssets()
        {
            try
            {
                var assets = await _context.Assets.ToListAsync();

                var dtos = assets.Select(a => new AssetDto
                {
                    Id = a.AssetId,
                    AssetType = a.AssetType,
                    AssetName = a.AssetName,
                    SerialNumber = a.SerialNumber,
                    Status = a.Status,
                    Location = a.Location
                }).ToList();

                return Ok(dtos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching assets");
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Get a specific asset
        /// </summary>
        [HttpGet("{id}")]
        public async Task<ActionResult<AssetDto>> GetAsset(int id)
        {
            try
            {
                var asset = await _context.Assets.FindAsync(id);
                if (asset == null)
                {
                    return NotFound();
                }

                var dto = new AssetDto
                {
                    Id = asset.AssetId,
                    AssetType = asset.AssetType,
                    AssetName = asset.AssetName,
                    SerialNumber = asset.SerialNumber,
                    Status = asset.Status,
                    Location = asset.Location
                };

                return Ok(dto);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching asset");
                return StatusCode(500, "Internal server error");
            }
        }
    }

    // DTOs
    public class AssetDto
    {
        public int Id { get; set; }
        public string AssetType { get; set; }
        public string AssetName { get; set; }
        public string SerialNumber { get; set; }
        public string Status { get; set; }
        public string? Location { get; set; }
    }

    public class AssetAssignmentDto
    {
        public string Id { get; set; }
        public string AssetType { get; set; }
        public string AssetName { get; set; }
        public string SerialNumber { get; set; }
        public DateTime AssignmentDate { get; set; }
        public string Status { get; set; }
        public string Condition { get; set; }
        public string? Location { get; set; }
    }
}
