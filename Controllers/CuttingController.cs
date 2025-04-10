using System;
using Microsoft.AspNetCore.Mvc;
using MetalCuttingOptimizer.Models;
using MetalCuttingOptimizer.Services;
using Microsoft.Extensions.Logging;

namespace MetalCuttingOptimizer.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class CuttingController : ControllerBase
    {
        private readonly CuttingOptimizationService _optimizationService;
        private readonly ILogger<CuttingController> _logger;

        public CuttingController(CuttingOptimizationService optimizationService, ILogger<CuttingController> logger)
        {
            _optimizationService = optimizationService;
            _logger = logger;
        }

        [HttpPost("optimize")]
        public IActionResult OptimizeCutting([FromBody] OptimizationRequest request)
        {
            var billet = new SteelBillet
            {
                Length = request.BilletLength,
                Width = request.BilletWidth,
                Thickness = request.BilletThickness,
                Grade = request.SteelGrade ?? "DefaultGrade",
                CostPerUnit = request.CostPerUnit
            };

            var result = _optimizationService.OptimizeCutting(billet, request.RequiredLength, request.RequiredWidth);
            _logger.LogInformation($"Optimization completed for billet ID: {result.BilletId}");
            return Ok(result);
        }

        [HttpGet("export/{billetId}")]
        public IActionResult ExportToExcel(int billetId)
        {
            _logger.LogInformation($"Exporting to Excel for billet ID: {billetId}");
            var result = _optimizationService.GetCuttingResultById(billetId);

            if (result == null)
            {
                _logger.LogWarning($"No result found for billet ID: {billetId}");
                return NotFound();
            }

            var excelData = _optimizationService.ExportToExcel(result);
            return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "cutting_optimization.xlsx");
        }
    }

    public class OptimizationRequest
    {
        public double BilletLength { get; set; }
        public double BilletWidth { get; set; }
        public double BilletThickness { get; set; }
        public string? SteelGrade { get; set; }
        public decimal CostPerUnit { get; set; }
        public double RequiredLength { get; set; }
        public double RequiredWidth { get; set; }
    }
} 