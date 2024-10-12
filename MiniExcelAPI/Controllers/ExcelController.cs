using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MiniExcelLibs;

namespace MiniExcelAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly IWebHostEnvironment _environment;

        public ExcelController(IWebHostEnvironment environment)
        {
            _environment = environment;
        }

        // GET: api/excel/export
        [HttpGet]
        public async Task<IActionResult> ExportToExcel()
        {
            // Sample data for export
            var data = new List<dynamic>
            {
                new { Name = "Alice", Age = 30, Department = "HR" },
                new { Name = "Bob", Age = 25, Department = "IT" },
                new { Name = "Charlie", Age = 35, Department = "Finance" }
            };

            using (var stream = new MemoryStream())
            {
                // Export data to Excel using MiniExcel
                stream.SaveAs(data);

                // Return the Excel file as a download
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Employees.xlsx");
            }
        }

        [HttpPost]
        public async Task<IActionResult> ImportFromExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            try
            {
                using (var stream = new MemoryStream())
                {
                    await file.CopyToAsync(stream);
                    stream.Position = 0; // Reset position to the start of the stream

                    // Use MiniExcel to query the Excel file
                    var rows = stream.Query(useHeaderRow: true).ToList();

                    return Ok(rows);
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error importing Excel: {ex.Message}");
            }
        }

    }

}
