using System.IO;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using Newtonsoft.Json;

[ApiController]
[Route("api/[controller]")]
public class ExcelController : ControllerBase
{
    [HttpPost("upload")]
    public IActionResult Upload(IFormFile file)
    {
        if (file == null || file.Length == 0)
            return BadRequest("No file uploaded.");

        try
        {
            using var stream = new MemoryStream();
            file.CopyTo(stream);

            using var package = new ExcelPackage(stream);
            var worksheet = package.Workbook.Worksheets[0];

            var rowCount = worksheet.Dimension.Rows;
            var colCount = worksheet.Dimension.Columns;

            var data = new List<Dictionary<string, string>>();

            // Read header row
            var headers = new List<string>();
            for (var col = 1; col <= colCount; col++)
            {
                headers.Add(worksheet.Cells[1, col].Text);
            }

            // Read data rows
            for (var row = 2; row <= rowCount; row++)
            {
                var rowData = new Dictionary<string, string>();
                for (var col = 1; col <= colCount; col++)
                {
                    rowData[headers[col - 1]] = worksheet.Cells[row, col].Text;
                }
                data.Add(rowData);
            }

            return Ok(data);
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"Internal server error: {ex.Message}");
        }
    }
}
