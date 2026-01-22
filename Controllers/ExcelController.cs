using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Text.Json;
using Microsoft.AspNetCore.Authorization;

namespace ExcelFilterApi.Controllers
{
    [ApiController]
    [Route("api/excel")]
    public class ExcelController : ControllerBase
    {
        [HttpPost("upload")]
        [AllowAnonymous]
        [Microsoft.AspNetCore.Authorization.Authorize]
        public async Task<IActionResult> UploadExcel([FromForm] ExcelUploadDto dto)
        {
            if (dto.File == null || dto.File.Length == 0)
                return BadRequest("No file uploaded");

            using var stream = new MemoryStream();
            await dto.File.CopyToAsync(stream);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(stream);
            var ws = package.Workbook.Worksheets[0];

            if (ws.Dimension == null)
                return BadRequest("Worksheet is empty");

            int rowCount = ws.Dimension.Rows;
            int colCount = ws.Dimension.Columns;

            var headers = new List<string>();
            for (int col = 1; col <= colCount; col++)
            {
                headers.Add(ws.Cells[1, col].Value?.ToString()?.Trim() ?? "");
            }

            var columnName = dto.ColumnName?.Trim();
            if (string.IsNullOrEmpty(columnName))
                return BadRequest("ColumnName is required");

            if (!headers.Contains(columnName))
                return BadRequest($"Column '{columnName}' not found");

            int columnIndex = headers.IndexOf(columnName) + 1;

            var filteredRows = new List<Dictionary<string, string>>();
            var originalRowIndexes = new List<int>();
            var rowImages = new Dictionary<int, List<string>>();

            for (int row = 2; row <= rowCount; row++)
            {
                var cellValue = ws.Cells[row, columnIndex].Value?.ToString()?.Trim() ?? "";

                bool includeRow = false;

                if (cellValue.Contains(dto.FilterValue) || cellValue.Contains("כולם"))
                    includeRow = true;

                if (includeRow && cellValue.Contains("חוץ") && cellValue.Contains(dto.FilterValue))
                    includeRow = false;

                if (!includeRow)
                    continue;

                var rowDict = new Dictionary<string, string>();
                for (int col = 1; col <= colCount; col++)
                {
                    rowDict[headers[col - 1]] =
                        ws.Cells[row, col].Value?.ToString()?.Trim() ?? "";
                }

                filteredRows.Add(rowDict);
                originalRowIndexes.Add(row);
            }

            foreach (var drawing in ws.Drawings)
            {
                if (drawing is ExcelPicture pic && pic.Image != null)
                {
                    try
                    {
                        int picRow = pic.From.Row + 1;
                        var imageBytes = pic.Image.ImageBytes;

                        if (imageBytes != null && imageBytes.Length > 0)
                        {
                            var base64Image = Convert.ToBase64String(imageBytes);

                            if (!rowImages.ContainsKey(picRow))
                                rowImages[picRow] = new List<string>();

                            rowImages[picRow].Add(base64Image);
                        }
                    }
                    catch { }
                }
            }

            var filteredRowsWithImages = new List<Dictionary<string, object>>();
            for (int i = 0; i < filteredRows.Count; i++)
            {
                var rowData = new Dictionary<string, object>();
                foreach (var kvp in filteredRows[i])
                {
                    rowData[kvp.Key] = kvp.Value;
                }

                int originalRow = originalRowIndexes[i];
                if (rowImages.ContainsKey(originalRow))
                {
                    rowData["_images"] = rowImages[originalRow];
                }

                filteredRowsWithImages.Add(rowData);
            }

            var newPackage = new ExcelPackage();
            var newSheet = newPackage.Workbook.Worksheets.Add("Filtered");

            newSheet.DefaultColWidth = 15;
            newSheet.DefaultRowHeight = 80;

            for (int i = 0; i < headers.Count; i++)
            {
                var cell = newSheet.Cells[1, i + 1];
                cell.Value = headers[i];
                cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                cell.Style.Font.Bold = true;
            }

            // נתונים
            for (int i = 0; i < filteredRows.Count; i++)
            {
                for (int j = 0; j < headers.Count; j++)
                {
                    var cell = newSheet.Cells[i + 2, j + 1];
                    cell.Value = filteredRows[i][headers[j]];
                    cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    cell.Style.WrapText = true; // עטיפת טקסט אם צריך
                }
            }

            const int IMAGE_WIDTH = 100;
            const int IMAGE_HEIGHT = 100;

            foreach (var drawing in ws.Drawings)
            {
                if (drawing is ExcelPicture pic && pic.Image != null)
                {
                    try
                    {
                        int picRow = pic.From.Row + 1;
                        int picCol = pic.From.Column + 1;

                        int matchIndex = originalRowIndexes.IndexOf(picRow);

                        if (matchIndex >= 0)
                        {
                            var imageBytes = pic.Image.ImageBytes;
                            if (imageBytes == null || imageBytes.Length == 0)
                                continue;

                            using var imgStream = new MemoryStream(imageBytes);

                            var newPic = newSheet.Drawings.AddPicture(
                                $"pic_{matchIndex}_{picCol}",
                                imgStream
                            );

                            newPic.SetPosition(
                                matchIndex + 1,
                                5,
                                picCol - 1,
                                5
                            );

                            newPic.SetSize(IMAGE_WIDTH, IMAGE_HEIGHT);

                            newPic.EditAs = OfficeOpenXml.Drawing.eEditAs.OneCell;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to copy image: {ex.Message}");
                        continue;
                    }
                }
            }

            var excelBytes = newPackage.GetAsByteArray();
            var json = JsonSerializer.Serialize(filteredRowsWithImages);

            return Ok(new
            {
                ExcelBase64 = Convert.ToBase64String(excelBytes),
                Json = json
            });
        }
    }

    public class ExcelUploadDto
    {
        public IFormFile File { get; set; } = default!;
        public string ColumnName { get; set; } = "";
        public string FilterValue { get; set; } = "";
    }
}