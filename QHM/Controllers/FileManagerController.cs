using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.IO;

namespace QHM.Controllers
{

    [ApiController]
    [Route("[controller]")]
    public class FileManagerController : Controller
    {
        [HttpPost("/createfile")]
        public IActionResult CreateFile(Data data )
        {
            var memoryStream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(memoryStream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Backbone");

                // Đặt tiêu đề cột
                worksheet.Cells[1, 1].Value = "Index";
                worksheet.Cells[1, 2].Value = "X";
                worksheet.Cells[1, 3].Value = "Y";
                worksheet.Cells[1, 4].Value = "Weight";
                worksheet.Cells[1, 5].Value = "IsBackBone";
                worksheet.Cells[1, 6].Value = "Moment";
                worksheet.Cells[1, 7].Value = "IsAccess";
                worksheet.Cells[1, 8].Value = "Backbone Distance";
                worksheet.Cells[1, 9].Value = "Backbone Index";

                // Ghi dữ liệu từ danh sách vào file Excel
                int row = 2;
                foreach (var item in data.MyPoints)
                {
                    worksheet.Cells[row, 1].Value = (item.index + 1).ToString();

                    worksheet.Cells[row, 2].Value = item.x.ToString();
                    worksheet.Cells[row, 3].Value= item.y.ToString();
                    worksheet.Cells[row, 4].Value = item.W.ToString();
                    if (item.IsBackbone)
                        worksheet.Cells[row, 5].Value = "Yes";
                    else
                        worksheet.Cells[row, 5].Value = "No";
                    if (item.IsBackbone)
                        worksheet.Cells[row, 6].Value = item.moment.ToString();
                    else
                        worksheet.Cells[row, 6].Value = "N/A";

                    if (item.IsAccess)
                        worksheet.Cells[row, 7].Value = "Yes";
                    else
                        worksheet.Cells[row, 7].Value = "No";

                    if (item.IsAccess)
                        worksheet.Cells[row, 8].Value = item.backboneDistance.ToString();
                    else
                        worksheet.Cells[row, 8].Value = "N/A";
                    if (item.IsAccess)
                        worksheet.Cells[row, 9].Value = (item.backbone.index + 1).ToString();
                    else
                        worksheet.Cells[row, 9].Value = "N/A";

                    row++;
                }
                var backboneList = data.MyPoints.Where(it => it.IsBackbone).ToList();

                ExcelWorksheet worksheet2 = package.Workbook.Worksheets.Add("Lưu lượng giữa các nút backbone");
                worksheet2.Cells[1, 1].Value = "T_b";
                for (int i = 2; i< backboneList.Count + 2; i++)
                {
                    worksheet2.Cells[1, i].Value = backboneList[i - 2].index;
                }
                row = 2;
                foreach(var item in data.Tb) {
                    worksheet2.Cells[row, 1].Value = backboneList[row - 2].index;
                    for (int i = 2;i < item.Count() + 2;i++)
                    {
                        worksheet2.Cells[row, i].Value = item[i - 2];
                    }
                    row++;
                }
                package.SaveAs(memoryStream); // Lưu file Excel
            }

            memoryStream.Position = 0;
            return new FileContentResult(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = "Sample.xlsx"
            };
        }
        [HttpPost("/test")]
        public IActionResult test()
        {
            return Ok();
        }
    }
}
