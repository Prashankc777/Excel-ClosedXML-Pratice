﻿using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ClosedXML_Pratice.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ClosedXMLExcelSheet : ControllerBase
    {
        public ClosedXMLExcelSheet()
        {
            
        }

        public class Student
        {
            public int Id { get; set; }
            public string Name { get; set; } = string.Empty;
            public string RollNumber { get; set; } = string.Empty;
        }

        [HttpGet, Route("download-excel-report"), AllowAnonymous]
        public async Task<IActionResult> DownloadExcelReport()
        {
            try
            {
                List<Student> students = new List<Student>();

                for (int i = 0; i < 50; i++)
                {
                    students.Add(new Student()
                    {
                        Id = i,
                        Name = "Prashan" + i,
                        RollNumber = "100" + i
                    });
                }

                using var workBook = new XLWorkbook();
                //Adding the workSheet
                var ws = workBook.Worksheets.Add("Student");
                var ws01 = workBook.Worksheets.Add("Grade");
                var ws02 = workBook.Worksheets.Add("Grade01", 2); // pachadi ko number chai position

                //Getting worksheet name 
                IXLWorksheet worksheet = workBook.Worksheet(ws.Name);
                var xx = worksheet.Name;

                var currentRow = 1;
                ws.Cell(currentRow, 1).Value = "StudentId";
                ws.Cell(currentRow, 2).Value = "Name";
                ws.Cell(currentRow, 3).Value = "Roll";
                ws.Cell(currentRow, 4).Value = "Int";

                currentRow++;

                ws.Range(ws.Cell(currentRow - 1, 1), ws.Cell(currentRow ++, 1)).Merge();
                ws.Range(ws.Cell(currentRow - 1, 1), ws.Cell(currentRow, 1)).Value = "Merge";



                //Suru ko thulo heading yo ho 
                var temp = ws01.Range(ws01.Cell(1, 1), ws01.Cell(1, 10));
                temp.Merge();
                temp.Value = "Hello";
                temp.Style.Font.Bold = true;
                temp.Style.Font.FontSize = 50;
                temp.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);



                foreach (var student in students)
                {
                    currentRow++;
                    ws.Cell(currentRow, 1).Value = student.Id;
                    //Alignment in center
                    ws.Cell(currentRow, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    ws.Cell(currentRow, 2).Value = student.Name;
                    ws.Cell(currentRow, 3).Value = student.RollNumber;
                    ws.Cell(currentRow, 4).Value = 2;
                    ws.Cell(currentRow, 4).SetDataType(XLDataType.Number);
                }

                //single work sheet color
                ws.Cell(2, 3).Style.Fill.SetBackgroundColor(XLColor.Cyan);

                //Range ma color lagaune
                IXLRange range01 = ws.Range(ws.Cell(4, 2).Address, ws.Cell(6, 4).Address);
                //range01.Style.Fill.SetBackgroundColor(XLColor.Cyan);
                //range01.Style.Fill.SetBackgroundColor(XLColor.FromHtml("#FF996515"));


                //lining in all excel sheet 
                IXLRange range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(students.Count + 1, 3).Address);
                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                range.Style.Border.RightBorder = XLBorderStyleValues.Thin;

                //Auto fit column 
                ws.Columns().AdjustToContents();

                await using var stream = new MemoryStream();
                workBook.SaveAs(stream);
                var content = stream.ToArray();
                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "Students.xlsx"
                );
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);
            }
        }
    }
}
