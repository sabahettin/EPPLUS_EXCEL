using EPPLUS_EXCEL.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Drawing;

namespace EPPLUS_EXCEL.Controllers
{
    public class EmployeeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult CreateExcelFile()
        {
            // EPPlus lisansımızı ayarlıyoruz.
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //Örnek listemizi ayarlıyoruz.
            var employees = new List<Employee>
            {
                new Employee{Id=1,Name="ALİ",SurName="VELİ",Email="ali.veli@mail.com" },
                new Employee{Id=2,Name="AHMET",SurName="YILMAZ",Email="ahmet.yilmaz@mail.com" },
                new Employee{Id=3,Name="SEDA",SurName="NUR",Email="seda.nur@mail.com" },
            };
            //Excel işlemlerimizi ve tablomuz için özel stillerimizi ayarıyoruz.
            var stream = new MemoryStream();
            using (var excelPackage = new ExcelPackage(stream))
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("Employess");
                worksheet.Cells["A1"].Value = "Employee List";
                using (var t = worksheet.Cells["A1:C1"])
                {
                    t.Merge = true;
                    t.Style.Font.Color.SetColor(Color.White);
                    t.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                    t.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    t.Style.Fill.BackgroundColor.SetColor(Color.Green);
                }
                worksheet.Cells["A3"].Value = "ID";
                worksheet.Cells["B3"].Value = "NAME";
                worksheet.Cells["C3"].Value = "SURNAME";
                worksheet.Cells["D3"].Value = "MAIL";
                worksheet.Cells["A3:D3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells["A3:D3"].Style.Fill.BackgroundColor.SetColor(Color.Red);
                worksheet.Cells["A3:D3"].Style.Font.Bold = true;

                int row = 4;
                foreach (var item in employees)
                {
                    worksheet.Cells[row, 1].Value = item.Id;
                    worksheet.Cells[row, 2].Value = item.Name;
                    worksheet.Cells[row, 3].Value = item.SurName;
                    worksheet.Cells[row, 4].Value = item.Email;
                    row++;
                }
                excelPackage.Save();
                stream.Position = 0;
                return File(stream, "application/vnd.opemxmlformats-officedocument.spreadsheetml.sheet", "employees.xlsx");
            }
        }

        public IActionResult ReadExcelFile()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ReadExcelFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                // Dosya seçilmediğinde veya boş olduğunda hata mesajı gösterilebilir.
                ViewBag.ErrorMessage = "Dosya Seçiniz.";
                return View();
            }

            // EPPlus lisansımızı ayarlıyoruz.
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Seçilen dosyanın bellekte tutulması
            using (var stream = new MemoryStream())
            {
                file.CopyTo(stream);
                stream.Position = 0;

                ExcelPackage excelPackage = new ExcelPackage(stream);
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();

                int rows = worksheet.Dimension.Rows;
                int columns = worksheet.Dimension.Columns;

                var employees = new List<Employee>();
                for (int i = 4; i <= rows; i++)
                {
                    var employee = new Employee();
                    employee.Id = int.Parse(worksheet.Cells[i, 1].Value.ToString());
                    employee.Name = worksheet.Cells[i, 2].Value.ToString();
                    employee.SurName = worksheet.Cells[i, 3].Value.ToString();
                    employee.Email = worksheet.Cells[i, 4].Value.ToString();

                    employees.Add(employee);
                }

                return View(employees);
            }
        }

    }
}
