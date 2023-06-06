using ExcelMVC.Models;
using System.Data;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace ExcelMVC.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _hostingEnvironment;

        public HomeController(IWebHostEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Upload(IFormFile file)
        {
            if (file != null && file.Length > 0)
            {
                var filePath = Path.Combine(_hostingEnvironment.WebRootPath, "uploads", file.FileName);
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                try
                {
                    var dt = ReadExcelToDataTable(filePath);
                    return View(dt);
                }
                catch (Exception ex)
                {
                    ViewBag.ErrorMessage = "Error occurred while processing the Excel file: " + ex.Message;
                    System.IO.File.Delete(filePath); // If it's not an ExcelSheet Delete the uploaded file
                    return View("Index");
                }
            }

            return RedirectToAction("Index");
        }

        private DataTable ReadExcelToDataTable(string filePath)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets[0];
                    var dt = new DataTable();

                    foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns])
                    {
                        dt.Columns.Add(firstRowCell.Text);
                    }

                    for (var rowNum = 2; rowNum <= worksheet.Dimension.Rows; rowNum++)
                    {
                        var worksheetRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.Columns];
                        var row = dt.Rows.Add();
                        foreach (var cell in worksheetRow)
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                    }

                    return dt;
                }
            }
            catch (Exception ex)
            {
                // Handle the EPPlus exception (e.g., log, display an error message, etc.)
                Console.WriteLine("Invalid Excel file: " + ex.Message);
                throw new ApplicationException("Error occurred while reading Excel file: " + ex.Message);
            }
        }
    }
 
}
