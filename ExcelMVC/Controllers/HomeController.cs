using ExcelMVC.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Data;
using System.IO;

namespace ExcelMVC.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _hostingEnvironment;

        public HomeController(IWebHostEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        // GET: Home/Index
        public IActionResult Index()
        {
            return View();
        }

        // POST: Home/Upload
        [HttpPost]
        public IActionResult Upload(IFormFile file)
        {
            var uploadsFolderPath = Path.Combine(_hostingEnvironment.WebRootPath, "uploads");
            if (!Directory.Exists(uploadsFolderPath))
            {
                Directory.CreateDirectory(uploadsFolderPath);
            }

            // Check if a file was uploaded
            if (file != null && file.Length > 0)
            {
                // Combine the file path with the uploads folder path
                var filePath = Path.Combine(_hostingEnvironment.WebRootPath, "uploads", file.FileName);

                try
                {
                    // Save the uploaded file to disk
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }
                }
                catch (UnauthorizedAccessException ex)
                {
                    // Handle the UnauthorizedAccessException
                    Console.WriteLine("Access to the path is denied: " + ex.Message);
                    // Displays a user-friendly error message when the file path is denied, etc.
                }

                try
                {
                    // Read the Excel file into a DataTable
                    var dt = ReadExcelToDataTable(filePath);
                    return View(dt); // Display the DataTable on the view
                }
                catch (Exception ex)
                {
                    ViewBag.ErrorMessage = "Error occurred while processing the Excel file: " + ex.Message;
                    System.IO.File.Delete(filePath); // Delete the uploaded file in case of an exception
                    return View("Index"); // Return to the Index view
                }
            }

            return RedirectToAction("Index"); // No file uploaded, redirect to the Index view
        }

        // Read the Excel file into a DataTable
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

                    // Create columns in the DataTable based on the first row of the worksheet
                    foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns])
                    {
                        dt.Columns.Add(firstRowCell.Text);
                    }

                    // Iterate through the rows of the worksheet and add them to the DataTable
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

        [HttpPost]
        public IActionResult ExportToExcel([FromBody] List<List<string>> data)
        {
            if (data != null && data.Count > 0)
            {
                using (var package = new ExcelPackage())
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets.Add("Sheet1");

                    // Write the data to the worksheet
                    for (int row = 0; row < data.Count; row++)
                    {
                        for (int col = 0; col < data[row].Count; col++)
                        {
                            worksheet.Cells[row + 1, col + 1].Value = data[row][col];
                        }
                    }

                    // Generate the Excel file bytes
                    var fileBytes = package.GetAsByteArray();

                    // Set the content type and file name for the response
                    var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    var fileName = "ExcelData.xlsx";

                    // Return the Excel file as a download response
                    return File(fileBytes, contentType, fileName);
                }
            }

            return BadRequest("No data provided for export.");
        }
    }
 }
