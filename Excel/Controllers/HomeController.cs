using Aspose.Cells;
using Excel.Extentions;
using Excel.Models;
using Excel.Models.Context;
using Excel.Models.Entities;
using Excel.Models.ViewModels;
using Excel.Services;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Excel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly ExcelContext _context;
        private readonly IWebHostEnvironment _hostEnvironment;

        public HomeController(ILogger<HomeController> logger, ExcelContext context, IWebHostEnvironment hostEnvironment)
        {
            _logger = logger;
            _context = context;
            _hostEnvironment = hostEnvironment;
        }


        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost, ValidateAntiForgeryToken]
        public async Task<IActionResult> Index(IFormFile file,string entityName)
        {
            if (!ModelState.IsValid)
                return View();

            if (file == null || string.IsNullOrWhiteSpace(entityName))
                return View();

            var validateFileExtention = file.FileName.IsValidFile();
            if (validateFileExtention.FileExtention == "" || validateFileExtention.IsValid == false)
                return View();

            var watchTime = new Stopwatch();
            watchTime.Start();

            //AsposeExcelOperation contains many methods and function about take and save data to database or convert csv file to xlsx 
            AsposeExcelOperation convertToExcel = new AsposeExcelOperation(_hostEnvironment, _context);

            string savefilefirstdirectoryPath =await convertToExcel.SaveExcelFileRoot(file);

            var book = new Workbook(savefilefirstdirectoryPath);

            var check5 = book.Worksheets[0].Cells.Rows[0];

            string saveFilePathWithName = convertToExcel.ConvertandSaveXlsxToCsv(savefilefirstdirectoryPath);

            var rowResult = convertToExcel.KeepInMemory(savefilefirstdirectoryPath);

            var convertResult = convertToExcel.ConvertCsvToExcleFile(saveFilePathWithName);
            if (!convertResult)
                return NotFound("somthing went wrong");

            if (rowResult)
            {
                try
                {
                    System.IO.File.Delete(saveFilePathWithName);
                    System.IO.File.Delete(savefilefirstdirectoryPath);
                }
                catch (Exception exp)
                {
                    return View(exp.Message);
                }
            }

            watchTime.Stop();
            string timeSapn = watchTime.ElapsedMilliseconds.ToString();

            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
