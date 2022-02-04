using Excel.Models.Context;
using Excel.Models.Entities;
using Excel.Models.ViewModels;
using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Excel.Services
{
    public class EpplusImportExcel
    {
        private readonly IWebHostEnvironment _hostEnviroment;
        private readonly ExcelContext _context;
        public EpplusImportExcel(IWebHostEnvironment hostEnviroment, ExcelContext context)
        {
            _hostEnviroment = hostEnviroment;
            _context = context;
        }
        public bool ImportData(out int count, string fileExtention)
        {
            bool result = false;
            count = 0;
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                string path = $"{_hostEnviroment.ContentRootPath }\\ExcelTest\\Untitledspreadsheet{fileExtention}";
                var mainPackage = new ExcelPackage(new System.IO.FileInfo(path));

                // data in excel file starts from clomun 1 row 1
                int startColumn = 1;

                // data in excel file starts from row 1 clomun 1
                int startRow = 1;

                //read sheet 1
                ExcelWorksheet workSheet = mainPackage.Workbook.Worksheets[0];
                var lastRow = workSheet.Cells.Where(cell => !string.IsNullOrEmpty(cell.Value?.ToString() ?? string.Empty)).LastOrDefault().End.Row;
                var starts = workSheet.Dimension.Start;
                var end = workSheet.Cells.Where(a => !string.IsNullOrEmpty(a.Value?.ToString() ?? string.Empty)).LastOrDefault().End.Column;
                List<int> lengthList = new List<int>();
                for (int i = 0; i < end; i++)
                {
                    lengthList.Add(i);
                }

                int[] lengthArray = lengthList.ToArray();

                //object className = null;
                //for (int i = 1; i <= lastRow; i++)
                //{
                //    for (int d = starts.Column; d < end; d++)
                //    {
                //        className = workSheet.Cells[i, d].Value;

                //    }
                //}

                object className = null;
                for (int i = 1; i <= lastRow; i++)
                {
                    for (int d = starts.Column; d < end; d++)
                    {
                        // className = workSheet.Cells[i, lengthArray[d]].Value;

                        var shas = SaveToDatabaseAsync(new GeneralViewModel
                        {
                            FirstName = workSheet.Cells[i, lengthArray[d - 1]].Value.ToString(),
                            LastName = workSheet.Cells[i, lengthArray[d]].Value.ToString(),
                            Email = workSheet.Cells[i, lengthArray[d + 1]].Value.ToString(),
                            PhoneNumber = workSheet.Cells[i, lengthArray[d + 2]].Value.ToString(),
                            Count = int.Parse(workSheet.Cells[i, lengthArray[d + 3]].Value.ToString()),
                        });

                    }
                }
            }
            catch (Exception exp)
            {
                throw;
            }

            return result;
        }
        public async Task<bool> SaveToDatabaseAsync(GeneralViewModel viewModel)
        {
            var result = false;
            try
            {
                var sampleModel = new Sample
                {
                    Address = viewModel.Address,
                    Count = viewModel.Count,
                    Email = viewModel.Email,
                    FirstName = viewModel.FirstName,
                    IsIS = viewModel.IsIS,
                    LastName = viewModel.LastName,
                    PhoneNumber = viewModel.PhoneNumber,
                };

                var relationModel = new Relation
                {
                    CarName = viewModel.CarName,
                    Color = viewModel.Color,
                    MadeOn = viewModel.MadeOn,
                    Price = viewModel.Price,
                };

                await _context.Samples.AddAsync(sampleModel);
                await _context.Relations.AddAsync(relationModel);

                await _context.SaveChangesAsync();


                var relation = new Sample_Relation
                {
                    SampleId = sampleModel.Id,
                    RelationId = relationModel.Id
                };

                await _context.Sample_Relations.AddAsync(relation);

                await _context.SaveChangesAsync();
                result = true;
            }
            catch
            {
                result = false;
            }
            return result;
        }

    }
}
