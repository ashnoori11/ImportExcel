using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Cells;
using Excel.Extentions;
using Excel.Models.Context;
using Excel.Models.ViewModels;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;

namespace Excel.Services
{
    public class AsposeExcelOperation
    {
        private readonly IWebHostEnvironment _hostEnviroment;
        private readonly ExcelContext _context;

        public AsposeExcelOperation(IWebHostEnvironment hostEnviroment, ExcelContext context)
        {
            _hostEnviroment = hostEnviroment;
            _context = context;
        }
        public bool ConvertCsvToExcleFile(string csvFilePath)
        {
            try
            {

                TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
                loadOptions.ConvertNumericData = false;

                Workbook workbook = new Workbook(csvFilePath, loadOptions);
                Worksheet sheet = workbook.Worksheets[0];

                object[,] dataArray = sheet.Cells.ExportArray(0, 0, sheet.Cells.MaxDataRow + 1, sheet.Cells.MaxDataColumn + 1);

                if (dataArray != null)
                {
                    string length = "Array Length " + dataArray.Length;
                }


                Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet2.Cells.ImportTwoDimensionArray(dataArray, 0, 0);
                string savePath = $"{_hostEnviroment.WebRootPath}/FirstSavefile/worked.xlsx";
                workbook.Save(savePath, SaveFormat.Xlsx);

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        public bool ImportToDatabase(string csvFilePath)
        {
            try
            {
                // For complete examples and data files, please go to https://github.com/aspose-cells/Aspose.Cells-for-.NET
                // The path to the documents directory.
                // string dataDir = RunExamples.GetDataDir(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

                string dataDir = csvFilePath;

                // Create directory if it is not already present.
                bool IsExists = System.IO.Directory.Exists(dataDir);
                if (!IsExists)
                    System.IO.Directory.CreateDirectory(dataDir);

                // Instantiating a Workbook object
                Workbook workbook = new Workbook();

                // Obtaining the reference of the worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Creating an array containing names as string values
                string[] names = new string[] { "laurence chen", "roman korchagin", "kyle huang" };

                // Importing the array of names to 1st row and first column vertically
                worksheet.Cells.ImportArray(names, 0, 0, true);

                // Saving the Excel file
                workbook.Save(dataDir + "DataImport.out.xls");
                return true;
            }
            catch
            {
                return false;
            }
        }
        public bool SaveCsvWithMultipleEncodingAsExcelFile(string csvFilePath)
        {
            try
            {
                string filePath = csvFilePath;

                // Set Multi Encoded Property to True
                TxtLoadOptions options = new TxtLoadOptions();
                options.IsMultiEncoded = true;

                // Load the CSV file into Workbook
                Workbook workbook = new Workbook(filePath, options);

                // Save it in XLSX format
                workbook.Save(filePath + ".out.xlsx", SaveFormat.Xlsx);

                return true;
            }
            catch
            {
                return false;
            }
        }

        // it works
        public bool KeepInMemory(string csvFilePath)
        {
            List<string> row1 = new List<string>();
            List<GeneralViewModel> ListViewModel = new List<GeneralViewModel>();
            EpplusImportExcel importExcel = new EpplusImportExcel(_hostEnviroment,_context);
            List<string> AllRowsValues = new List<string>();
            GeneralViewModel viewModel =new GeneralViewModel();
            bool returnResult = false;

            try
            {
                byte[] filebytes=File.ReadAllBytes(csvFilePath);
                using (var memory=new MemoryStream(filebytes))
                {
                    var book = new Workbook(memory);
                    var worksheet = book.Worksheets[0];
                    int allRows = worksheet.Cells.MaxDataRow;
                    int allColumns = worksheet.Cells.MaxDataColumn;
                    int[] columnsArray = new int[allColumns+1];

                    for (int i = 0; i <= allColumns ; i++)
                    {
                        columnsArray[i] = i;
                    }

                    for (int i = 1; i <= allRows; i++)
                    {
                        foreach (var item in columnsArray)
                        {
                            AllRowsValues.Add(worksheet.Cells[i, item].Value.ToString());
                        }

                    }

                    //int calculate = (AllRowsValues.Count() / allRows);
                    List<string> AddTest = new List<string>();

                    int endRecent = 0;
                    int end = allColumns;
                    int allrows = allRows;
                    int starttIndex = 0;

                    for(int i = endRecent; i <= end; i++)
                    {
                        AddTest = SeprateListByNumber(AllRowsValues,starttIndex,end,endRecent);
                        returnResult = InsertToSqlServer(AddTest).Result;
                        if (returnResult)
                        {
                            int y = 0;
                            //endRecent = end+1;
                            //int sum = (i + 1);
                            //end = end * sum;
                            end++;
                            endRecent = end;
                            starttIndex = end;
                            end = end + allColumns+y;
                            y++;
                        }
                    }

                    memory.Dispose();
                }
           
                return true;
            }
            catch(Exception exp)
            {
                string message = exp.Message;
                string trace = exp.StackTrace.ToString();
                return false;
            }
        }
        public List<string> SeprateListByNumber(List<string> list,int? startNum,int? endNum,int? lastSeprateNums)
        {
            List<string> resultList = new List<string>();

            int end = (int)endNum;
            if(lastSeprateNums > 0)
            {
                end = (int)endNum + (int)lastSeprateNums;
            }

            try
            {
                for (int i = (int)startNum; i <= end; i++)
                {
                    resultList.Add(list[i].ToString());
                }

                return resultList;
            }
            catch
            {
                return resultList;
            }
        }
        public async Task<bool> InsertToSqlServer(List<string> columnsValuesOfRow)
        {
            bool result = false;
            EpplusImportExcel dbConnection = new EpplusImportExcel(_hostEnviroment, _context);
            GeneralViewModel viewModel = new GeneralViewModel();
            Type propertiesCount = typeof(GeneralViewModel);

            int propCount = propertiesCount.GetProperties().Count();
            Dictionary<string,string> requiredFields = VerifyUploadExtention.GetDisplayNameList<GeneralViewModel>();


            try
            {
                viewModel = new GeneralViewModel
                {
                    FirstName = columnsValuesOfRow[1],
                    LastName = columnsValuesOfRow[2],
                    Email = columnsValuesOfRow[3],
                    PhoneNumber = columnsValuesOfRow[4],
                    Count = int.Parse(columnsValuesOfRow[5].ToString()),
                    IsIS = columnsValuesOfRow[6] == "1" ? true : false,
                    CarName = columnsValuesOfRow[7],
                    Color = columnsValuesOfRow[8],
                    Price = columnsValuesOfRow[9],
                    MadeOn = Convert.ToDateTime(columnsValuesOfRow[10]),
                    Address = columnsValuesOfRow[11],
                };

               result= await dbConnection.SaveToDatabaseAsync(viewModel);

                return result;
            }
            catch (Exception exp)
            {
                return result;
            }
        }

        //----------------------

        public async Task<string> SaveExcelFileRoot(IFormFile file)
        {
            try
            {
                string newFileName = Guid.NewGuid().ToString();
                string savefilefirstdirectoryPath = $"{_hostEnviroment.WebRootPath}/FirstSavefile/{newFileName}{Path.GetExtension(file.FileName)}";
                if (!Directory.Exists($"{_hostEnviroment.WebRootPath}/FirstSavefile"))
                {
                    Directory.CreateDirectory($"{_hostEnviroment.WebRootPath}/FirstSavefile");
                }

                using (var stream = new FileStream(savefilefirstdirectoryPath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                return savefilefirstdirectoryPath;
            }
            catch (Exception exp)
            {
                return "";
            }
        }
        public string ConvertandSaveXlsxToCsv(string path)
        {
            var book = new Workbook(path);

            string directoryPath = $"{_hostEnviroment.ContentRootPath}/ExcelAsposeFiles";
            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            string newCsvfileName = Guid.NewGuid().ToString();
            string saveFilePathWithName = $"{directoryPath}/{newCsvfileName}.csv";

            // save XLSX as CSV
            book.Save(saveFilePathWithName, SaveFormat.Auto);

            return saveFilePathWithName;
        }
    }
}
