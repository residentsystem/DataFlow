using System;
using System.IO;
using OfficeOpenXml;
using DataFlow.Models;
using DataFlow.Services;
using DataFlow.Converters;

namespace DataFlow
{
    class Program
    {
        static int Main(string[] args)
        {
            // Parse and test command line arguments
            Parser parser = new Parser();

            if ((parser.ParsingArguments(args)) == 1) {
                Console.WriteLine("\nUsage: DataFlow.exe -linux | -windows excelfile.xlsx");
                return (int)Parsing.Error;
            }

            // Add Dependency Injection to this project
            DependencyInjectionService.AddServiceCollection();

            // Get excel file properties
            string filepath = args[1];
            Document excel = new Document(filepath);
            FileInfo excelfile = new FileInfo(excel.FilePath);

            // Exit program if file is not found
            excel.FileNotFound(excelfile);

            // Extract spreadsheet file name from path, without the file extension
            string excelfilename = Path.GetFileNameWithoutExtension(excel.FilePath);

            // Create a new folder and name it the same has the Excel spreadsheet   
            Folder directory = new Folder(excelfilename);
            directory.CreateNew();

            // Get the worksheet from the excel spreadsheet and validate its content
            ExcelWorksheet worksheet = excel.GetWorksheet(excelfile);

            // Access the dimension address of the worksheet and get number of rows and column
            int rowcount = excel.GetRowCount(worksheet);
            int colcount = excel.GetColCount(worksheet);

            // Apply converter service based on command line argument
            if (args[0] == "-windows") 
            {
                WindowsConverter converter = new WindowsConverter(excel.FilePath, excel.Delimiter, directory.FolderName);
                ConverterScriptService service = new ConverterScriptService(converter);
                service.CreateScript(worksheet, rowcount, colcount);
            }
            else if (args[0] == "-linux") 
            {
                LinuxConverter converter = new LinuxConverter(excel.FilePath, excel.Delimiter, directory.FolderName);
                ConverterScriptService service = new ConverterScriptService(converter);
                service.CreateScript(worksheet, rowcount, colcount);
            }
            else if (args[0] == "-csv") 
            {
                CsvConverter converter = new CsvConverter(excel.FilePath, excel.Delimiter, directory.FolderName);
                ConverterFileService service = new ConverterFileService(converter);
                service.CreateFile(worksheet, rowcount, colcount);
            }
            return (int)Parsing.Success;
        }
    }
}
