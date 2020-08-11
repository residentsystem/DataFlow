﻿using System;
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
            Parser argument = new Parser();

            if ((argument.ParsingArguments(args)) == 1) {
                Console.WriteLine("\nUsage: DataFlow.exe <-linux|-windows> <excelfile>");
                return 1;
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
                ConverterService flow = new ConverterService(converter);
                flow.CreateScript(worksheet, rowcount, colcount);
            }
            else if (args[0] == "-linux") 
            {
                LinuxConverter converter = new LinuxConverter(excel.FilePath, excel.Delimiter, directory.FolderName);
                ConverterService flow = new ConverterService(converter);
                flow.CreateScript(worksheet, rowcount, colcount);
            }
            return 0;
        }
    }
}