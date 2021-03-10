using System;
using System.IO;
using System.Collections.Generic; 
using OfficeOpenXml;

namespace DataFlow.Models
{
    public class Document
    {
        public string FilePath { get; set; }
        public string Delimiter { get; set; }

        List<string> ListOfCellValues = new List<string>();

        public Document(string filepath, string delimiter = ",")
        {
            this.FilePath = filepath;
            this.Delimiter = delimiter;
        }

        public void FileNotFound(FileInfo file)
        {
            if (!file.Exists)
            {
                Console.WriteLine($"\nFile {FilePath} was not found!");
                Console.WriteLine("Usage: AppFlowConvert.exe <-linux|-windows> <excelfile>");
                Environment.Exit(0);
            }
        }

        public void CreateNew(FileInfo file)
        {
            try
            {
                if (!file.Exists)
                {
                    file.Delete();
                    file = new FileInfo(FilePath);
                }
            }
            catch (DriveNotFoundException ex)
            {
                Console.WriteLine($"Argument Exception Error: '{ex}'");
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"Argument Exception Error: '{ex}'");
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Argument Exception Error: '{ex}'");
            }
            catch (NotSupportedException ex)
            {
                Console.WriteLine($"Argument Exception Error: '{ex}'");
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine($"Argument Exception Error: '{ex}'");
            }
        }

        public ExcelWorksheet GetWorksheet(FileInfo excelfile)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Represents a top-level object to access all parts of an Excel file 
            ExcelPackage package = new ExcelPackage(excelfile);

            // Represents an Excel worksheet and allow access to all its properties and methods 
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            // Access the dimension address of the worksheet and get number of rows and column
            int rowcount = worksheet.Dimension.Rows;
            int colcount = worksheet.Dimension.Columns;

            // Iterate through worksheet starting at second row 
            for (int row = 2; row <= rowcount; row++)
            {
                for (int col = 1; col <= colcount; col++)
                {
                    ListOfCellValues.Add(worksheet.Cells[row, col].Value.ToString().Replace(" ", String.Empty));                    
                }

                string flow = ListOfCellValues[0];

                // Last character of a cell string must not end with a delimiter character
                for (int col = 1; col < colcount; col++)
                {
                    if (ListOfCellValues[col].Substring(ListOfCellValues[col].Length-1) == Delimiter)
                    {
                        Console.WriteLine($"\nError in file: {excelfile.Name} WifLine: {flow} Column: {++col}: Items are missing.");
                        Console.WriteLine($"Please add missing value(s) or make sure the content does not end with a delimiter.");
                    }                  
                }

                // Split multiple values contain in each cells into separate arrays    
                String[] SourceNames = ListOfCellValues[1].Split(Delimiter);
                String[] SourceIPs = ListOfCellValues[2].Split(Delimiter);
                String[] DestinationNames = ListOfCellValues[3].Split(Delimiter);
                String[] DestinationIPs = ListOfCellValues[4].Split(Delimiter);
                String[] Protocols = ListOfCellValues[5].Split(Delimiter);
                String[] Ports = ListOfCellValues[6].Split(Delimiter);

                // Destination and ip cells must contain the same number of elements 
                if (SourceNames.Length != SourceIPs.Length || DestinationNames.Length != DestinationIPs.Length)
                {
                    Console.WriteLine($"\nError in file: {excelfile.Name} WifLine: {flow}: Destination column and Ip column are of different length.");
                    Console.WriteLine($"Please make sure that both columns contain the same amount of items.");
                    Environment.Exit(0);
                }     
                ListOfCellValues.Clear();
            }
            return worksheet;
        }

        public int GetRowCount(ExcelWorksheet worksheet)
        {
            return worksheet.Dimension.Rows;
        } 

        public int GetColCount(ExcelWorksheet worksheet)
        {
            return worksheet.Dimension.Columns;
        } 
    }
}