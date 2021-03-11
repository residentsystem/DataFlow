using System;
using System.IO;
using System.Text;
using System.Collections.Generic; 
using OfficeOpenXml;
using DataFlow.Models;
using DataFlow.Interfaces;

namespace DataFlow.Converters
{
    public class CsvConverter : IConverterFileService
    {
        private string filepath;

        private string delimiter;

        private string folderpath;

        Dictionary<string, List<string>> WorkSheet = new Dictionary<string, List<string>>()
        {
            { "Header", new List<string>() },
            { "Data", new List<string>() }
        };

        Dictionary<string, String[]> Flow = new Dictionary<string, String[]>()
        {
            { "SourceName", new string[5] },
            { "SourceIP", new string[5] },
            { "DestinationName", new string[5] },
            { "DestinationIP", new string[5] },
            { "Protocol", new string[5] },
            { "Port", new string[5] }
        };

        public CsvConverter(string filepath, string delimiter, string folderpath)
        {
            // Remove extension from filename
            this.filepath = Path.GetFileNameWithoutExtension(filepath);
            this.delimiter = delimiter;
            this.folderpath = folderpath;
        }

        public void ConvertToCsv(ExcelWorksheet worksheet, int rowcount, int colcount)
        {
            Document csv = new Document($"./{folderpath}/{filepath}Rules.csv", "|");
            FileInfo csvfile = new FileInfo(csv.FilePath);

            // Create a new csv file if it does not exist
            csv.CreateNew(csvfile);

            // Iterate through worksheet first row and get all headers
            for (int row = 1; row < 2; row++)
            {
                for (int col = 1; col <= colcount; col++)
                {
                    WorkSheet["Header"].Add(worksheet.Cells[row, col].Value.ToString().Replace(" ", String.Empty));                       
                }
            }

            // Append headers to csv string builder
            StringBuilder csvbuilder = new StringBuilder();
            csvbuilder.AppendLine(string.Join(csv.Delimiter, WorkSheet["Header"][1], WorkSheet["Header"][3], WorkSheet["Header"][4], WorkSheet["Header"][6], WorkSheet["Header"][0]));
 
            // Iterate through worksheet starting on second row and get all data
            for (int row = 2; row <= rowcount; row++)
            {
                for (int col = 1; col <= colcount; col++)
                {
                    WorkSheet["Data"].Add(worksheet.Cells[row, col].Value.ToString().Replace("\n", String.Empty));                    
                }

                // Split flows into different categories
                SplitFlow(WorkSheet, ref Flow);

                // Combine different length and iterate through all possible flows
                int numberofsourcename = Flow["SourceName"].Length;
                int numberofdestinationip = Flow["DestinationIP"].Length;
                int numberofport = Flow["Port"].Length;

                // Append all lines to build the csv file
                for (int source = 0; source < numberofsourcename; source++)
                {
                    for (int destination = 0; destination < numberofdestinationip; destination++)
                    {
                        for (int port = 0; port < numberofport; port++)
                        {
                            csvbuilder.AppendLine(string.Join(csv.Delimiter, Flow["SourceName"][source], Flow["DestinationName"][destination], Flow["DestinationIP"][destination], Flow["Port"][port], WorkSheet["Data"][0]) + $"{csv.Delimiter}");
                        }
                    }
                }
                WorkSheet["Header"].Clear();
                WorkSheet["Data"].Clear();
            }

            // Write to files from string builder  
            File.WriteAllText(csv.FilePath, csvbuilder.ToString());
        }
        public void SplitFlow(Dictionary<string, List<string>> WorkSheet, ref Dictionary<string, String[]> Flow)
        {
            Flow["SourceName"] = WorkSheet["Data"][1].Split(delimiter);
            Flow["SourceIP"] = WorkSheet["Data"][2].Split(delimiter);
            Flow["DestinationName"] = WorkSheet["Data"][3].Split(delimiter);
            Flow["DestinationIP"] = WorkSheet["Data"][4].Split(delimiter);
            Flow["Protocol"] = WorkSheet["Data"][5].Split(delimiter);
            Flow["Port"] = WorkSheet["Data"][6].Split(delimiter);
        }
    }
}