using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic; 
using OfficeOpenXml;
using DataFlow.Models;
using DataFlow.Interfaces;

namespace DataFlow.Converters
{
    public class LinuxConverter : IConverterService
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
            { "DestinationName", new string[5] },
            { "DestinationIP", new string[5] },
            { "Protocol", new string[5] },
            { "Port", new string[5] }
        };

        public LinuxConverter(string filepath, string delimiter, string folderpath)
        {
            //Remove extension from filename
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

            // Append first row to csv string builder
            StringBuilder csvbuilder = new StringBuilder();
            csvbuilder.AppendLine(string.Join(csv.Delimiter, WorkSheet["Header"][1], WorkSheet["Header"][2], WorkSheet["Header"][3], WorkSheet["Header"][5], WorkSheet["Header"][0], "Status"));
 
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

        public void ConvertToScript(ExcelWorksheet worksheet, int rowcount, int colcount)
        {
            Document bash = new Document($"./{folderpath}/{filepath}Script.sh");
            FileInfo bashfile = new FileInfo(bash.FilePath);

            // Create a new Powershell script if it does not exist
            bash.CreateNew(bashfile);

            // Iterate through worksheet first row and get all headers
            for (int row = 1; row < 2; row++)
            {
                for (int col = 1; col <= colcount; col++)
                {
                    WorkSheet["Header"].Add(worksheet.Cells[row, col].Value.ToString());                        
                }
            }

            // Append script headers
            StringBuilder scriptbuilder = new StringBuilder();
            scriptbuilder.Append("#! /usr/bin/env bash\n");
            scriptbuilder.Append("## Bash script: Test For Open Ports\n");
 
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

                // Append script headers
                scriptbuilder.Append($"## WifLine: {WorkSheet["Data"][0]}\n");

                // Append all lines to build the script
                for (int source = 0; source < numberofsourcename; source++)
                {
                    if (source < 1) 
                    {
                        scriptbuilder.Append($"touch Wifline{WorkSheet["Data"][0]}_Result.txt\n");
                    }

                    for (int destination = 0; destination < numberofdestinationip; destination++)
                    {
                        for (int port = 0; port < numberofport; port++)
                        {
                            scriptbuilder.Append($"echo \"Flow: {WorkSheet["Data"][0]} - Source: {Flow["SourceName"][source]} - Destination: {Flow["DestinationName"][destination]} ({Flow["DestinationIP"][destination]}) - Port: {Flow["Port"][port]} - Date: $(date)\" >> Wifline{WorkSheet["Data"][0]}_Result.txt\n");
                            scriptbuilder.Append($"./portwass check -p tcp {Flow["DestinationIP"][destination]} {Flow["Port"][port]} >> Wifline{WorkSheet["Data"][0]}_Result.txt\n");
                        }
                    }
                }
                WorkSheet["Header"].Clear();
                WorkSheet["Data"].Clear();
            }              

            // Write to the file from but first remove new line  
            File.WriteAllText(bash.FilePath, scriptbuilder.ToString());
        }

        public void ConvertToMultipleScript(ExcelWorksheet worksheet, int rowcount, int colcount)
        {
            List<KeyValuePair<string, string>> ListOfPortQry = new List<KeyValuePair<string, string>>();

            List<string> ListOfSourceNames = new List<string>();

            List<string> ListOfDistinctSourceNames = new List<string>();

            // Iterate through worksheet starting on second row and get all data
            for (int row = 2; row <= rowcount; row++)
            {
                for (int col = 1; col <= colcount; col++)
                {
                    WorkSheet["Data"].Add(worksheet.Cells[row, col].Value.ToString().Replace("\n", String.Empty));                        
                }

                string flow = WorkSheet["Data"][0];

                // Split flows into different categories
                SplitFlow(WorkSheet, ref Flow);

                // Combine different length and iterate through all possible flows
                int numberofsourcename = Flow["SourceName"].Length;
                int numberofdestinationip = Flow["DestinationIP"].Length;
                int numberofport = Flow["Port"].Length;

                // Append all lines to build the script
                for (int source = 0; source < numberofsourcename; source++)
                {
                    for (int destination = 0; destination < numberofdestinationip; destination++)
                    {
                        for (int port = 0; port < numberofport; port++)
                        {
                            // Get list of all source servers build list of port scanner tool commands 
                            ListOfSourceNames.Add(Flow["SourceName"][source]);
                            ListOfPortQry.Add(new KeyValuePair<string, string>(Flow["SourceName"][source], $"echo \"Flow: {flow} - Source: {Flow["SourceName"][source]} - Destination: {Flow["DestinationName"][destination]} ({Flow["DestinationIP"][destination]}) - Port: {Flow["Port"][port]} - Date: $(date)\" >> {Flow["SourceName"][source]}_Result.txt"));
                            ListOfPortQry.Add(new KeyValuePair<string, string>(Flow["SourceName"][source], $"./portwass check -p tcp {Flow["DestinationIP"][destination]} {Flow["Port"][port]} >> {Flow["SourceName"][source]}_Result.txt"));
                        }
                    }
                }
                WorkSheet["Header"].Clear();
                WorkSheet["Data"].Clear();
            }

            // Build a list of distinct server names and return the number of elements 
            ListOfDistinctSourceNames = ListOfSourceNames.Distinct().ToList();
            int numberofdistinctsourcenames = ListOfDistinctSourceNames.Count();

            // Use key/value arrays based on the number distinct servers
            StringBuilder[] scriptbuilder = new StringBuilder[numberofdistinctsourcenames];
            Document[] scriptdocument = new Document[numberofdistinctsourcenames];
            FileInfo[] scriptfile = new FileInfo[numberofdistinctsourcenames];

            // Append all lines to build multiple scripts
            for (int value = 0; value < numberofdistinctsourcenames; value++)
            {
                scriptbuilder[value] = new StringBuilder();
                scriptdocument[value] = new Document($"./{folderpath}/{ListOfDistinctSourceNames[value]}Script.sh");
                scriptfile[value] = new FileInfo(scriptdocument[value].FilePath);

                try 
                {
                    // Create a new Bash script if it does not exist
                    if (!scriptfile[value].Exists)
                    {
                        scriptfile[value].Delete();
                        scriptfile[value] = new FileInfo(scriptdocument[value].FilePath);
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

            for (int value = 0; value < numberofdistinctsourcenames; value++)
            {
                scriptbuilder[value].Append("#! /usr/bin/env bash\n");
                scriptbuilder[value].Append("## Bash script: Test For Open Ports\n");
                scriptbuilder[value].Append($"## Source: {ListOfDistinctSourceNames[value]}\n");
                scriptbuilder[value].Append($"touch {ListOfDistinctSourceNames[value]}_Result.txt\n");
                foreach (KeyValuePair<string, string> portqry in ListOfPortQry)
                {
                    if (ListOfDistinctSourceNames[value] == portqry.Key)
                    {
                        scriptbuilder[value].Append($"{portqry.Value}\n");
                    }
                }
            }

            for (int value = 0; value < numberofdistinctsourcenames; value++)
            {
                // Write to file 
                File.WriteAllText(scriptdocument[value].FilePath, scriptbuilder[value].ToString());
            }
        }

        public void ConvertToServerPool(ExcelWorksheet worksheet, int rowcount, int colcount)
        {
            List<string> ListOfSourceNames = new List<string>();

            List<string> ListOfDistinctSourceNames = new List<string>();

            // Iterate through worksheet starting on second row and get all data 
            for (int row = 2; row <= rowcount; row++)
            {
                for (int col = 2; col < 3; col++)
                {
                    WorkSheet["Data"].Add(worksheet.Cells[row, col].Value.ToString().Replace("\n", String.Empty));                        
                }

                // Remove the delimiter from the list and return the number of elements    
                Flow["SourceName"] = WorkSheet["Data"][0].Split(delimiter);
                int numberofsourcename = Flow["SourceName"].Length;

                for (int source = 0; source < numberofsourcename; source++)
                {
                    ListOfSourceNames.Add(Flow["SourceName"][source]);
                }
                WorkSheet["Data"].Clear();
            }

            ListOfDistinctSourceNames = ListOfSourceNames.Distinct().ToList();

            try 
            {
                Document txt = new Document($"./{folderpath}/Servers.txt");
                FileInfo txtfile = new FileInfo(txt.FilePath);

                // Create a new txt file if it does not exist
                txt.CreateNew(txtfile);

                // Create a file to write to
                using (StreamWriter sw = txtfile.CreateText()) 
                {
                    foreach (string source in ListOfDistinctSourceNames)
                    {
                        sw.WriteLine($"{source}");
                    }
                    sw.Close();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Creating server list failed: {0}", e.ToString());
            }
        }
        public void SplitFlow(Dictionary<string, List<string>> WorkSheet, ref Dictionary<string, String[]> Flow)
        {
            Flow["SourceName"] = WorkSheet["Data"][1].Split(delimiter);
            Flow["DestinationName"] = WorkSheet["Data"][2].Split(delimiter);
            Flow["DestinationIP"] = WorkSheet["Data"][3].Split(delimiter);
            Flow["Protocol"] = WorkSheet["Data"][4].Split(delimiter);
            Flow["Port"] = WorkSheet["Data"][5].Split(delimiter);
        }
    }
}