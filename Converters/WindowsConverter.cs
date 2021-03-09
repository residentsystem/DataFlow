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
    public class WindowsConverter : IConverterService
    {
        private string filepath;

        private string delimiter;

        private string folderpath;

        Dictionary<string, List<string>> WorkSheet = new Dictionary<string, List<string>>()
        {
            { "Header", new List<string>() },
            { "Flow", new List<string>() }
        };

        Dictionary<string, String[]> Cell = new Dictionary<string, String[]>()
        {
            { "SourceName", new string[5] },
            { "DestinationName", new string[5] },
            { "DestinationIP", new string[5] },
            { "Protocol", new string[5] },
            { "Port", new string[5] }
        };

        public WindowsConverter(string filepath, string delimiter, string folderpath)
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

            // Ensures we create a new csv file if it does not exist
            csv.CreateNew(csvfile);

            // Iterate through worksheet first row and get all values from individual cells   
            for (int row = 1; row < 2; row++)
            {
                for (int col = 1; col <= colcount; col++)
                {
                    WorkSheet["Header"].Add(worksheet.Cells[row, col].Value.ToString().Replace(" ", String.Empty));                       
                }
            }

            // Append worksheet first row to csv string builder
            StringBuilder csvbuilder = new StringBuilder();
            csvbuilder.AppendLine(string.Join(csv.Delimiter, WorkSheet["Header"][1], WorkSheet["Header"][2], WorkSheet["Header"][3], WorkSheet["Header"][5], WorkSheet["Header"][0], "Status"));
 
            // Iterate through worksheet from second row and get values of individual cells
            for (int row = 2; row <= rowcount; row++)
            {
                for (int col = 1; col <= colcount; col++)
                {
                    WorkSheet["Flow"].Add(worksheet.Cells[row, col].Value.ToString().Replace("\n", String.Empty));                    
                }

                string flow = WorkSheet["Flow"][0];

                // Split values contained in each cells into separate arrays    
                Cell["SourceName"] = WorkSheet["Flow"][1].Split(delimiter);
                Cell["DestinationName"] = WorkSheet["Flow"][2].Split(delimiter);
                Cell["DestinationIP"] = WorkSheet["Flow"][3].Split(delimiter);
                Cell["Protocol"] = WorkSheet["Flow"][4].Split(delimiter);
                Cell["Port"] = WorkSheet["Flow"][5].Split(delimiter);

                // Combine different length and iterate through all possibilities
                int sourcecount = Cell["SourceName"].Length;
                int destinationcount = Cell["DestinationIP"].Length;
                int portcount = Cell["Port"].Length;

                // Append all lines to build the csv file
                for (int source = 0; source < sourcecount; source++)
                {
                    for (int destination = 0; destination < destinationcount; destination++)
                    {
                        for (int port = 0; port < portcount; port++)
                        {
                            csvbuilder.AppendLine(string.Join(csv.Delimiter, Cell["SourceName"][source], Cell["DestinationName"][destination], Cell["DestinationIP"][destination], Cell["Port"][port], flow) + $"{csv.Delimiter}");
                        }
                    }
                }
                WorkSheet["Header"].Clear();
                WorkSheet["Flow"].Clear();
            }

            // Write to files from string builder  
            File.WriteAllText(csv.FilePath, csvbuilder.ToString());
        }

        public void ConvertToScript(ExcelWorksheet worksheet, int rowcount, int colcount)
        {
            Document powershell = new Document($"./{folderpath}/{filepath}Script.ps1");
            FileInfo powershellfile = new FileInfo(powershell.FilePath);

            // Create a new Powershell script if it does not exist
            powershell.CreateNew(powershellfile);

            // Iterate through worksheet from first row and get values of individual cells 
            for (int row = 1; row < 2; row++)
            {
                for (int col = 1; col <= colcount; col++)
                {
                    WorkSheet["Header"].Add(worksheet.Cells[row, col].Value.ToString());                        
                }
            }

            // Append script headers
            StringBuilder scriptbuilder = new StringBuilder();
            scriptbuilder.AppendLine("## Powershell script: Test For Open Ports");
 
            // Iterate through worksheet from second row and get values of individual cells 
            for (int row = 2; row <= rowcount; row++)
            {
                for (int col = 1; col <= colcount; col++)
                {
                    WorkSheet["Flow"].Add(worksheet.Cells[row, col].Value.ToString().Replace("\n", String.Empty));                       
                }

                string flow = WorkSheet["Flow"][0];

                // Split values contained in each cells into separate arrays    
                Cell["SourceName"] = WorkSheet["Flow"][1].Split(delimiter);
                Cell["DestinationName"] = WorkSheet["Flow"][2].Split(delimiter);
                Cell["DestinationIP"] = WorkSheet["Flow"][3].Split(delimiter);
                Cell["Protocol"] = WorkSheet["Flow"][4].Split(delimiter);
                Cell["Port"] = WorkSheet["Flow"][5].Split(delimiter);

                // Combine different length and iterate through all possibilities
                int sourcecount = Cell["SourceName"].Length;
                int destinationcount = Cell["DestinationIP"].Length;
                int portcount = Cell["Port"].Length;

                // Append script headers
                scriptbuilder.AppendLine($"## WifLine: {flow}");

                // Append all lines to build the script
                for (int source = 0; source < sourcecount; source++)
                {
                    for (int destination = 0; destination < destinationcount; destination++)
                    {
                        for (int port = 0; port < portcount; port++)
                        {
                            scriptbuilder.AppendLine($"Add-Content Wifline{flow}_Result.txt \"Flow: {flow} - Source: {Cell["SourceName"][source]} - Destination: {Cell["DestinationName"][destination]} ({Cell["DestinationIP"][destination]}) - Port: {Cell["Port"][port]} - Date: $(Get-Date)\"");
                            scriptbuilder.AppendLine($".\\PortQry.exe -n {Cell["DestinationIP"][destination]} -p tcp -e {Cell["Port"][port]} | find \": FILTERED\" >> Wifline{flow}_Result.txt");
                        }
                    }
                }
                WorkSheet["Header"].Clear();
                WorkSheet["Flow"].Clear();
            }              

            // Write to the file from but first remove new line  
            File.WriteAllText(powershell.FilePath, scriptbuilder.ToString().Replace("\n", String.Empty));
        }

        public void ConvertToMultipleScript(ExcelWorksheet worksheet, int rowcount, int colcount)
        {
            List<string> ListOfSourceNames = new List<string>();

            List<KeyValuePair<string, string>> ListOfPortQry = new List<KeyValuePair<string, string>>();

            List<string> ListOfDistinctSourceNames = new List<string>();

            // Iterate through worksheet from second row and get values of individual cells
            for (int row = 2; row <= rowcount; row++)
            {
                for (int col = 1; col <= colcount; col++)
                {
                    WorkSheet["Flow"].Add(worksheet.Cells[row, col].Value.ToString().Replace("\n", String.Empty));                        
                }

                string flow = WorkSheet["Flow"][0];

                // Split values contained in each cells into separate arrays    
                Cell["SourceName"] = WorkSheet["Flow"][1].Split(delimiter);
                Cell["DestinationName"] = WorkSheet["Flow"][2].Split(delimiter);
                Cell["DestinationIP"] = WorkSheet["Flow"][3].Split(delimiter);
                Cell["Protocol"] = WorkSheet["Flow"][4].Split(delimiter);
                Cell["Port"] = WorkSheet["Flow"][5].Split(delimiter);

                // Combine different length and iterate through all possibilities
                int sourcecount = Cell["SourceName"].Length;
                int destinationcount = Cell["DestinationIP"].Length;
                int portcount = Cell["Port"].Length;

                // Append all lines to build the script
                for (int source = 0; source < sourcecount; source++)
                {
                    for (int destination = 0; destination < destinationcount; destination++)
                    {
                        for (int port = 0; port < portcount; port++)
                        {
                            // Get list of all source servers 
                            // Build list of port scanner tool commands 
                            ListOfSourceNames.Add(Cell["SourceName"][source]);
                            ListOfPortQry.Add(new KeyValuePair<string, string>(Cell["SourceName"][source], $"Add-Content {Cell["SourceName"][source]}_Result.txt \"Flow: {flow} - Source: {Cell["SourceName"][source]} - Destination: {Cell["DestinationName"][destination]} ({Cell["DestinationIP"][destination]}) - Port: {Cell["Port"][port]} - Date: $(Get-Date)\""));
                            ListOfPortQry.Add(new KeyValuePair<string, string>(Cell["SourceName"][source], $".\\PortQry.exe -n {Cell["DestinationIP"][destination]} -p tcp -e {Cell["Port"][port]} | find \": FILTERED\" >> {Cell["SourceName"][source]}_Result.txt"));
                        }
                    }
                }
                WorkSheet["Header"].Clear();
                WorkSheet["Flow"].Clear();
            }

            // Build a list of distinct server names and return the number of elements
            ListOfDistinctSourceNames = ListOfSourceNames.Distinct().ToList();
            int sourcenamecount = ListOfDistinctSourceNames.Count();

            // Use key/value arrays based on the number distinct servers
            StringBuilder[] scriptbuilder = new StringBuilder[sourcenamecount];
            Document[] scriptdocument = new Document[sourcenamecount];
            FileInfo[] scriptfile = new FileInfo[sourcenamecount];

            // Append all lines to build multiple scripts
            for (int value = 0; value < sourcenamecount; value++)
            {
                scriptbuilder[value] = new StringBuilder();
                scriptdocument[value] = new Document($"./{folderpath}/{ListOfDistinctSourceNames[value]}Script.ps1");
                scriptfile[value] = new FileInfo(scriptdocument[value].FilePath);

                try
                {
                    // Ensures we create a new Powershell script if it does not exist
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

            for (int value = 0; value < sourcenamecount; value++)
            {
                scriptbuilder[value].AppendLine("## Powershell script: Test For Open Ports");
                scriptbuilder[value].AppendLine($"## Source: {ListOfDistinctSourceNames[value]}");
                foreach (KeyValuePair<string, string> portqry in ListOfPortQry)
                {
                    if (ListOfDistinctSourceNames[value] == portqry.Key)
                    {
                        scriptbuilder[value].AppendLine($"{portqry.Value}");
                    }
                }
            }

            for (int value = 0; value < sourcenamecount; value++)
            {
                // Write to files but first remove new line 
                File.WriteAllText(scriptdocument[value].FilePath, scriptbuilder[value].ToString().Replace("\n", String.Empty));
            }
        }

        public void ConvertToServerPool(ExcelWorksheet worksheet, int rowcount, int colcount)
        {
            List<string> ListOfSourceNames = new List<string>();

            List<string> ListOfDistinctSourceNames = new List<string>();

            // Iterate through worksheet from second row and get values from second column of every row 
            for (int row = 2; row <= rowcount; row++)
            {
                for (int col = 2; col < 3; col++)
                {
                    WorkSheet["Flow"].Add(worksheet.Cells[row, col].Value.ToString().Replace("\n", String.Empty));                        
                }

                // Remove the delimiter from the list and return the number of elements     
                Cell["SourceName"] = WorkSheet["Flow"][0].Split(delimiter);
                int sourcecount = Cell["SourceName"].Length;

                for (int source = 0; source < sourcecount; source++)
                {
                    ListOfSourceNames.Add(Cell["SourceName"][source]);
                }
                WorkSheet["Flow"].Clear();
            }

            ListOfDistinctSourceNames = ListOfSourceNames.Distinct().ToList();

            try 
            {
                Document txt = new Document($"./{folderpath}/Servers.txt");
                FileInfo txtfile = new FileInfo(txt.FilePath);

                // Ensures we create a new txt file if it does not exist
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
    }
}