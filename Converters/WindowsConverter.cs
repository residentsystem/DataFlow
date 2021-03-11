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
    public class WindowsConverter : IConverterScriptService
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

        public WindowsConverter(string filepath, string delimiter, string folderpath)
        {
            // Remove extension from filename
            this.filepath = Path.GetFileNameWithoutExtension(filepath);
            this.delimiter = delimiter;
            this.folderpath = folderpath;
        }

        public void ConvertToScript(ExcelWorksheet worksheet, int rowcount, int colcount)
        {
            Document powershell = new Document($"./{folderpath}/{filepath}Script.ps1");
            FileInfo powershellfile = new FileInfo(powershell.FilePath);

            // Create a new Powershell script if it does not exist
            powershell.CreateNew(powershellfile);

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
            scriptbuilder.AppendLine("## Powershell script: Test For Open Ports");
 
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
                scriptbuilder.AppendLine($"## WifLine: {WorkSheet["Data"][0]}");

                // Append all lines to build the script
                for (int source = 0; source < numberofsourcename; source++)
                {
                    for (int destination = 0; destination < numberofdestinationip; destination++)
                    {
                        for (int port = 0; port < numberofport; port++)
                        {
                            scriptbuilder.AppendLine($"Add-Content Wifline{WorkSheet["Data"][0]}_Result.txt \"Flow: {WorkSheet["Data"][0]} - Source: {Flow["SourceName"][source]} - Destination: {Flow["DestinationName"][destination]} ({Flow["DestinationIP"][destination]}) - Port: {Flow["Port"][port]} - Date: $(Get-Date)\"");
                            scriptbuilder.AppendLine($".\\PortQry.exe -n {Flow["DestinationIP"][destination]} -p tcp -e {Flow["Port"][port]} | find \": FILTERED\" >> Wifline{WorkSheet["Data"][0]}_Result.txt");
                        }
                    }
                }
                WorkSheet["Header"].Clear();
                WorkSheet["Data"].Clear();
            }              

            // Write to the file from but first remove new line  
            File.WriteAllText(powershell.FilePath, scriptbuilder.ToString().Replace("\n", String.Empty));
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
                            ListOfPortQry.Add(new KeyValuePair<string, string>(Flow["SourceName"][source], $"Add-Content {Flow["SourceName"][source]}_Result.txt \"Flow: {WorkSheet["Data"][0]} - Source: {Flow["SourceName"][source]} ({Flow["SourceIP"][source]}) - Destination: {Flow["DestinationName"][destination]} ({Flow["DestinationIP"][destination]}) - Port: {Flow["Port"][port]} - Date: $(Get-Date)\""));
                            ListOfPortQry.Add(new KeyValuePair<string, string>(Flow["SourceName"][source], $".\\PortQry.exe -n {Flow["DestinationIP"][destination]} -p tcp -e {Flow["Port"][port]} | find \": FILTERED\" >> {Flow["SourceName"][source]}_Result.txt"));
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
                scriptdocument[value] = new Document($"./{folderpath}/{ListOfDistinctSourceNames[value]}Script.ps1");
                scriptfile[value] = new FileInfo(scriptdocument[value].FilePath);

                try
                {
                    // Create a new Powershell script if it does not exist
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

            for (int value = 0; value < numberofdistinctsourcenames; value++)
            {
                // Write to files but first remove new line 
                File.WriteAllText(scriptdocument[value].FilePath, scriptbuilder[value].ToString().Replace("\n", String.Empty));
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
            Flow["SourceIP"] = WorkSheet["Data"][2].Split(delimiter);
            Flow["DestinationName"] = WorkSheet["Data"][3].Split(delimiter);
            Flow["DestinationIP"] = WorkSheet["Data"][4].Split(delimiter);
            Flow["Protocol"] = WorkSheet["Data"][5].Split(delimiter);
            Flow["Port"] = WorkSheet["Data"][6].Split(delimiter);
        }
    }
}