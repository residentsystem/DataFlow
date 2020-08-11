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
        
        List<string> ListOfCellValues = new List<string>();

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

            // Ensures we create a new csv file if it does not exist
            csv.CreateNew(csvfile);

            // Iterate through worksheet first row and get all values from individual cells   
            for (int row = 1; row < 2; row++)
            {
                for (int col = 1; col <= colcount; col++)
                {
                    ListOfCellValues.Add(worksheet.Cells[row, col].Value.ToString().Replace(" ", String.Empty));                       
                }
            }

            // Append worksheet first row to csv string builder
            StringBuilder csvbuilder = new StringBuilder();
            csvbuilder.AppendLine(string.Join(csv.Delimiter, ListOfCellValues[1], ListOfCellValues[2], ListOfCellValues[3], ListOfCellValues[5], ListOfCellValues[0], "Status"));
            ListOfCellValues.Clear();
 
            // Iterate through worksheet from second row and get values of individual cells
            for (int row = 2; row <= rowcount; row++)
            {
                for (int col = 1; col <= colcount; col++)
                {
                    ListOfCellValues.Add(worksheet.Cells[row, col].Value.ToString().Replace("\n", String.Empty));                    
                }

                string flow = ListOfCellValues[0];

                // Split values contained in each cells into separate arrays    
                String[] SourceNames = ListOfCellValues[1].Split(delimiter);
                String[] DestinationNames = ListOfCellValues[2].Split(delimiter);
                String[] DestinationIPs = ListOfCellValues[3].Split(delimiter);
                String[] Protocols = ListOfCellValues[4].Split(delimiter);
                String[] Ports = ListOfCellValues[5].Split(delimiter);

                // Combine different length and iterate through all possibilities  
                int sourcecount = SourceNames.Length;
                int destinationcount = DestinationIPs.Length;
                int portcount = Ports.Length;

                // Append all lines to build the csv file
                for (int source = 0; source < sourcecount; source++)
                {
                    for (int destination = 0; destination < destinationcount; destination++)
                    {
                        for (int port = 0; port < portcount; port++)
                        {
                            csvbuilder.AppendLine(string.Join(csv.Delimiter, SourceNames[source], DestinationNames[destination], DestinationIPs[destination], Ports[port], flow) + $"{csv.Delimiter}");
                        }
                    }
                }
                ListOfCellValues.Clear();
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

            // Iterate through worksheet from first row and get values of individual cells 
            for (int row = 1; row < 2; row++)
            {
                for (int col = 1; col <= colcount; col++)
                {
                    ListOfCellValues.Add(worksheet.Cells[row, col].Value.ToString());                        
                }
            }
            ListOfCellValues.Clear();

            // Append script headers
            StringBuilder scriptbuilder = new StringBuilder();
            scriptbuilder.Append("#! /usr/bin/env bash\n");
            scriptbuilder.Append("## Bash script: Test For Open Ports\n");
 
            // Iterate through worksheet from second row and get values of individual cells 
            for (int row = 2; row <= rowcount; row++)
            {
                for (int col = 1; col <= colcount; col++)
                {
                    ListOfCellValues.Add(worksheet.Cells[row, col].Value.ToString().Replace("\n", String.Empty));                       
                }

                string flow = ListOfCellValues[0];

                // Split values contained in each cells into separate arrays    
                String[] SourceNames = ListOfCellValues[1].Split(delimiter);
                String[] DestinationNames = ListOfCellValues[2].Split(delimiter);
                String[] DestinationIPs = ListOfCellValues[3].Split(delimiter);
                String[] Protocols = ListOfCellValues[4].Split(delimiter);
                String[] Ports = ListOfCellValues[5].Split(delimiter);

                // Combine different length and iterate through all possibilities
                int sourcecount = SourceNames.Length;
                int destinationcount = DestinationIPs.Length;
                int portcount = Ports.Length;

                // Append script headers
                scriptbuilder.Append($"## WifLine: {flow}\n");

                // Append all lines to build the script
                for (int source = 0; source < sourcecount; source++)
                {
                    if (source < 1) 
                    {
                        scriptbuilder.Append($"touch Wifline{flow}_Result.txt\n");
                    }

                    for (int destination = 0; destination < destinationcount; destination++)
                    {
                        for (int port = 0; port < portcount; port++)
                        {
                            scriptbuilder.Append($"echo \"Flow: {flow} - Source: {SourceNames[source]} - Destination: {DestinationNames[destination]} ({DestinationIPs[destination]}) - Port: {Ports[port]} - Date: $(date)\" >> Wifline{flow}_Result.txt\n");
                            scriptbuilder.Append($"./portwass check -p tcp {DestinationIPs[destination]} {Ports[port]} >> Wifline{flow}_Result.txt\n");
                        }
                    }
                }
                ListOfCellValues.Clear();
            }              

            // Write to the file from but first remove new line  
            File.WriteAllText(bash.FilePath, scriptbuilder.ToString());
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
                    ListOfCellValues.Add(worksheet.Cells[row, col].Value.ToString().Replace("\n", String.Empty));                        
                }

                string flow = ListOfCellValues[0];

                // Split values contained in each cells into separate arrays    
                String[] SourceNames = ListOfCellValues[1].Split(delimiter);
                String[] DestinationNames = ListOfCellValues[2].Split(delimiter);
                String[] DestinationIPs = ListOfCellValues[3].Split(delimiter);
                String[] Protocols = ListOfCellValues[4].Split(delimiter);
                String[] Ports = ListOfCellValues[5].Split(delimiter);

                // Combine different length and iterate through all possibilities
                int sourcecount = SourceNames.Length;
                int destinationcount = DestinationIPs.Length;
                int portcount = Ports.Length;

                // Append all lines to build the script
                for (int source = 0; source < sourcecount; source++)
                {
                    for (int destination = 0; destination < destinationcount; destination++)
                    {
                        for (int port = 0; port < portcount; port++)
                        {
                            // Get list of all source servers 
                            // Build list of port scanner tool commands 
                            ListOfSourceNames.Add(SourceNames[source]);
                            ListOfPortQry.Add(new KeyValuePair<string, string>(SourceNames[source], $"echo \"Flow: {flow} - Source: {SourceNames[source]} - Destination: {DestinationNames[destination]} ({DestinationIPs[destination]}) - Port: {Ports[port]} - Date: $(date)\" >> {SourceNames[source]}_Result.txt"));
                            ListOfPortQry.Add(new KeyValuePair<string, string>(SourceNames[source], $"./portwass check -p tcp {DestinationIPs[destination]} {Ports[port]} >> {SourceNames[source]}_Result.txt"));
                        }
                    }
                }
                ListOfCellValues.Clear();
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
                scriptdocument[value] = new Document($"./{folderpath}/{ListOfDistinctSourceNames[value]}Script.sh");
                scriptfile[value] = new FileInfo(scriptdocument[value].FilePath);

                try 
                {
                    // Ensures we create a new Bash script if it does not exist
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

            for (int value = 0; value < sourcenamecount; value++)
            {
                // Write to file 
                File.WriteAllText(scriptdocument[value].FilePath, scriptbuilder[value].ToString());
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
                    ListOfCellValues.Add(worksheet.Cells[row, col].Value.ToString().Replace("\n", String.Empty));                        
                }

                // Remove the delimiter from the list and return the number of elements    
                String[] SourceNames = ListOfCellValues[0].Split(delimiter);
                int sourcecount = SourceNames.Length;

                for (int source = 0; source < sourcecount; source++)
                {
                    ListOfSourceNames.Add(SourceNames[source]);
                }
                ListOfCellValues.Clear();
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