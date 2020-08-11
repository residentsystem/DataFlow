using System;
using System.IO;
using System.Security;

namespace DataFlow.Models
{
    public class Folder
    {
        public string FolderName { get; set; }

        public Folder(string foldername = "Rules")
        {
            this.FolderName = foldername;
        }

        public void CreateNew()
        {
            // Get folder path where scripts will be located
            DirectoryInfo scriptfolder = new DirectoryInfo(FolderName);

            try
            {
                // Create folder if it does not already exist
                if (!scriptfolder.Exists)
                {
                    scriptfolder.Create();
                }
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine($"Argument Exception Error: '{ex}'");
            }
            catch (DirectoryNotFoundException ex)
            {
                Console.WriteLine($"Argument Exception Error: '{ex}'");
            }
            catch (IOException ex)
            {
                Console.WriteLine($"IO Exception Error: '{ex}'");
            }
            catch (SecurityException ex)
            {
                Console.WriteLine($"Security Exception Error: '{ex}'");
            }
            catch (NotSupportedException ex)
            {
                Console.WriteLine($"Security Exception Error: '{ex}'");
            }
        }
    }
}