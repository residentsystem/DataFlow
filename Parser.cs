using System;
using DataFlow.Models;

namespace DataFlow
{
    class Parser 
    {
        public int ParsingArguments(String[] args)
        {
            // The right amount of arguments have to be supplied.
            if (args.Length < 2) {
                Console.WriteLine("\nMissing argument: Please supply all arguments.");

                if (args.Length == 1)
                {
                    ParsingSingleString(args[0]);
                }

                return (int)Parsing.Error;
            }
            else if (args.Length > 2) {
                Console.WriteLine("\nToo many arguments: Please supply a maximum of 2 arguments.");
                ParsingMultipleString(args);

                return (int)Parsing.Error;
            }

            // Validate all values if the correct amount of arguments have been supplied.
            if (args.Length == 2) {

                if (!(args[0] == "-windows" || args[0] == "-linux")) {
                    Console.WriteLine("\nValue Error: Please specify a valid argument.");
                    ParsingMultipleString(args);

                    return (int)Parsing.Error;
                }
            }
            // Return if parsed correctly
            return (int)Parsing.Success;
        }

        public void ParsingMultipleString(String[] args)
        {
            // Test if multiple arguments were supplied as a string.
            foreach(string arg in args)
            {
                bool IsArgumentString = int.TryParse(arg, out int argument);

                if (IsArgumentString) {
                    Console.WriteLine("Type Error: All arguments must be of type string.");
                    break;
                }  
            }
        }

        public void ParsingSingleString(string arg)
        {
            // Test if a single argument was supplied as a string.
            bool IsArgumentString = int.TryParse(arg, out int argument);

            if (IsArgumentString){
                Console.WriteLine("Type Error: The argument must be a string.");
            } 
        }
    }
}