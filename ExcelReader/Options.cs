using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    class Options
    {
        [Option('r', "read", Required = true,HelpText = "Input file to be processed.")]
        public string InputFile { get; set; }

        [Option('t', "type", Required = true, HelpText = "Type of upload")]
        public string UploadType { get; set; }

        [Option('v', "verbose", DefaultValue = true, HelpText = "Prints all messages to standard output.")]
        public bool Verbose { get; set; }
    }
}
