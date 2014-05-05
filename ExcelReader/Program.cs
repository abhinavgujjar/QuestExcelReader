using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using System.IO;
using ExcelReader.Models;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;


namespace ExcelReader
{
    class Program
    {

        static void Main(string[] args)
        {
            var options = new Options();

            var result = CommandLine.Parser.Default.ParseArguments(args, options);
            if (result)
            {
                // Values are available here
                if (options.Verbose) Console.WriteLine("Filename: {0}", options.InputFile);
            }
            else
            {
                Console.WriteLine("Error in arguments");
                return;
            }

            string docName = options.InputFile;

            FileInfo existingFile = LoadFile(docName);
            if (existingFile == null)
            {
                Console.WriteLine("File opening error");
                return;
            }

            var config = LoadConfiguraiton(options.UploadType);
            if (config == null)
            {
                Console.WriteLine("Configuration File Error");
                return;
            }


            using (var package = new ExcelPackage(existingFile))
            {
                var wb = package.Workbook;
                var worksheet = wb.Worksheets.First();

                IImporter importer = GetImporter(options.UploadType, worksheet, config);
                importer.Import();

                Console.WriteLine("Done Importing");
                
                package.Save();
            }

            Console.ReadLine();
        }

        private static IImporter GetImporter(string importType, ExcelWorksheet worksheet, JToken config)
        {
            IImporter importer;

            switch (importType)
            {
                case "StudentProfile":
                    importer = new Importer(worksheet, config);
                    break;
                case "Incremental":
                    importer = new ScoresImporter(worksheet, config);
                    break;
                default:
                    importer = new Importer(worksheet, config);
                    break;
            }

            return importer;
        }

        private static FileInfo LoadFile(string docName)
        {
            FileInfo existingFile;
            existingFile = new FileInfo(docName);
            if (!existingFile.Exists)
            {
                Console.WriteLine(String.Format("Could not find file {0}", existingFile.FullName));
            }

            return existingFile;
        }

        private static JToken LoadConfiguraiton(string uploadType)
        {
            JToken targetConfig = null;
            try
            {

                using (var reader = File.OpenText(@"import.config"))
                {
                    var config = (JObject)JToken.ReadFrom(new JsonTextReader(reader));


                    foreach (var entry in config["configurations"])
                    {
                        if ((string)entry["file"] == uploadType)
                        {
                            targetConfig = entry;
                        }
                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Error reading configuration file!");
                Console.WriteLine(e.Message);
            }

            return targetConfig;
        }

    }
}
