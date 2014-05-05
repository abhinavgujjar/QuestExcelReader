using ExcelReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Gimli
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            openFileDialog1 = new OpenFileDialog();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            

            openFileDialog1.Filter = "Excel files (*.xls, *.xlsx)|*.xlsx; *.xls|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxFile.Text = openFileDialog1.FileName;
                Log(string.Format("Awesome sauce - You've selected {0} to be imported",  openFileDialog1.FileName));
                
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var existingFile = openFileDialog1.OpenFile();
            using (var package = new ExcelPackage(existingFile))
            {
                var wb = package.Workbook;
                var worksheet = wb.Worksheets.First();

                var importOption = importOptions.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked);

                var config = LoadConfiguraiton(importOption.Text);
                if (config == null)
                {
                    Console.WriteLine("Configuration File Error");
                    return;
                }

                try
                {
                    IImporter importer = GetImporter(importOption.Text, worksheet, config);
                    var result = importer.Import();

                    if (result.Failed)
                    {
                        Log("Ooops. Something went wrong during the import");
                        Log(result.Message);
                    }
                    else
                    {
                        buttonImport.Enabled = false;
                        Log("Hurrah !!! All is well");
                        Log( string.Format("{0} records imported", result.NumberOfRecords));
                    }



                    package.Save();
                }
                catch (Exception excp)
                {
                    Log("Oops! Something went wrong with the import. Contact Abhijeet Mehta");
                    Log(excp.Message);
                }

            }
        }

        private static IImporter GetImporter(string importType, ExcelWorksheet worksheet, JToken config)
        {
            IImporter importer;

            switch (importType)
            {
                case "Profile":
                    importer = new Importer(worksheet, config);
                    break;
                case "Scores":
                    importer = new ScoresImporter(worksheet, config);
                    break;
                default:
                    importer = new Importer(worksheet, config);
                    break;
            }

            return importer;
        }

        private JToken LoadConfiguraiton(string uploadType)
        {
            JToken targetConfig = null;
            try
            {
                using (var reader = File.OpenText(@"legacyimport.config"))
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
                Log("Config file is a little messed up - " + e.Message);
                Log("Aborting! It's ok. It happens... ");
            }

            return targetConfig;
        }

        private void Log(string message)
        {
            var current = textBoxConsole.Text;

            textBoxConsole.Text = current + Environment.NewLine + "> " + message;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(openFileDialog1.FileName))
            {
                Log("Umm.. you want to try selecting file first?");
                return;
            }

            try
            {
                Log("Starting Validation...");
                //open the file and validate the entries
                var existingFile = openFileDialog1.OpenFile();

                using (var package = new ExcelPackage(existingFile))
                {
                    var wb = package.Workbook;
                    var worksheet = wb.Worksheets.First();

                    var importOption = importOptions.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked);

                    var config = LoadConfiguraiton(importOption.Text);
                    if (config == null)
                    {
                        Console.WriteLine("Configuration File Error");
                        return;
                    }

                    IImporter importer = GetImporter(importOption.Text, worksheet, config);
                    var result = importer.Validate();

                    Log(result.Message);

                    if (result.Valid)
                    {
                        Log("Press Import to continue... ");
                        buttonImport.Enabled = true;
                    }
                    else
                    {
                        Log("Oh no! We can't import this file just yet. If you know what you're doing, go ahead and correct the import file and try again. Else - contact Abhijeet Mehta");
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
            }
        }
    }
}
