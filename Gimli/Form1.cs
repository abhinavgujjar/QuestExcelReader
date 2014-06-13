using ExcelReader;
using ExcelReader.Models;
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
        IImporter importer;
        JToken _config;
        QSStagingDbContext db = new QSStagingDbContext();

        public Form1()
        {
            InitializeComponent();
            openFileDialog1 = new OpenFileDialog();
            _config = LoadConfiguraiton();
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
                Log(string.Format("Awesome sauce - You've selected {0} to be imported", openFileDialog1.FileName));

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var existingFile = openFileDialog1.OpenFile();
            using (var package = new ExcelPackage(existingFile))
            {
                var wb = package.Workbook;
                var worksheet = wb.Worksheets.First();

                try
                {
                    var result = importer.Import();


                    if (result.Failed)
                    {
                        Log("Ooops. Something went wrong during the import");
                        Log(result.Message);
                    }
                    else
                    {
                        Log(String.Format("Found {0} Students to import", result.ImportStudents.Count()));
                        Log(String.Format("Saving to database, please wait ... "));

                        //all well till now for the center
                        foreach (var student in result.ImportStudents)
                        {
                            db.StudentProfiles.Add(student.Profile);
                            db.SaveChanges();

                            student.Placement.StudentUid = student.Profile.Uid;
                            student.Placement.StudentId = student.Profile.Id;
                            db.Placements.Add(student.Placement);
                            db.SaveChanges();

                            foreach (var score in student.Scores)
                            {
                                score.StudentId = student.Profile.Id;

                                score.Total = getTotalForsubject(score.Subject);
                               
                            }
                            db.LegacySubjectScores.AddRange(student.Scores);
                            db.SaveChanges();
                        }


                        buttonImport.Enabled = false;
                        Log("Hurrah! Saved to Database. ");
                        Log("-------------------------------------------");
                    }

                    package.Save();
                }
                catch (Exception excp)
                {
                    Log("Oops! Something went wrong with the import. Contact Abhijeet Mehta");
                    Log(excp.Message);
                    Log(excp.StackTrace);
                }

            }
        }

        private decimal getTotalForsubject(string subject)
        {
            var start = subject.IndexOf('(');
            var end = subject.IndexOf(')');

            if (start > 0 && end > 0 && end > start)
            {
                var rawTotal = subject.Substring(start + 1, end - start - 1);

                decimal total;
                Decimal.TryParse(rawTotal, out total);

                return total;
            }

            return 0;
        }


        private JToken LoadConfiguraiton()
        {
            JToken targetConfig = null;
            try
            {
                using (var reader = File.OpenText(@"legacyimport.config"))
                {
                    targetConfig = (JObject)JToken.ReadFrom(new JsonTextReader(reader));
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
            textBoxConsole.AppendText(Environment.NewLine + "> " + message);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(openFileDialog1.FileName))
            {
                Log("Umm.. you want to try selecting file first?");
                return;
            }

            Stream existingFile = null;
            try
            {
                Log("Starting Validation...");
                //open the file and validate the entries
                existingFile = openFileDialog1.OpenFile();
            }
            catch (Exception exp)
            {
                Log("Could not open file");
                Log(exp.Message);
                Log(exp.StackTrace);
            }

            try
            {
                using (var package = new ExcelPackage(existingFile))
                {
                    var wb = package.Workbook;
                    var worksheet = wb.Worksheets.First();

                    importer = new Importer(worksheet, _config);
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
                Log(ex.Message);
                if (ex.InnerException != null)
                    Log(ex.InnerException.Message);
            }
        }
    }
}
