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
        JToken _config;
        QSStagingDbContext db = new QSStagingDbContext();

        public Form1()
        {
            InitializeComponent();
            openFileDialog1 = new OpenFileDialog();
            _config = LoadConfiguraiton();

            if (_config == null)
            {
                buttonImportProfile.Enabled = false;
                buttonImportScores.Enabled = false;
            }
            else
            {
                Log("Select File to import...");
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
                    var importer = new Importer(worksheet, _config);
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


                        buttonImportProfile.Enabled = false;
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

                Log("Loaded configuration file successfully." );

            }
            catch (Exception e)
            {
                Log("Config file is a little messed up - " + e.Message);
                Log("FIX CONFIG FILE");
            }

            return targetConfig;
        }

        private void Log(string message)
        {
            textBoxConsole.AppendText(Environment.NewLine + "> " + message);
        }

          private void buttonImport_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(openFileDialog1.FileName))
            {
                Log("No file selected for import. Select file...");
                return;
            }

            using (var existingFile = openFileDialog1.OpenFile())
            {
                using (var package = new ExcelPackage(existingFile))
                {
                    var wb = package.Workbook;
                    var worksheet = wb.Worksheets.First();

                    try
                    {
                        var importer = new Importer(worksheet, _config);
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
                                var targetStudent = db.StudentProfiles.Where(p => p.Uid == student.Profile.Uid).SingleOrDefault();
                                if (targetStudent == null)
                                {
                                    db.StudentProfiles.Add(student.Profile);
                                }
                                else
                                {
                                    //this will set the changed values to the taget student record from the database.
                                    db.Entry(targetStudent).CurrentValues.SetValues(student);
                                }

                                db.SaveChanges();
                            }


                            buttonImportProfile.Enabled = false;
                            Log("Hurrah! Saved to Database. ");
                            Log("-------------------------------------------");
                        }

                    }
                    catch (Exception excp)
                    {
                        Log("Oops! Something went wrong with the import. Contact Abhijeet Mehta");
                        Log(excp.Message);
                        Log(excp.StackTrace);
                    }
                }
            }
        }

        private void buttonSelectFile_Click(object sender, EventArgs e)
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
    }
}
