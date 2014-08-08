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
using System.Text.RegularExpressions;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

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
                        

                       

                        //all well till now for the center
                        foreach (var student in result.ImportStudents)
                        {
                            db.StudentProfiles.Add(student.Profile);
                           
                            db.SaveChanges();
                            Log("Hurrah! Saved Student Profile to Database. ");

                            /*student.Placement.StudentUid = student.Profile.Uid;
                            student.Placement.StudentId = student.Profile.Id;
                            db.Placements.Add(student.Placement);
                            db.SaveChanges();*/
                            

                            foreach (var score in student.Scores)
                            {
                                score.StudentId = student.Profile.Id;

                                score.Total = getTotalForsubject(score.Subject);
                               
                            }
                            db.LegacySubjectScores.AddRange(student.Scores);
                            db.SaveChanges();



                            foreach (var placement in result.ImportStudents)
                            {
                                
                                db.Placements.Add(placement.Placement);
                                db.SaveChanges();

                                Log("Hurrah! Saved Placement to Database. ");
                            }

                            
                        }


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

          private void buttonImportProfile_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(openFileDialog1.FileName))
            {
                Log("No file selected for import. Select file...");
                return;
            }
       try
          {           
            using ( var existingFile = openFileDialog1.OpenFile())
            {
                using (var package = new ExcelPackage(existingFile))
                {
                    var wb = package.Workbook;
                    var worksheet = wb.Worksheets.First();

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
                            int validateflag = 0;
                            //all well till now for the center
                            foreach (var student in result.ImportStudents)
                            {
                               
                                //Validation Checking Started for Student Profile

                                //Fistname Validation
                                if (Regex.IsMatch(student.Profile.FirstName.Trim(), @"^[a-zA-Z .]*$") == false)
                                {
                                    Log(String.Format("First name format incorrect for Student Id:- {0} ", student.Profile.Uid));
                                    validateflag = 1;
                                }

                                //Lastname Validation
                                if (Regex.IsMatch(student.Profile.LastName.Trim(), @"^[a-zA-Z .]*$") == false)
                                {
                                    Log(String.Format("Last name format incorrect for Student Id:- {0} ", student.Profile.Uid));
                                    validateflag = 1;
                                }               

                                //Age Validation
                                if (Regex.IsMatch(student.Profile.Age.ToString(), @"^[0-9]+$")==false && student.Profile.Age<=100 )                               
                                {
                                    Log(String.Format("Age value or format incorrect for Student Id:- {0} ", student.Profile.Uid));
                                    validateflag = 1;
                                }   

                                //Email Validation
                                if (Regex.IsMatch(student.Profile.Email, @"^(?!\.)(""([^""\r\\]|\\[""\r\\])*""|"+ @"([-a-z0-9!#$%&'*+/=?^_`{|}~]|(?<!\.)\.)*)(?<!\.)" + @"@[a-z0-9][\w\.-]*[a-z0-9]\.[a-z][a-z\.]*[a-z]$") == false)
                                {
                                    if (String.IsNullOrEmpty(student.Profile.Email) == false)
                                    {
                                        Log(String.Format("Email format incorrect for Student Id:- {0} ", student.Profile.Uid));
                                        validateflag = 1;
                                    }
                                }

                                //Mobile Number Validation
                                if (Regex.IsMatch(student.Profile.Mobile, @"^[0-9]+$") == false || student.Profile.Mobile.Length != 10)
                                {
                                    if (String.IsNullOrEmpty(student.Profile.Mobile)==false)
                                    {
                                        Log(String.Format("Mobile Number format incorrect for Student Id:- {0} ", student.Profile.Uid));
                                        validateflag = 1;
                                    }
                                    
                                }

                                //Training Centre VAlidation
                                if (Regex.IsMatch(student.Profile.TrainingCenter.Trim(), @"^[a-zA-Z .]*$") == false)
                                {                    
                                        Log(String.Format("Training Centre format incorrect for Student Id:- {0} ", student.Profile.Uid));
                                        validateflag = 1;
                                }

                                //Batch Number Validation
                                if (Regex.IsMatch(student.Profile.BatchNumber.ToString(), @"^[0-9]+$") == false)
                                {
                                    Log(String.Format("Batch Number format incorrect for Student Id:- {0} ", student.Profile.Uid));
                                    validateflag = 1;
                                }

                                //Location Validation
                                if (Regex.IsMatch(student.Profile.OrganisationName.Trim(), @"^[a-zA-Z0-9]*$") == false)
                                {
                                    Log(String.Format("Loacation format incorrect for Student Id:- {0} ", student.Profile.Uid));
                                    validateflag = 1;
                                }

                                //State Validation
                                if (Regex.IsMatch(student.Profile.State.Trim(), @"^[a-zA-Z .]*$") == false)
                                {
                                    Log(String.Format("State format incorrect for Student Id:- {0} ", student.Profile.Uid));
                                    validateflag = 1;
                                }
                               
                                //Parent Name Validation
                                if (Regex.IsMatch(student.Profile.ParentName.Trim(), @"^[a-zA-Z .]*$") == false)
                                {
                                    Log(String.Format("Parent Name format incorrect for Student Id:- {0} ", student.Profile.Uid));
                                    validateflag = 1;
                                }

                                //Permanenet Address Validation
                                //Not needed as address is not mandatory and can have any format.

                                //Parent Contact Validation
                               
                                if (Regex.IsMatch(student.Profile.ParentContact, @"^[0-9]+$") == false || student.Profile.Mobile.Length != 10)
                                {
                                    if (String.IsNullOrEmpty(student.Profile.ParentContact) == false)
                                    {
                                        Log(String.Format("Parent Contact format incorrect for Student Id:- {0} ", student.Profile.Uid));
                                        validateflag = 1;
                                    }

                                }

                                //Batch Start Date Validation
                                
                                DateTime date;
                                if (!DateTime.TryParseExact(student.Profile.BatchStart,
                                   "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None,
                                    out date))
                                {
                                    Log(String.Format("Batch Start  Date format incorrect for Student Id:- {0} ", student.Profile.Uid));
                                    validateflag = 1;
                                }

                                //Batch End Date Validation
                                DateTime date1;
                                if (!DateTime.TryParseExact(student.Profile.BatchEnd,
                                   "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None,
                                    out date1))
                                {
                                    Log(String.Format("Batch End Date format incorrect for Student Id:- {0} ", student.Profile.Uid));
                                    validateflag = 1;
                                }
                            
                                var targetStudent = db.StudentProfiles.Where(p => p.Uid == student.Profile.Uid).SingleOrDefault();
                                var updateflag = 0;

                                if (targetStudent == null) //&& validateflag != 1)
                                {
                                    db.StudentProfiles.Add(student.Profile);
                                    db.SaveChanges();
                                    Log(String.Format("Profile Records added for Student Id:- {0} ", student.Profile.Uid));

                                }
                                else
                                {
                                    updateflag = 1;
                                }

                                if(updateflag==1)// && validateflag!=1)
                                {
                                    //this will set the changed values to the taget student record from the database.
                                    // db.Entry(targetStudent).CurrentValues.SetValues(student);
                                    targetStudent.FirstName = student.Profile.FirstName;
                                    targetStudent.LastName = student.Profile.LastName;
                                    targetStudent.Age = student.Profile.Age;
                                    targetStudent.Gender = student.Profile.Gender;
                                    targetStudent.Education = student.Profile.Education;
                                    targetStudent.Email = student.Profile.Email;
                                    targetStudent.Mobile = student.Profile.Mobile;
                                    targetStudent.MaritalStatus = student.Profile.MaritalStatus;
                                    targetStudent.TrainingCenter = student.Profile.TrainingCenter;
                                    targetStudent.BatchNumber = student.Profile.BatchNumber;
                                    targetStudent.OrganisationName=student.Profile.OrganisationName;
                                    targetStudent.State = student.Profile.State;
                                    targetStudent.ParentOccupation=student.Profile.ParentOccupation;
                                    targetStudent.WorkExperience = student.Profile.WorkExperience;
                                    targetStudent.FamilyMonthlyIncome = student.Profile.FamilyMonthlyIncome;
                                    targetStudent.Prefix = student.Profile.Prefix;
                                    targetStudent.ParentName = student.Profile.ParentName;
                                    targetStudent.PermanentAddress = student.Profile.PermanentAddress;
                                    targetStudent.ParentContact = student.Profile.ParentContact;
                                    targetStudent.BatchStart = student.Profile.BatchStart;
                                    targetStudent.BatchEnd = student.Profile.BatchEnd;
                                    targetStudent.EmploymentStatus = student.Profile.EmploymentStatus;
                                    db.SaveChanges();
                                    Log(String.Format("Profile Records updated for Student Id:- {0} ", student.Profile.Uid));
                                }

                                //db.SaveChanges();

                            }
                            



                            foreach (var placement in result.ImportPlacements)
                            {
                                var targetPlacement = db.Placements.Where(p => p.StudentUid == placement.Placement.StudentUid).SingleOrDefault();
                                if (targetPlacement == null)
                                {
                                    db.Placements.Add(placement.Placement);
                                    db.SaveChanges();
                                    Log(String.Format("Placement Records added for Student Id:- {0} ", placement.Placement.StudentUid));
                                }
                                else
                                {
                                    //this will set the changed values to the taget student record from the database.
                                   // db.Entry(targetPlacement).CurrentValues.SetValues(placement);
                                    
                                    targetPlacement.EmploymentStatus = placement.Placement.EmploymentStatus;
                                    targetPlacement.Position = placement.Placement.Position;
                                    targetPlacement.Company = placement.Placement.Company;
                                    targetPlacement.Location = placement.Placement.Location;                                  
                                    targetPlacement.NatureOfJob = placement.Placement.NatureOfJob;
                                    targetPlacement.Salary = placement.Placement.Salary;
                                    targetPlacement.IfContEduReason = placement.Placement.IfContEduReason;
                                    targetPlacement.IfDropOutReason = placement.Placement.IfDropOutReason;
                                    targetPlacement.UpdatedContact = placement.Placement.UpdatedContact;
                                    db.SaveChanges();
                                    Log(String.Format("Placement Records updated for Student Id:- {0} ", placement.Placement.StudentUid));
                                }

                              
                                
                            }



                            foreach (var postplacement in result.ImportPostPlacements)
                            {
                                var targetPostPlacement = db.PostPlacement.Where(p => p.StudentUid == postplacement.PostPlacement.StudentUid).SingleOrDefault();
                                if (targetPostPlacement == null)
                                {
                                    db.PostPlacement.Add(postplacement.PostPlacement);
                                    db.SaveChanges();
                                    Log(String.Format("Post Placement Records added for Student Id:- {0} ", postplacement.PostPlacement.StudentUid));
                                }
                                else
                                {
                                    //this will set the changed values to the taget student record from the database.
                                    // db.Entry(targetPlacement).CurrentValues.SetValues(placement);

                                    targetPostPlacement.ContinueJob = postplacement.PostPlacement.ContinueJob;
                                    targetPostPlacement.Company = postplacement.PostPlacement.Company;
                                    targetPostPlacement.Position = postplacement.PostPlacement.Position;
                                    targetPostPlacement.Salary = postplacement.PostPlacement.Salary;
                                    targetPostPlacement.UpdatedContact = postplacement.PostPlacement.UpdatedContact;
                                    db.SaveChanges();
                                    Log(String.Format("Post Placement Records updated for Student Id:- {0} ", postplacement.PostPlacement.StudentUid));
                                }



                            }
                           
                            Log("-------------------------------------------");
                        }
                
                    }
            }
       }
            
                    catch (Exception ex)
                    {                        
                        //Log(ex.StackTrace);
                        Log("---------------------------------------------------------------------------");
                        Log("---------------------------------------------------------------------------");
                        Log("                                                                           ");
                        Log(ex.Message);
                        Log("Oops! Something went wrong with the import. Contact Abhijeet Mehta");                    
            }
    }
        

          private void buttonImportScores_Click(object sender, EventArgs e)
          {
              if (String.IsNullOrEmpty(openFileDialog1.FileName))
              {
                  Log("No file selected for import. Select file...");
                  return;
              }
              try
              {
                  var existingFile = openFileDialog1.OpenFile();
                  // Open and read the XlSX file.
                  const int startRow = 1;
                  using (var package = new ExcelPackage(existingFile))
                  {
                      // Get the work book in the file
                      ExcelWorkbook workBook = package.Workbook;
                      if (workBook != null)
                      {
                          if (workBook.Worksheets.Count > 0)
                          {

                              // Get the first worksheet
                              ExcelWorksheet currentWorksheet = workBook.Worksheets.First();
                              var SubjectInfo = new Subject();

                              int temp = 1;
                              for (int i = 1; i <= currentWorksheet.Dimension.End.Column; i++)
                              {
                                  if (currentWorksheet.Cells[startRow, i].Value != null)
                                  {

                                      SubjectInfo.SubjectName = currentWorksheet.Cells[startRow, i].Value.ToString();

                                      for (int j = temp; j <= currentWorksheet.Dimension.End.Column; j++)
                                      {
                                          if (currentWorksheet.Cells[startRow + 1, j].Value.ToString().Trim().Equals("Course Total"))
                                          {
                                              temp = j + 1;

                                              SubjectInfo.TotalMarks = currentWorksheet.Cells[startRow + 2, j].Value.ToString();
                                              var targetSubject = db.Subjects.Where(p => p.SubjectName == SubjectInfo.SubjectName).SingleOrDefault();
                                              if (targetSubject == null)
                                              {
                                                  db.Subjects.Add(SubjectInfo);
                                                  db.SaveChanges();
                                                  Log(String.Format("Subjects Added:- {0} ", SubjectInfo.SubjectName));
                                                  break;
                                              }
                                              else
                                              {
                                                  targetSubject.TotalMarks = SubjectInfo.TotalMarks;
                                                  db.SaveChanges();
                                                  Log(String.Format("Subjects Updated:- {0} ", SubjectInfo.SubjectName));
                                                  break;
                                              }

                                          }
                                      }
                                  }
                              }

                              int k = 6;
                              SubjectScore SubjectScoreInfo = new SubjectScore();
                              for (int i = 1; i <= currentWorksheet.Dimension.End.Column; i++)
                              {
                                  if (currentWorksheet.Cells[startRow, i].Value != null)
                                  {
                                      SubjectScoreInfo.Subject = currentWorksheet.Cells[startRow, i].Value.ToString();

                                      for (int j = k; j <= currentWorksheet.Dimension.End.Column; j++)
                                      {

                                          if (Convert.ToInt32(currentWorksheet.Cells[startRow + 2, j].Value) != 1)
                                          {
                                              if (currentWorksheet.Cells[startRow + 1, j].Value.ToString().Trim().Equals("Course Total"))
                                              {
                                                  k = j + 1;
                                                  break;
                                              }
                                              else
                                              {
                                                  if (currentWorksheet.Cells[startRow + 1, j].Value.ToString().Trim().Equals("%") == false)
                                                  {
                                                      if (currentWorksheet.Cells[startRow + 1, j].Value.ToString().Trim().Equals("CGPA") == false)
                                                      {

                                                          for (int n = 4; n <= currentWorksheet.Dimension.End.Row; n++)
                                                          {
                                                              if ((String)(currentWorksheet.Cells[n, 3].Value) == null)
                                                              {
                                                              }
                                                              else
                                                              {
                                                                  SubjectScoreInfo.StudentUID = (String)(currentWorksheet.Cells[n, 3].Value);
                                                                  SubjectScoreInfo.Lessons = currentWorksheet.Cells[startRow + 1, j].Value.ToString();
                                                                  SubjectScoreInfo.Score = Convert.ToInt32(currentWorksheet.Cells[n, j].Value);
                                                                  var targetScoreInfo = db.SubjectScores.Where(p => p.StudentUID == SubjectScoreInfo.StudentUID && p.Subject == SubjectScoreInfo.Subject && p.Lessons == SubjectScoreInfo.Lessons).SingleOrDefault();
                                                                  if (targetScoreInfo == null)
                                                                  {
                                                                      db.SubjectScores.Add(SubjectScoreInfo);
                                                                      db.SaveChanges();
                                                                      Log(String.Format("Student Id:- {0} Subjects :- {1} Lession :- {2}, Score :-{3} added", SubjectScoreInfo.StudentUID, SubjectScoreInfo.Subject, SubjectScoreInfo.Lessons, SubjectScoreInfo.Score));

                                                                  }
                                                                  else
                                                                  {
                                                                      targetScoreInfo.Subject = SubjectScoreInfo.Subject;
                                                                      targetScoreInfo.Lessons = SubjectScoreInfo.Lessons;
                                                                      targetScoreInfo.Score = SubjectScoreInfo.Score;
                                                                      db.SaveChanges();
                                                                      Log(String.Format("Student Id:- {0} Subjects :- {1} Lession :- {2}, Score :-{3} updated", SubjectScoreInfo.StudentUID, SubjectScoreInfo.Subject, SubjectScoreInfo.Lessons, SubjectScoreInfo.Score));

                                                                  }
                                                              }

                                                          }

                                                      }

                                                  }
                                              }

                                          }
                                      }

                                  }
                              }
                          }
                      }
                  }
                  Log("All records successfully saved");
              }
              catch (Exception excp)
              {
                  //Log(excp.StackTrace);
                  Log("---------------------------------------------------------------------------");
                  Log("---------------------------------------------------------------------------");
                  Log("                                                                           ");
                  Log("File format is not correct");
                  Log("Restart and try again");
                  Log(excp.Message);
                  Log("Oops! Something went wrong with the import. Contact Abhijeet Mehta");
                  
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

        private void generateReport_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.Show();
            
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

       



    }

}
