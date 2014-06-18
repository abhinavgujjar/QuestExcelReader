using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using ExcelReader.Models;
using Newtonsoft.Json;
using System.IO;
using Newtonsoft.Json.Linq;

namespace ExcelReader
{
    public class Importer : ExcelReader.IImporter
    {
        ExcelWorksheet _workSheet;

        JToken _config;

        WorksheetHelper _helper;

        QSStagingDbContext db = new QSStagingDbContext();


        public Importer(ExcelWorksheet workSheet, JToken config)
        {
            _workSheet = workSheet;
            _config = config;
            _helper = new WorksheetHelper(_workSheet, _config);

        }

        public ValidationResult Validate()
        {
            try
            {
                Validator validator = new Validator(_workSheet, _config);
                return validator.Validate();
            }
            catch (Exception e)
            {
                throw new ApplicationException("Error during validation. Are you sure the file is formatted corectly?", e);
            }

        }

        public ImportResult Import()
        {
            ImportResult result = new ImportResult();

            int rowIndex = (int)_config["dataRowStart"];

            bool first = true;
            List<string> messages = new List<string>();
            List<FullStudent> imports = new List<FullStudent>();

            while (true)
            {
                var studentId = _helper.getCellValue("StudentId", rowIndex);
                if (String.IsNullOrWhiteSpace(studentId))
                {
                    //reached the end of records
                    break;
                }


                try
                {
                    FullStudent student = new FullStudent();

                    var profile = importProfile(rowIndex);
                    student.Profile = profile;

                    imports.Add(student);
                   
                }
                catch (Exception e)
                {
                    throw new ApplicationException("Error loading at row : " + rowIndex, e);
                }

                rowIndex++;
                first = false;
            }

            result.ImportStudents = imports;

            return result;
        }

        private bool skipSubject(string subjectName, JToken _config)
        {
            var shouldSkip = false;
            foreach (var item in _config["skipColumns"])
            {
                if ((string)item == subjectName || subjectName.ToLower().Contains(((string)item).ToLower()))
                {
                    shouldSkip = true;
                }
            }
            return shouldSkip;
        }

        private List<Legacy_SubjectScore> importScores(int rowIndex, int studentId, bool first)
        {
            List<Legacy_SubjectScore> scores = new List<Legacy_SubjectScore>();

            int dataRowStart = (int)_config["dataRowStart"];
            int subjectRowIndex = (int)_config["subjectRowIndex"];
            int categoryRowIndex = (int)_config["categoryRowIndex"];
            int columnStartIndex = (int)_config["columnStartIndex"];
            //determine number of columns to traverse
            int columnIndex = columnStartIndex;
            var category = string.Empty;

            while (true)
            {
                //detect end of 
                if (_workSheet.Cells[subjectRowIndex, columnIndex].FirstOrDefault() == null || _workSheet.Cells[subjectRowIndex, columnIndex].First().Value == null)
                {
                    break;
                }

                var subjectName = Convert.ToString(_workSheet.Cells[subjectRowIndex, columnIndex].First().Value);

                var newCategory = Convert.ToString(_workSheet.Cells[categoryRowIndex, columnIndex].First().Value);
                if (!string.IsNullOrWhiteSpace(newCategory))
                {
                    category = newCategory;
                }

                if (skipSubject(subjectName, _config))
                {
                    columnIndex++;
                    continue;
                };

                if (_workSheet.Cells[rowIndex, columnIndex].FirstOrDefault() != null
                    && _workSheet.Cells[rowIndex, columnIndex].First().Value != null)
                {
                    var rawScore = Convert.ToString(_workSheet.Cells[rowIndex, columnIndex].First().Value);

                    decimal score;
                    decimal.TryParse(rawScore, out score);
                    
                    var subjectScore = new Legacy_SubjectScore()
                     {
                         Score = score,
                         StudentId = studentId,
                         Subject = subjectName,
                         Category = category
                     };

                    scores.Add(subjectScore);
                }

                columnIndex++;
            }

            return scores;
        }

        private Placement importPlacement(int rowIndex)
        {
            var placementRecord = new Placement();
            //placementRecord.StudentUid = studentProfile.Uid;

            placementRecord.OfferLetter = _helper.getCellValue("OfferLetter", rowIndex);
            placementRecord.CourseCompletionStatus = _helper.getCellValue("CourseCompletionStatus", rowIndex);
            placementRecord.Company = _helper.getCellValue("Company", rowIndex);
            placementRecord.EmploymentStatus = _helper.getCellValue("EmploymentStatus", rowIndex);
            placementRecord.Comments = _helper.getCellValue("Comments", rowIndex);
            placementRecord.Position = _helper.getCellValue("Position", rowIndex);
            placementRecord.Salary = _helper.getCellValue("Salary", rowIndex);
            placementRecord.Location = _helper.getCellValue("CompanyLocation", rowIndex);

            return placementRecord;
        }

        private StudentProfile importProfile(int rowIndex)
        {
            ImportResult result = new ImportResult();

            StudentProfile student = new StudentProfile();

            GetExcelRow(rowIndex, student);

            return student;
        }

        private void GetExcelRow(int rowIndex, StudentProfile student)
        {
            student.FirstName = _helper.getCellValue("FirstName", rowIndex);
            student.LastName = _helper.getCellValue("LastName", rowIndex);
            student.Uid = _helper.getCellValue("StudentId", rowIndex);
            student.Gender = _helper.getCellValue("Gender", rowIndex);
            student.MobileNumber = _helper.getCellValue("Mobile", rowIndex);
            student.Education = _helper.getCellValue("Education", rowIndex);
            student.MaritalStatus = _helper.getCellValue("MaritalStatus", rowIndex);
            student.Email = _helper.getCellValue("Email", rowIndex);
            student.Age = Convert.ToInt32(_helper.getCellValue("Age", rowIndex));
            student.WorkExperience = _helper.getCellValue("WorkExperience", rowIndex);
            student.ParentName = _helper.getCellValue("ParentName", rowIndex);
            student.ParentContact = _helper.getCellValue("ParentContact", rowIndex);
            student.PermanentAddress = _helper.getCellValue("Address", rowIndex);
            student.FamilyMonthlyIncome = _helper.getCellValue("FamilyIncome", rowIndex);
            student.TrainingCenter = _helper.getCellValue("TrainingCentre", rowIndex);
            student.Location = _helper.getCellValue("Location", rowIndex);
            student.BatchStart = _helper.getCellValue("BatchStart", rowIndex);
            student.BatchEnd = _helper.getCellValue("BatchEnd", rowIndex);
            student.Demographics = _helper.getCellValue("Demographics", rowIndex);
            student.EmploymentStatus = _helper.getCellValue("EmploymentStatus", rowIndex);
            student.State = _helper.getCellValue("State", rowIndex);

        }


    }
}
