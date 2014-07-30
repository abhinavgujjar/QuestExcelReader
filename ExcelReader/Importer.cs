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
            List<FullStudent> imports1 = new List<FullStudent>();
            List<FullStudent> imports2 = new List<FullStudent>();

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
                    FullStudent placementRecord = new FullStudent();
                    FullStudent postPlacementRecord = new FullStudent();
                    var profile = importProfile(rowIndex);
                    var placement = importPlacement(rowIndex);
                    var postplacement = importPostPlacement(rowIndex); 
                    placementRecord.Placement = placement;
                    student.Profile = profile;
                    postPlacementRecord.PostPlacement = postplacement; 
                    imports.Add(student);
                    imports1.Add(placementRecord);
                    imports2.Add(postPlacementRecord);
                }

                    

                catch (Exception e)
                {
                    throw new ApplicationException("Error loading at row : " + rowIndex, e);
                }

                rowIndex++;
                first = false;
            }

            result.ImportStudents = imports;
            result.ImportPlacements = imports1;
            result.ImportPostPlacements = imports2;
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

        /*private Placement importPlacement(int rowIndex)
        {
            var placementRecord = new Placement();
            //placementRecord.StudentUid = studentProfile.Uid;
           // placementRecord.StudentId = Convert.ToInt32(_helper.getCellValue("StudentId", rowIndex));
            //placementRecord.StudentUid = _helper.getCellValue("StudentUid", rowIndex);
            placementRecord.OfferLetter = _helper.getCellValue("OfferLetter", rowIndex);
            placementRecord.CourseCompletionStatus = _helper.getCellValue("CourseCompletionStatus", rowIndex);
            placementRecord.Company = _helper.getCellValue("Company", rowIndex);
            placementRecord.EmploymentStatus = _helper.getCellValue("EmploymentStatus", rowIndex);
            placementRecord.Comments = _helper.getCellValue("Comments", rowIndex);
            placementRecord.Position = _helper.getCellValue("Position", rowIndex);
            placementRecord.Salary = _helper.getCellValue("Salary", rowIndex);
            placementRecord.Location = _helper.getCellValue("CompanyLocation", rowIndex);

            return placementRecord;
        }*/

        private StudentProfile importProfile(int rowIndex)
        {
            ImportResult result = new ImportResult();

            StudentProfile student = new StudentProfile();

            GetExcelRow(rowIndex, student);

            return student;
        }

        private Placement importPlacement(int rowIndex)
        {
            ImportResult result = new ImportResult();

            Placement placementRecord = new Placement();

            GetExcelRow(rowIndex, placementRecord);

            return placementRecord;
        }

        private Post_Placement importPostPlacement(int rowIndex)
        {
            ImportResult result = new ImportResult();

            Post_Placement postPlacementRecord = new Post_Placement();

            GetExcelRow(rowIndex, postPlacementRecord);

            return postPlacementRecord;
        }

        private void GetExcelRow(int rowIndex, Placement placementRecord)
        {
                placementRecord.StudentUid = _helper.getCellValue("StudentId", rowIndex);
                placementRecord.EmploymentStatus = _helper.getCellValue("EmploymentStatus", rowIndex);
                placementRecord.Position = _helper.getCellValue("Designation", rowIndex);
                placementRecord.Company = _helper.getCellValue("Company", rowIndex);
                placementRecord.Location = _helper.getCellValue("CompanyLocation", rowIndex);
                placementRecord.NatureOfJob = _helper.getCellValue("JobType", rowIndex);
                placementRecord.Salary = (_helper.getCellValue("Salary", rowIndex));
                placementRecord.IfContEduReason = _helper.getCellValue("ContinueEducation", rowIndex);
                placementRecord.IfDropOutReason = _helper.getCellValue("DropOutReason", rowIndex);
                placementRecord.UpdatedContact = _helper.getCellValue("UpdatedContactDetail", rowIndex);
           
        }

        private void GetExcelRow(int rowIndex, Post_Placement postPlacementRecord)
        {
            postPlacementRecord.StudentUid = _helper.getCellValue("StudentId", rowIndex);
            postPlacementRecord.ContinueJob = _helper.getCellValue("ContinueJob", rowIndex);
            postPlacementRecord.Company = _helper.getCellValue("PostCompany", rowIndex);
            postPlacementRecord.Position = _helper.getCellValue("PostDesignation", rowIndex);
            postPlacementRecord.Salary = _helper.getCellValue("PostSalary", rowIndex);
            postPlacementRecord.UpdatedContact = _helper.getCellValue("PostUpdatedContact", rowIndex);
        }


            
        private void GetExcelRow(int rowIndex, StudentProfile student)
        {
            student.FirstName = _helper.getCellValue("FirstName", rowIndex);
            student.LastName = _helper.getCellValue("LastName", rowIndex);
            student.OrganisationName = _helper.getCellValue("Oraganisation", rowIndex);
            student.TrainingCenter = _helper.getCellValue("TrainingCentre", rowIndex);
            student.BatchNumber = _helper.getCellValue("BatchNumber", rowIndex);
            student.Prefix = _helper.getCellValue("Prefix", rowIndex);
            student.Uid = _helper.getCellValue("StudentId", rowIndex);
            student.Gender = _helper.getCellValue("Gender", rowIndex);        
            student.Education = _helper.getCellValue("Education", rowIndex);
            student.MaritalStatus = _helper.getCellValue("MaritalStatus", rowIndex);
            student.Email = _helper.getCellValue("Email", rowIndex);
            student.Age = Convert.ToInt32(_helper.getCellValue("Age", rowIndex));
            student.WorkExperience = _helper.getCellValue("WorkExperience", rowIndex);
            student.ParentName = _helper.getCellValue("ParentName", rowIndex);
            student.ParentContact = _helper.getCellValue("ParentContact", rowIndex);
            student.PermanentAddress = _helper.getCellValue("Address", rowIndex);
            student.FamilyMonthlyIncome = _helper.getCellValue("FamilyIncome", rowIndex);
            student.ParentOccupation = _helper.getCellValue("ParentOccupation", rowIndex);        
            student.BatchStart = _helper.getCellValue("BatchStart", rowIndex);
            student.BatchEnd = _helper.getCellValue("BatchEnd", rowIndex);
            student.State = _helper.getCellValue("State", rowIndex);
            student.Mobile = _helper.getCellValue("Mobile", rowIndex);
            student.EmploymentStatus = _helper.getCellValue("EmploymentStatus", rowIndex);

        }


    }
}
