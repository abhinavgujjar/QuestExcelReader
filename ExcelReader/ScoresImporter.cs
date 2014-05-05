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
    public class ScoresImporter : IImporter
    {
        ExcelWorksheet _workSheet;

        JToken _config;

        WorksheetHelper _helper;

        public ScoresImporter(ExcelWorksheet workSheet, JToken config)
        {
            _workSheet = workSheet;
            _config = config;
            _helper = new WorksheetHelper(_workSheet, _config);
        }

        public ImportResult Import()
        {
            ImportPlacement();
            ImportScores();

            return new ImportResult();
        }
        public ValidationResult Validate()
        {
            //ensure that all the records have a student id
            return new ValidationResult() { Valid = true };
        }


        private void ImportPlacement()
        {
            string studentIdColumn = (string)_config["studentIdColumn"];
            string trainingCentreColumn = (string)_config["trainingCentreColumn"];

            int dataRowStart = (int)_config["dataRowStart"];

            int rowIndex = dataRowStart;

            QSStagingDbContext db = new QSStagingDbContext();

            while (true)
            {
                //reached the end of the records
                if (_workSheet.Cells[rowIndex, 1].FirstOrDefault() == null) { break; }

                if (_workSheet.Cells[rowIndex, 1].FirstOrDefault().Value == null) { break; }
                
                if (String.IsNullOrWhiteSpace(_workSheet.Cells[rowIndex, 1].FirstOrDefault().Value.ToString())) { break; }


                String studentId = _workSheet.Cells[studentIdColumn + rowIndex.ToString()].First().Value as string;
                String trainingCentre = _workSheet.Cells[trainingCentreColumn + rowIndex.ToString()].First().Value as string;

                var studentProfile = db.StudentProfiles.Where(
                    p => p.LegacyUid == studentId
                    && p.TrainingCenter == trainingCentre).SingleOrDefault();

                if (studentProfile != null)
                {

                    var placementRecord = db.Placements.Where(
                        p => p.StudentUid == studentProfile.Uid).SingleOrDefault();

                    if (placementRecord == null)
                    {
                        placementRecord = new Placement();
                        placementRecord.StudentUid = studentProfile.Uid;

                        db.Placements.Add(placementRecord);
                    }

                    placementRecord.OfferLetter = _helper.getCellValue("OfferLetter", rowIndex);
                    placementRecord.CourseCompletionStatus = _helper.getCellValue("CourseCompletionStatus", rowIndex);
                    placementRecord.Company = _helper.getCellValue("Company", rowIndex);
                    placementRecord.EmploymentStatus = _helper.getCellValue("EmploymentStatus", rowIndex);
                    placementRecord.Comments = _helper.getCellValue("Comments", rowIndex);
                    placementRecord.Position = _helper.getCellValue("Position", rowIndex);
                    placementRecord.Salary = _helper.getCellValue("Salary", rowIndex);
                    placementRecord.Location = _helper.getCellValue("Location", rowIndex);
                }
                rowIndex++;
            }

            db.SaveChanges();
        }

        public bool ImportScores()
        {
            QSStagingDbContext db = new QSStagingDbContext();
            var students = new List<StudentProfile>();
            string studentIdColumn = (string)_config["studentIdColumn"];
            string trainingCentreColumn = (string)_config["trainingCentreColumn"];
            int dataRowStart = (int)_config["dataRowStart"];
            int subjectRowIndex = (int)_config["subjectRowIndex"];
            int categoryRowIndex = (int)_config["categoryRowIndex"];
            int columnStartIndex = (int)_config["columnStartIndex"];

            int subjectTotalRow = subjectRowIndex + 1;

            //determine number of columns to traverse
            int columnIndex = columnStartIndex;

            while (true)
            {
                if (_workSheet.Cells[subjectRowIndex, columnIndex].FirstOrDefault() == null)
                {
                    break;
                }

                var subjectName = _workSheet.Cells[subjectRowIndex, columnIndex].First().Value as string;
                var category = _workSheet.Cells[categoryRowIndex, columnIndex].First().Value as string;

                if (skipSubject(subjectName, _config))
                {
                    columnIndex++;
                    continue;
                };

                double? subjectTotal = _workSheet.Cells[subjectTotalRow, columnIndex].First().Value as double?;
                var targetSubject = db.Subjects.Where(s => s.Name.ToLower() == subjectName).SingleOrDefault();

                if (targetSubject == null)
                {
                    //create an entry 
                    targetSubject = new Subject()
                    {
                        Name = subjectName
                    };

                    db.Subjects.Add(targetSubject);
                    db.SaveChanges();
                }

                //pick up scores for the subject for each student
                int rowIndex = dataRowStart;
                while (true)
                {
                    if (_workSheet.Cells[rowIndex, 1].FirstOrDefault() == null)
                    {
                        break;
                    }

                    String studentId = _workSheet.Cells[studentIdColumn + rowIndex.ToString()].First().Value as string;
                    String trainingCentre = _workSheet.Cells[trainingCentreColumn + rowIndex.ToString()].First().Value as string;

                    var studentProfile = db.StudentProfiles.Where(
                    p => p.LegacyUid == studentId
                    && p.TrainingCenter == trainingCentre).SingleOrDefault();

                    if (studentId == null)
                    {
                        Console.WriteLine("could not find student Id, skipping column");
                    }
                    else
                    {
                        if (_workSheet.Cells[rowIndex, columnIndex].FirstOrDefault() != null
                            && _workSheet.Cells[rowIndex, columnIndex].First().Value != null)
                        {
                            var subjectScore = db.SubjectScores.Where(s => s.StudentUID == studentId && s.SubjectId == targetSubject.Id).SingleOrDefault();

                            if (subjectScore == null)
                            {
                                subjectScore = new SubjectScore()
                                {
                                    SubjectId = targetSubject.Id,
                                    StudentUID = studentId
                                };

                                db.SubjectScores.Add(subjectScore);
                            }

                            int score = Convert.ToInt32(_workSheet.Cells[rowIndex, columnIndex].First().Value);

                            subjectScore.Score = score;
                            subjectScore.Subject = subjectName;
                            subjectScore.Total = subjectTotal.HasValue ? Convert.ToInt32(subjectTotal.Value) : 0;

                            
                        }
                    }
                    rowIndex++;
                }

                columnIndex++;
            }

            db.SaveChanges();

            return true;
        }

        private bool skipSubject(string subjectName, JToken _config)
        {
            var shouldSkip = false;
            foreach ( var item in _config["skipColumns"])
            {
                if ( (string)item == subjectName || subjectName.ToLower().Contains(((string)item).ToLower() ) )
                {
                    shouldSkip = true;
                }
            }
            return shouldSkip;
        }

        
    }
}
