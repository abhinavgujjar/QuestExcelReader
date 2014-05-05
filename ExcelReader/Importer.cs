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

        public int NumberOfBufferRecords { get; set; }

        public Importer(ExcelWorksheet workSheet, JToken config)
        {
            _workSheet = workSheet;
            _config = config;
            _helper = new WorksheetHelper(_workSheet, _config);

            NumberOfBufferRecords = (int)_config["numberOfBufferRecords"];
        }

        public ValidationResult Validate()
        {
            Validator validator = new Validator(_workSheet, _config);
            return validator.Validate();

        }

        public ImportResult Import()
        {
            ImportResult result = new ImportResult();
            
            bool recordsExist = checkIfRecordsExistForBatch();

            QSStagingDbContext db = new QSStagingDbContext();

            var students = new List<StudentProfile>();
            int rowIndex = (int)_config["dataRowStart"];

            while (true)
            {
                if (_workSheet.Cells[rowIndex, 1].First().Value == null)
                {
                    //reached the end of records
                    break;
                }

                var uid = _helper.getCellValue("Uid", rowIndex);

                StudentProfile student = null;
                if (!String.IsNullOrEmpty(uid))
                {
                    student = db.StudentProfiles.Where(s => s.Uid == uid).SingleOrDefault();
                }

                if (student == null)
                {
                    student = new StudentProfile();
                }

                GetExcelRow(rowIndex, student);

                if (String.IsNullOrEmpty(uid))
                {
                    uid = UidGenerator.GenerateUid(student.Name, student.TrainingCenter, student.Location, student.BatchNumer);

                    if (db.StudentProfiles.Where(s => s.Uid == uid).FirstOrDefault() != null)
                    {
                        uid = uid + rowIndex;
                    }
                    

                    //check for conflict
                    if (students.Where(s => s.Uid == uid).Count() > 0)
                    {
                        uid = uid + rowIndex;
                    }

                    student.Uid = uid;

                    //update the excel sheet as well
                    _helper.updateCellValue("Uid", rowIndex, uid);

                    db.StudentProfiles.Add(student);
                }

                students.Add(student);


                rowIndex++;
            }

            result.NumberOfRecords = rowIndex;
            result.Message = "Import Successful";

            //var sampleStudent = students.First();

            ////add some buffer records
            //if (!recordsExist)
            //{
            //    for (int i = 0; i < NumberOfBufferRecords; i++)
            //    {
            //        var uid = UidGenerator.GenerateUid("buffer", sampleStudent.TrainingCenter, sampleStudent.Location, sampleStudent.BatchNumer);
            //        uid = uid + i;

            //        var student = getBufferEntry(sampleStudent, i, uid);

            //        students.Add(student);

            //        AddExcelRow(rowIndex, i, uid, student);
            //    }
            //}

            

            return result;
        }

        private bool checkIfRecordsExistForBatch()
        {
            var dataStartRow = (int)_config["dataRowStart"];
            var centreName = _helper.getCellValue("TrainingCentre", dataStartRow);
            var batchNumber = Convert.ToInt32( _helper.getCellValue("BatchNumber", dataStartRow));
            var location = _helper.getCellValue("Location", dataStartRow);

            using (QSStagingDbContext db = new QSStagingDbContext())
            {
                var matchingRecords = db.StudentProfiles.Where(p => p.TrainingCenter == centreName 
                    && p.BatchNumer ==  batchNumber
                    && p.Location == location);

                return matchingRecords.Count() > 0;
            }
        }

        private void GetExcelRow(int rowIndex, StudentProfile student)
        {
            student.Name = _helper.getCellValue("Name", rowIndex);
            student.MobileNumber = _helper.getCellValue("Mobile", rowIndex);
            student.Gender = _helper.getCellValue("Gender", rowIndex);
            student.Email = _helper.getCellValue("Email", rowIndex);
            student.Age = Convert.ToInt32(_helper.getCellValue("Age", rowIndex));
            student.Demographics = _helper.getCellValue("Demographics", rowIndex);
            student.Education = _helper.getCellValue("Education", rowIndex);
            student.EmploymentStatus = _helper.getCellValue("EmploymentStatus", rowIndex);
            student.FamilyMonthlyIncome = _helper.getCellValue("FamilyIncome", rowIndex);
            student.FullAddress = _helper.getCellValue("PresentAddress", rowIndex);
            student.PermanentAddress = _helper.getCellValue("PermanentAddress", rowIndex);
            student.State = _helper.getCellValue("State", rowIndex);
            student.Location = _helper.getCellValue("Location", rowIndex);
            student.ParentMobileNumber = _helper.getCellValue("ParentMobileNumber", rowIndex);
            student.TrainingCenter = _helper.getCellValue("TrainingCentre", rowIndex);
            student.BatchNumer = Convert.ToInt32(_helper.getCellValue("BatchNumber", rowIndex));
            student.LegacyUid = _helper.getCellValue("LegacyUid", rowIndex);
            
        }

        private void AddExcelRow(int rowIndex, int i, string uid, StudentProfile student)
        {
            _helper.updateCellValue("Uid", rowIndex + i, uid);
            _helper.updateCellValue("Name", rowIndex + i, student.Name);
            _helper.updateCellValue("Mobile", rowIndex + i, student.MobileNumber);
            _helper.updateCellValue("Gender", rowIndex + i, student.Gender);
            _helper.updateCellValue("Email", rowIndex + i, student.Email);
            _helper.updateCellValue("Age", rowIndex + i, student.Age.ToString());
            _helper.updateCellValue("Demographics", rowIndex + i, student.Demographics);
            _helper.updateCellValue("Education", rowIndex + i, student.Education);
            _helper.updateCellValue("EmploymentStatus", rowIndex + i, student.EmploymentStatus);
            _helper.updateCellValue("FamilyIncome", rowIndex + i, student.FamilyMonthlyIncome);
            _helper.updateCellValue("PresentAddress", rowIndex + i, student.FullAddress);
            _helper.updateCellValue("PermanentAddress", rowIndex + i, student.PermanentAddress);
            _helper.updateCellValue("State", rowIndex + i, student.State);
            _helper.updateCellValue("Location", rowIndex + i, student.Location);
            _helper.updateCellValue("ParentMobileNumber", rowIndex + i, student.ParentMobileNumber);
            _helper.updateCellValue("TrainingCentre", rowIndex + i, student.TrainingCenter);
            _helper.updateCellValue("BatchNumber", rowIndex + i, student.BatchNumer.ToString());
        }

        private static StudentProfile getBufferEntry(StudentProfile sampleStudent, int i, string uid)
        {
            var student = new StudentProfile()
            {
                Uid = uid,
                BatchNumer = sampleStudent.BatchNumer,
                TrainingCenter = sampleStudent.TrainingCenter,
                Location = sampleStudent.Location,
                MobileNumber = "0000000000",
                Name = "Buffer" + i,
                EmploymentStatus = sampleStudent.EmploymentStatus,
                ParentMobileNumber = "0000000000",
                PermanentAddress = "unknown",
                FullAddress = "unknown",
                Education = "unknown",
                Gender = "Unknown",
                Demographics = "unknown",
                Email = "unknown",
                Age = 0,
                State = sampleStudent.State,
                FamilyMonthlyIncome = "unknown"
            };
            return student;
        }


    }
}
