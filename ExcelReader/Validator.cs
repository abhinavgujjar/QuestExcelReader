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
    class Validator
    {
        ExcelWorksheet _workSheet;

        JToken _config;
        private WorksheetHelper _helper;

        public Validator(ExcelWorksheet workSheet, JToken config)
        {
            _workSheet = workSheet;
            _config = config;

            _helper = new WorksheetHelper(_workSheet, _config);

        }

        public ValidationResult Validate()
        {
            ValidationResult result = new ValidationResult();
            result.Valid = true;
            var worksheet = _workSheet;

            var dataStartRow = (int)_config["dataRowStart"];
            var centreName = _helper.getCellValue("TrainingCentre", dataStartRow);
            var batchNumber = _helper.getCellValue("BatchNumber", dataStartRow);
            var location = _helper.getCellValue("Location", dataStartRow);

            var duplicateValidation = ValidateDuplicates(centreName, Convert.ToInt32(batchNumber));
            if (!duplicateValidation.Valid)
            {
                return duplicateValidation;
            }

            int index = dataStartRow;
            var misMatch = false;
            while (true)
            {
                if (worksheet.Cells[index, 1].First().Value == null)
                {
                    //reached the end of records
                    break;
                }

                var indexCentre = _helper.getCellValue("TrainingCentre", index).Trim();
                var indexBatch = _helper.getCellValue("BatchNumber", index).Trim();
                var indexLocation = _helper.getCellValue("Location", index).Trim(); 

                if (string.Compare(indexCentre, centreName.Trim(), true) != 0  ||
                    string.Compare(indexBatch, batchNumber.Trim(), true) != 0 ||
                    string.Compare(indexLocation, location.Trim(), true) != 0)
                {
                    result.Valid = false;
                    Console.WriteLine("Centre information does not match for all the records. Please correct the excel sheet before uploading");
                    result.Message = String.Format("Mismatch Found : {0}|{1} ; {2}|{3}; {4}|{5}", indexCentre, centreName, indexBatch, batchNumber, indexLocation, location);
                    break;
                }
                index++;
            }

            if (result.Valid)
            {
                Console.WriteLine(string.Format("{0} records found", index - dataStartRow));
                Console.WriteLine(String.Format("Upload for Centre {0} ", centreName));
                result.Message = String.Format("Centre: {0}, Batch: {1}, Location: {2}", centreName, batchNumber, location);
            }

            return result;
        }

        public ValidationResult ValidateDuplicates(string centre, int batch)
        {
            //see if there are any records for the same center and batch
            var result = new ValidationResult();

            QSStagingDbContext db = new QSStagingDbContext();
            var duplicates = db.StudentProfiles.Where(p => p.TrainingCenter.ToLower().Trim() == centre && p.BatchNumer == batch).Count();

            if (duplicates > 0)
            {
                result.Valid = false;
                result.Message = String.Format("Existing records already found for Centre: {0} and batch: {1} ", centre, batch);
            }
            else
            {
                result.Valid = true;
            }
            return result;
        }

        
    }
}
