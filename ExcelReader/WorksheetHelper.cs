using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    class WorksheetHelper
    {
        private ExcelWorksheet _workSheet;
        JToken _config;

        public WorksheetHelper(ExcelWorksheet workSheet, JToken config)
        {
            _workSheet = workSheet;
            _config = config;
        }

        public string getCellValue(string columnName, int rowIndex)
        {
            var columnLetter = (string)_config["columns"][columnName];
            
            if (columnLetter == null)
            {
                return string.Empty;
            }

            var rawValue = _workSheet.Cells[columnLetter + rowIndex].FirstOrDefault();

            var value = rawValue != null && rawValue.Value != null ? rawValue.Value.ToString() : String.Empty;

            return value;
        }

        public void updateCellValue(string columnName, int rowIndex, string value)
        {
            var columnLetter = (string)_config["columns"][columnName];

            if (columnLetter == null)
            {
                return;
            }

            _workSheet.Cells[columnLetter + rowIndex].Value = value;
        }
    }
}
