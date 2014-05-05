using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelReader.Models
{
    public class StudentProfile
    {
        public int Id { get; set; }
        public string Uid { get; set; }
        public string LegacyUid { get; set; }


        public string Name { get; set; }
        public int Age { get; set; }
        public string Gender { get; set; }
        public String Education { get; set; }
        public string Email { get; set; }
        public string MobileNumber { get; set; }
        public string EmploymentStatus { get; set; }
        public string TrainingCenter { get; set; }
        public int BatchNumer { get; set; }
        public string Location { get; set; }
        public string State { get; set; }
        public string Demographics { get; set; }
        public string FamilyMonthlyIncome { get; set; }
        public string FullAddress { get; set; }
        public string PermanentAddress { get; set; }
        public string ParentMobileNumber { get; set; }

    }
}
