using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.Models
{
    public class Placement
    {
        public int Id { get; set; }
        public int StudentId { get; set; }
        public string StudentUid { get; set; }

        public string CourseCompletionStatus { get; set; }
        public string EmploymentStatus { get; set; }
        public string Company { get; set; }
        public string Position { get; set; }
        public string OfferLetter { get; set; }
        public string Salary { get; set; }
        public string Comments { get; set; }
        public string Location { get; set; }
    }
}
