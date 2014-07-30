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
     
        public string StudentUid { get; set; }    
        public string EmploymentStatus { get; set; }
        public string Company { get; set; }
        public string Position { get; set; }
        public string NatureOfJob { get; set; }
        public string IfContEduReason { get; set; }
        public string IfDropOutReason { get; set; }
        public string UpdatedContact { get; set; }       
        public string Salary { get; set; }    
        public string Location { get; set; }
    }
}
