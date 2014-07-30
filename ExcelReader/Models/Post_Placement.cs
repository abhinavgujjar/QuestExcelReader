using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.Models
{
    public class Post_Placement
    {
        public int Id { get; set; }
     
        public string StudentUid { get; set; }
        public string ContinueJob { get; set; }
        public string Company { get; set; }
        public string Position { get; set; }
        public string UpdatedContact { get; set; }
        public string Salary { get; set; }
    }
}
