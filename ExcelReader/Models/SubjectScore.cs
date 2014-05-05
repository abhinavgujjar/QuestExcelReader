using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.Models
{
    public class SubjectScore
    {
        public int Id { get; set; }
        public string StudentUID { get; set; }
        public int Total { get; set; }
        public int Score { get; set; }
        public int SubjectId { get; set; }
        public string Subject { get; set; }
    }
}
