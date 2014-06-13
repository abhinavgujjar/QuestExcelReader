using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.Models
{
    public class Legacy_SubjectScore
    {
        public int Id { get; set; }
        public int StudentId { get; set; }
        public decimal Total { get; set; }
        public decimal Score { get; set; }
        public string Subject { get; set; }
        public string Category { get; set; }
    }
}
