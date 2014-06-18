using ExcelReader.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReader
{
    public class FullStudent
    {
        public StudentProfile Profile { get; set; }
        public Placement Placement { get; set; }
        public List<Legacy_SubjectScore> Scores { get; set; }
    }
}
