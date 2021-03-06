﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReader
{
    public class ImportResult
    {
        public int NumberOfRecords { get; set; }
        public string Message { get; set; }
        public bool Failed { get; set; }
        public List<FullStudent> ImportStudents{ get; set; }
        public List<FullStudent> ImportPlacements{ get; set; }
        public List<FullStudent> ImportPostPlacements { get; set; }
    }
}
