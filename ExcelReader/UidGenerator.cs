using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelReader
{
    public class UidGenerator
    {
        public static string GenerateUid(string studentName, string centerName, string centerLocation, int batchNumber)
        {
            if ((studentName == null || centerName == null || centerLocation == null)
                ||
                (studentName.Length <= 3 || centerName.Length <= 3 || centerLocation.Length <= 3)
                )
            {
                throw new ArgumentException();
            }

            var uid =
                clean(centerName) +
                clean(centerLocation) +
                batchNumber +
                clean(studentName);

            return uid;
        }

        private static string clean(string input)
        {
            Regex rgx = new Regex("[^a-zA-Z]");
            input = rgx.Replace(input , "");

            return input.ToLower().Substring(0, 2);
        }

    }

    

}
