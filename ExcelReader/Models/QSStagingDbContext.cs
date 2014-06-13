using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.Models
{
    public class QSStagingDbContext: DbContext
    {
        public DbSet<StudentProfile> StudentProfiles { get; set; }
        public DbSet<SubjectScore> SubjectScores { get; set; }
        public DbSet<Subject> Subjects { get; set; }
        public DbSet<Placement> Placements { get; set; }
        public DbSet<Legacy_SubjectScore> LegacySubjectScores { get; set; }

    }
}
