using System;
namespace ExcelReader
{
    public interface IImporter
    {
        ImportResult Import();
        ValidationResult Validate();
    }
}
