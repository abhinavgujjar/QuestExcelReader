using ExcelReader;
using ExcelReader.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Security;
using System.ComponentModel;
using System.Collections.ObjectModel;
using Microsoft.Data.Edm;
using System.Data;
using System.Data.Entity;
using System.Drawing; 
using System.Runtime.Serialization;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Gimli
{
    public partial class Form2 : Form
    {
        QSStagingDbContext db = new QSStagingDbContext();
        static object useDefault = Type.Missing;
        public Form2()
        {
            InitializeComponent();
            /*comboBox1.Items.Add("Gender");
            comboBox1.Items.Add("Age");
            comboBox1.Items.Add("Economic Status");
            comboBox1.Items.Add("Education Status");
            comboBox1.Items.Add("Employment Status");
            comboBox1.Items.Add("Average Salary");*/

            var ReportOptions = new BindingList<KeyValuePair<int, string>>();
            ReportOptions.Add(new KeyValuePair<int, string>(0, "Gender"));
            ReportOptions.Add(new KeyValuePair<int, string>(1, "Age"));
            ReportOptions.Add(new KeyValuePair<int, string>(2, "Economic Status"));
            ReportOptions.Add(new KeyValuePair<int, string>(3, "Education Status"));
            ReportOptions.Add(new KeyValuePair<int, string>(4, "Employment Status"));
            ReportOptions.Add(new KeyValuePair<int, string>(5, "Average Salary"));

            comboBox1.DataSource = ReportOptions;
            comboBox1.ValueMember = "Key";
            comboBox1.DisplayMember = "Value";
            comboBox1.SelectedIndex = 0;
        }

        static void SetCellValue(Worksheet targetSheet, string cell,
    object value)
        {
            targetSheet.get_Range(cell, useDefault).set_Value(
                XlRangeValueDataType.xlRangeValueDefault, value);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {


        }


        private void generateReport_click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                // Declare variables that hold references to excel objects.
                ApplicationClass excelApplication = null;
                Workbook excelWorkBook = null;
                Worksheet targetSheet = null;
                PivotTable pivotTable = null;
                Range pivotData = null;
                Range pivotDestination = null;
                PivotField CENTRE = null;
                PivotField MALE = null;
                PivotField FEMALE = null;
                PivotField TOTAL_STUDENTS = null;


                // Declare helper variables.
                string workBookName = @"F:\test\pivottablesample.xlsx";
                string pivotTableName = @"REPORT BASED ON GENDER";
                string workSheetName = @"REPORT";


                try
                {
                    // Create an instance of Excel.
                    excelApplication = new ApplicationClass();

                    //Create a workbook and add a worksheet.
                    excelWorkBook = excelApplication.Workbooks.Add(
                        XlWBATemplate.xlWBATWorksheet);
                    targetSheet = (Worksheet)(excelWorkBook.Worksheets[1]);
                    targetSheet.Name = workSheetName;

                    // Add Data to the Worksheet.
                    SetCellValue(targetSheet, "A1", "BASED ON GENDER");
                    SetCellValue(targetSheet, "A2", "CENTRE NAME");
                    SetCellValue(targetSheet, "B2", "MALE");
                    SetCellValue(targetSheet, "C2", "FEMALE");
                    SetCellValue(targetSheet, "D2", "TOTAL");




                    var OrgCount = (from c in db.StudentProfiles
                                    group c.TrainingCenter by c.TrainingCenter into uniqueIds
                                    select uniqueIds.FirstOrDefault()).Count();





                    var OrgName = (from c in db.StudentProfiles
                                   select c.TrainingCenter).Distinct();


                    int i = 3;
                    foreach (var Org in OrgName)
                    {

                        SetCellValue(targetSheet, 'A' + (i).ToString(), Org);

                        var m = 0;
                        var f = 0;

                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount = (from b in context.StudentProfiles
                                             where b.TrainingCenter.Equals(Org) && b.Gender.Equals("Male")
                                             select b).Count();
                            m = MaleCount;
                            SetCellValue(targetSheet, 'B' + (i).ToString(), MaleCount);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount = (from b in context.StudentProfiles
                                               where b.TrainingCenter.Equals(Org) && b.Gender.Equals("Female")
                                               select b).Count();
                            f = FemaleCount;
                            SetCellValue(targetSheet, 'C' + (i).ToString(), FemaleCount);
                        }
                        var t = m + f;

                        SetCellValue(targetSheet, 'D' + (i).ToString(), t);

                        i++;
                    }



                    // Select a range of data for the Pivot Table.
                    Excel.Range last = targetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    pivotData = targetSheet.get_Range("A2", last);

                    if (excelApplication.Application.Sheets.Count < 2)
                    {
                        targetSheet = (Excel.Worksheet)excelWorkBook.Worksheets.Add();
                    }
                    else
                    {
                        targetSheet = (Excel.Worksheet)(excelWorkBook.Worksheets[2]);
                    }

                    targetSheet.Name = "REPORTS BASED ON GENDER";

                    // Select location of the Pivot Table.
                    pivotDestination = targetSheet.get_Range("A1", useDefault);

                    // Add a Pivot Table to the Worksheet.
                    excelWorkBook.PivotTableWizard(
                        XlPivotTableSourceType.xlDatabase,
                        pivotData,
                        pivotDestination,
                        pivotTableName,
                        true,
                        true,
                        true,
                        true,
                        useDefault,
                        useDefault,
                        false,
                        false,
                        XlOrder.xlDownThenOver,
                        0,
                        useDefault,
                        useDefault
                        );

                    // Set variables for used to manipulate the Pivot Table.
                    pivotTable =
                        (PivotTable)targetSheet.PivotTables(pivotTableName);
                    CENTRE = ((PivotField)pivotTable.PivotFields("CENTRE NAME"));
                    MALE = ((PivotField)pivotTable.PivotFields("MALE"));
                    FEMALE = ((PivotField)pivotTable.PivotFields("FEMALE"));
                    TOTAL_STUDENTS = ((PivotField)pivotTable.PivotFields("TOTAL"));

                    // Format the Pivot Table.
                    pivotTable.Format(XlPivotFormatType.xlReport2);
                    pivotTable.InGridDropZones = false;

                    // Set Centre as a Row Field.
                    CENTRE.Orientation =
                        XlPivotFieldOrientation.xlRowField;

                    // Set Value Field.
                    MALE.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    FEMALE.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL_STUDENTS.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL_STUDENTS.Function = XlConsolidationFunction.xlSum;

                    // Save the Workbook.
                    excelWorkBook.SaveAs(workBookName, useDefault, useDefault,
                        useDefault, useDefault, useDefault,
                        XlSaveAsAccessMode.xlNoChange, useDefault, useDefault,
                        useDefault, useDefault, useDefault);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    // Release the references to the Excel objects.
                    CENTRE = null;
                    TOTAL_STUDENTS = null;
                    MALE = null;
                    FEMALE = null;
                    pivotDestination = null;
                    pivotData = null;
                    pivotTable = null;
                    targetSheet = null;

                    excelApplication.Application.DisplayAlerts = false;
                    ((Microsoft.Office.Interop.Excel.Worksheet)excelWorkBook.Worksheets[2]).Delete();
                    excelApplication.Application.DisplayAlerts = true;

                    // Release the Workbook object.
                    if (excelWorkBook != null)
                        excelWorkBook = null;

                    // Release the ApplicationClass object.
                    if (excelApplication != null)
                    {
                        excelApplication.Quit();
                        excelApplication = null;
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    MessageBox.Show("Report Generated");
                }
            }


            if (comboBox1.SelectedIndex == 1)
            {

                // Declare variables that hold references to excel objects.
                ApplicationClass excelApplication = null;
                Workbook excelWorkBook = null;
                Worksheet targetSheet = null;
                PivotTable pivotTable = null;
                Range pivotData = null;
                Range pivotDestination = null;
                PivotField CENTRE = null;
                PivotField MALE1 = null;
                PivotField FEMALE1 = null;
                PivotField MALE2 = null;
                PivotField FEMALE2 = null;
                PivotField MALE3 = null;
                PivotField FEMALE3 = null;
                PivotField MALE4 = null;
                PivotField FEMALE4 = null;
                PivotField MALE5 = null;
                PivotField FEMALE5 = null;
                PivotField MALE6 = null;
                PivotField FEMALE6 = null;
                PivotField TOTAL1 = null;
                PivotField TOTAL2 = null;
                PivotField TOTAL3 = null;
                PivotField TOTAL4 = null;
                PivotField TOTAL5 = null;
                PivotField TOTAL6 = null;
                PivotField TOTAL_STUDENTS = null;


                // Declare helper variables.
                string workBookName = @"F:\test\pivottablesample.xlsx";
                string pivotTableName = @"BASED ON AGE";
                string workSheetName = @"REPORT";


                try
                {
                    // Create an instance of Excel.
                    excelApplication = new ApplicationClass();

                    //Create a workbook and add a worksheet.
                    excelWorkBook = excelApplication.Workbooks.Add(
                        XlWBATemplate.xlWBATWorksheet);
                    targetSheet = (Worksheet)(excelWorkBook.Worksheets[1]);
                    targetSheet.Name = workSheetName;

                    // Add Data to the Worksheet.
                    SetCellValue(targetSheet, "A1", "BASED ON AGE");
                    SetCellValue(targetSheet, "A3", "CENTRE NAME");
                    SetCellValue(targetSheet, "B2", "UPTO 18");
                    SetCellValue(targetSheet, "C2", "19-25");
                    SetCellValue(targetSheet, "D2", "26-30");
                    SetCellValue(targetSheet, "E2", "31-35");
                    SetCellValue(targetSheet, "F2", "36-40");
                    SetCellValue(targetSheet, "G2", "ABOVE 40");
                    SetCellValue(targetSheet, "B3", "MALE");
                    SetCellValue(targetSheet, "C3", "FEMALE");
                    SetCellValue(targetSheet, "D3", "TOTAL");
                    SetCellValue(targetSheet, "E3", "MALE ");
                    SetCellValue(targetSheet, "F3", "FEMALE ");
                    SetCellValue(targetSheet, "G3", "TOTAL ");
                    SetCellValue(targetSheet, "H3", " MALE");
                    SetCellValue(targetSheet, "I3", " FEMALE");
                    SetCellValue(targetSheet, "J3", " TOTAL");
                    SetCellValue(targetSheet, "K3", "MALE  ");
                    SetCellValue(targetSheet, "L3", "FEMALE  ");
                    SetCellValue(targetSheet, "M3", "TOTAL  ");
                    SetCellValue(targetSheet, "N3", "  MALE");
                    SetCellValue(targetSheet, "O3", "  FEMALE");
                    SetCellValue(targetSheet, "P3", "  TOTAL");
                    SetCellValue(targetSheet, "Q3", " MALE ");
                    SetCellValue(targetSheet, "R3", " FEMALE ");
                    SetCellValue(targetSheet, "S3", " TOTAL ");
                    SetCellValue(targetSheet, "T3", "COMPLETE TOTAL");


                    var OrgCount = (from c in db.StudentProfiles
                                    group c.TrainingCenter by c.TrainingCenter into uniqueIds
                                    select uniqueIds.FirstOrDefault()).Count();





                    var OrgName = (from c in db.StudentProfiles
                                   select c.TrainingCenter).Distinct();


                    int i = 4;
                    foreach (var Org in OrgName)
                    {

                        SetCellValue(targetSheet, 'A' + (i).ToString(), Org);

                        var m1 = 0;
                        var m2 = 0;
                        var m3 = 0;
                        var m4 = 0;
                        var m5 = 0;
                        var m6 = 0;
                        var f1 = 0;
                        var f2 = 0;
                        var f3 = 0;
                        var f4 = 0;
                        var f5 = 0;
                        var f6 = 0;

                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount1 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.Age<=18
                                              select b).Count();
                            m1 = MaleCount1;
                            SetCellValue(targetSheet, 'B' + (i).ToString(), MaleCount1);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount1 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.Age <= 18
                                                select b).Count();
                            f1 = FemaleCount1;
                            SetCellValue(targetSheet, 'C' + (i).ToString(), FemaleCount1);
                        }
                        var t1 = m1 + f1;

                        SetCellValue(targetSheet, 'D' + (i).ToString(), t1);


                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount2 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.Age > 18 && b.Age <= 25
                                              select b).Count();
                            m2 = MaleCount2;
                            SetCellValue(targetSheet, 'E' + (i).ToString(), MaleCount2);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount1 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.Age > 18 && b.Age <= 25
                                                select b).Count();
                            f2 = FemaleCount1;
                            SetCellValue(targetSheet, 'F' + (i).ToString(), FemaleCount1);
                        }
                        var t2 = m2 + f2;

                        SetCellValue(targetSheet, 'G' + (i).ToString(), t2);


                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount3 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.Age > 25 && b.Age <= 30
                                              select b).Count();
                            m3 = MaleCount3;
                            SetCellValue(targetSheet, 'H' + (i).ToString(), MaleCount3);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount3 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.Age > 25 && b.Age <= 30
                                                select b).Count();
                            f3 = FemaleCount3;
                            SetCellValue(targetSheet, 'I' + (i).ToString(), FemaleCount3);
                        }
                        var t3 = m3 + f3;

                        SetCellValue(targetSheet, 'J' + (i).ToString(), t3);

                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount4 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.Age > 30 && b.Age <= 35
                                              select b).Count();
                            m4 = MaleCount4;
                            SetCellValue(targetSheet, 'K' + (i).ToString(), MaleCount4);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount4 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.Age > 30 && b.Age <= 35
                                                select b).Count();
                            f4 = FemaleCount4;
                            SetCellValue(targetSheet, 'L' + (i).ToString(), FemaleCount4);
                        }
                        var t4 = m4 + f4;

                        SetCellValue(targetSheet, 'M' + (i).ToString(), t4);

                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount5 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.Age > 35 && b.Age <= 40
                                              select b).Count();
                            m5 = MaleCount5;
                            SetCellValue(targetSheet, 'N' + (i).ToString(), MaleCount5);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount5 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.Age > 35 && b.Age <= 40
                                                select b).Count();
                            f5 = FemaleCount5;
                            SetCellValue(targetSheet, 'O' + (i).ToString(), FemaleCount5);
                        }
                        var t5 = m5 + f5;

                        SetCellValue(targetSheet, 'P' + (i).ToString(), t5);

                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount6 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.Age >40
                                              select b).Count();
                            m6 = MaleCount6;
                            SetCellValue(targetSheet, 'Q' + (i).ToString(), MaleCount6);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount6 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.Age > 40
                                                select b).Count();
                            f6 = FemaleCount6;
                            SetCellValue(targetSheet, 'R' + (i).ToString(), FemaleCount6);
                        }
                        var t6 = m6 + f6;

                        SetCellValue(targetSheet, 'S' + (i).ToString(), t6);

                        SetCellValue(targetSheet, 'T' + (i).ToString(), t1 + t2 + t3 + t4 + t5 + t6);

                        i++;
                    }



                    // Select a range of data for the Pivot Table.
                    Excel.Range last = targetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    pivotData = targetSheet.get_Range("A3", last);
                    //pivotData.Merge(Type.Missing);

                    if (excelApplication.Application.Sheets.Count < 2)
                    {
                        targetSheet = (Excel.Worksheet)excelWorkBook.Worksheets.Add();
                    }
                    else
                    {
                        targetSheet = (Excel.Worksheet)(excelWorkBook.Worksheets[2]);
                    }

                    SetCellValue(targetSheet, "A1", "CENTRE NAME");
                    SetCellValue(targetSheet, "B1", "UPTO 18");
                    SetCellValue(targetSheet, "E1", "19-25");
                    SetCellValue(targetSheet, "H1", "26-30");
                    SetCellValue(targetSheet, "K1", "31-35");
                    SetCellValue(targetSheet, "N1", "36-40");
                    SetCellValue(targetSheet, "Q1", "ABOVE 40");
                    SetCellValue(targetSheet, "T1", "COMPLETE TOTAL");

                    Excel.Range Merge1 = targetSheet.get_Range("B1", "D1");
                    Merge1.Merge();

                    Excel.Range Merge2 = targetSheet.get_Range("E1", "G1");
                    Merge2.Merge();

                    Excel.Range Merge3 = targetSheet.get_Range("H1", "J1");
                    Merge3.Merge();

                    Excel.Range Merge4 = targetSheet.get_Range("K1", "M1");
                    Merge4.Merge();

                    Excel.Range Merge5 = targetSheet.get_Range("N1", "P1");
                    Merge5.Merge();

                    Excel.Range Merge6 = targetSheet.get_Range("Q1", "S1");
                    Merge6.Merge();

                    Excel.Range Heading = targetSheet.get_Range("A1", "T1");
                    Heading.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    Heading.Borders.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    Heading.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;        

                    targetSheet.Name = "BASED ON AGE";

                    // Select location of the Pivot Table.
                    pivotDestination = targetSheet.get_Range("A2", useDefault);
                    pivotDestination.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    // Add a Pivot Table to the Worksheet.
                    excelWorkBook.PivotTableWizard(
                        XlPivotTableSourceType.xlDatabase,
                        pivotData,
                        pivotDestination,
                        pivotTableName,
                        true,
                        true,
                        true,
                        true,
                        useDefault,
                        useDefault,
                        false,
                        false,
                        XlOrder.xlDownThenOver,
                        0,
                        useDefault,
                        useDefault
                        );

                    // Set variables for used to manipulate the Pivot Table.
                    pivotTable =
                        (PivotTable)targetSheet.PivotTables(pivotTableName);
                    CENTRE = ((PivotField)pivotTable.PivotFields("CENTRE NAME"));
                    MALE1 = (PivotField)pivotTable.PivotFields(2);
                    FEMALE1 = ((PivotField)pivotTable.PivotFields(3));
                    TOTAL1 = ((PivotField)pivotTable.PivotFields(4));
                    MALE2 = (PivotField)pivotTable.PivotFields(5);
                    FEMALE2 = ((PivotField)pivotTable.PivotFields(6));
                    TOTAL2 = ((PivotField)pivotTable.PivotFields(7));
                    MALE3 = (PivotField)pivotTable.PivotFields(8);
                    FEMALE3 = ((PivotField)pivotTable.PivotFields(9));
                    TOTAL3 = ((PivotField)pivotTable.PivotFields(10));
                    MALE4 = (PivotField)pivotTable.PivotFields(11);
                    FEMALE4 = ((PivotField)pivotTable.PivotFields(12));
                    TOTAL4 = ((PivotField)pivotTable.PivotFields(13));
                    MALE5 = (PivotField)pivotTable.PivotFields(14);
                    FEMALE5 = ((PivotField)pivotTable.PivotFields(15));
                    TOTAL5 = ((PivotField)pivotTable.PivotFields(16));
                    MALE6 = (PivotField)pivotTable.PivotFields(17);
                    FEMALE6 = ((PivotField)pivotTable.PivotFields(18));
                    TOTAL6 = ((PivotField)pivotTable.PivotFields(19));
                    TOTAL_STUDENTS = ((PivotField)pivotTable.PivotFields("COMPLETE TOTAL"));

                    // Format the Pivot Table.
                    pivotTable.Format(XlPivotFormatType.xlReport2);
                    pivotTable.InGridDropZones = false;

                    // Set Centre as a Row Field.
                    CENTRE.Orientation =
                        XlPivotFieldOrientation.xlRowField;

                    // Set Value Field.
                    MALE1.Orientation =
                         XlPivotFieldOrientation.xlDataField;
                    FEMALE1.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL1.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    MALE2.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    FEMALE2.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL2.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    MALE3.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    FEMALE3.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL3.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    MALE4.Orientation =
                         XlPivotFieldOrientation.xlDataField;
                    FEMALE4.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL4.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    MALE5.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    FEMALE5.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL5.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    MALE6.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    FEMALE6.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL6.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL_STUDENTS.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL_STUDENTS.Function = XlConsolidationFunction.xlSum;

                    // Save the Workbook.
                    excelWorkBook.SaveAs(workBookName, useDefault, useDefault,
                        useDefault, useDefault, useDefault,
                        XlSaveAsAccessMode.xlNoChange, useDefault, useDefault,
                        useDefault, useDefault, useDefault);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
                finally
                {
                    // Release the references to the Excel objects.
                    CENTRE = null;
                    TOTAL_STUDENTS = null;
                    MALE1 = null;
                    FEMALE1 = null;
                    MALE2 = null;
                    FEMALE2 = null;
                    MALE3 = null;
                    FEMALE3 = null;
                    MALE4 = null;
                    FEMALE4 = null;
                    MALE5 = null;
                    FEMALE5 = null;
                    MALE6 = null;
                    FEMALE6 = null;
                    TOTAL1 = null;
                    TOTAL2 = null;
                    TOTAL3 = null;
                    TOTAL4 = null;
                    TOTAL5 = null;
                    TOTAL6 = null;
                    pivotDestination = null;
                    pivotData = null;
                    pivotTable = null;
                    targetSheet = null;

                    excelApplication.Application.DisplayAlerts = false;
                    ((Microsoft.Office.Interop.Excel.Worksheet)excelWorkBook.Worksheets[2]).Delete();
                    excelApplication.Application.DisplayAlerts = true;

                    // Release the Workbook object.
                    if (excelWorkBook != null)
                        excelWorkBook = null;

                    // Release the ApplicationClass object.
                    if (excelApplication != null)
                    {
                        excelApplication.Quit();
                        excelApplication = null;
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    MessageBox.Show("Report Generated");
                }
            }


            if (comboBox1.SelectedIndex == 2)
            {
                // Declare variables that hold references to excel objects.
                ApplicationClass excelApplication = null;
                Workbook excelWorkBook = null;
                Worksheet targetSheet = null;
                PivotTable pivotTable = null;
                Range pivotData = null;
                Range pivotDestination = null;
                PivotField CENTRE = null;
                PivotField MALE1 = null;
                PivotField FEMALE1 = null;
                PivotField MALE2 = null;
                PivotField FEMALE2 = null;
                PivotField MALE3 = null;
                PivotField FEMALE3 = null;
                PivotField TOTAL1 = null;
                PivotField TOTAL2 = null;
                PivotField TOTAL3 = null;
                PivotField TOTAL_STUDENTS = null;


                // Declare helper variables.
                string workBookName = @"F:\test\pivottablesample.xlsx";
                string pivotTableName = @"BASED ON ECONOMIC STATUS";
                string workSheetName = @"REPORT";


                try
                {
                    // Create an instance of Excel.
                    excelApplication = new ApplicationClass();

                    //Create a workbook and add a worksheet.
                    excelWorkBook = excelApplication.Workbooks.Add(
                        XlWBATemplate.xlWBATWorksheet);
                    targetSheet = (Worksheet)(excelWorkBook.Worksheets[1]);
                    targetSheet.Name = workSheetName;

                    // Add Data to the Worksheet.
                    SetCellValue(targetSheet, "A1", "BASED ON ECONOMIC STATUS");
                    SetCellValue(targetSheet, "A3", "CENTRE NAME");
                    SetCellValue(targetSheet, "B2", "LESS THAN 5000");
                    SetCellValue(targetSheet, "C2", "5000 TO 10000");
                    SetCellValue(targetSheet, "D2", "MORE THAN 10000");
                    SetCellValue(targetSheet, "B3", "MALE");
                    SetCellValue(targetSheet, "C3", "FEMALE");
                    SetCellValue(targetSheet, "D3", "TOTAL");
                    SetCellValue(targetSheet, "E3", "MALE ");
                    SetCellValue(targetSheet, "F3", "FEMALE ");
                    SetCellValue(targetSheet, "G3", "TOTAL ");
                    SetCellValue(targetSheet, "H3", " MALE");
                    SetCellValue(targetSheet, "I3", " FEMALE");
                    SetCellValue(targetSheet, "J3", " TOTAL");
                    SetCellValue(targetSheet, "K3", "COMPLETE TOTAL");


                    var OrgCount = (from c in db.StudentProfiles
                                    group c.TrainingCenter by c.TrainingCenter into uniqueIds
                                    select uniqueIds.FirstOrDefault()).Count();





                    var OrgName = (from c in db.StudentProfiles
                                   select c.TrainingCenter).Distinct();


                    int i = 4;
                    foreach (var Org in OrgName)
                    {

                        SetCellValue(targetSheet, 'A' + (i).ToString(), Org);

                        var m1 = 0;
                        var m2 = 0;
                        var m3 = 0;
                        var f1 = 0;
                        var f2 = 0;
                        var f3 = 0;

                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount1 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.FamilyMonthlyIncome.Equals("Less than Rs.5000")
                                              select b).Count();
                            m1 = MaleCount1;
                            SetCellValue(targetSheet, 'B' + (i).ToString(), MaleCount1);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount1 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.FamilyMonthlyIncome.Equals("Less than Rs.5000")
                                                select b).Count();
                            f1 = FemaleCount1;
                            SetCellValue(targetSheet, 'C' + (i).ToString(), FemaleCount1);
                        }
                        var t1 = m1 + f1;

                        SetCellValue(targetSheet, 'D' + (i).ToString(), t1);


                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount2 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.FamilyMonthlyIncome.Equals("Rs.5000 - Rs.10000")
                                              select b).Count();
                            m2 = MaleCount2;
                            SetCellValue(targetSheet, 'E' + (i).ToString(), MaleCount2);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount1 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.FamilyMonthlyIncome.Equals("Rs.5000 - Rs.10000")
                                                select b).Count();
                            f2 = FemaleCount1;
                            SetCellValue(targetSheet, 'F' + (i).ToString(), FemaleCount1);
                        }
                        var t2 = m2 + f2;

                        SetCellValue(targetSheet, 'G' + (i).ToString(), t2);


                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount3 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.FamilyMonthlyIncome.Equals("More than Rs.10000")
                                              select b).Count();
                            m3 = MaleCount3;
                            SetCellValue(targetSheet, 'H' + (i).ToString(), MaleCount3);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount3 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.FamilyMonthlyIncome.Equals("More than Rs.10000")
                                                select b).Count();
                            f3 = FemaleCount3;
                            SetCellValue(targetSheet, 'I' + (i).ToString(), FemaleCount3);
                        }
                        var t3 = m3 + f3;

                        SetCellValue(targetSheet, 'J' + (i).ToString(), t3);

                        SetCellValue(targetSheet, 'K' + (i).ToString(), t1 + t2 + t3);

                        i++;
                    }



                    // Select a range of data for the Pivot Table.
                    Excel.Range last = targetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    pivotData = targetSheet.get_Range("A3", last);
                    //pivotData.Merge(Type.Missing);

                    if (excelApplication.Application.Sheets.Count < 2)
                    {
                        targetSheet = (Excel.Worksheet)excelWorkBook.Worksheets.Add();
                    }
                    else
                    {
                        targetSheet = (Excel.Worksheet)(excelWorkBook.Worksheets[2]);
                    }

                    SetCellValue(targetSheet, "A1", "CENTRE NAME");
                    SetCellValue(targetSheet, "B1", "LESS THAN 5000");
                    SetCellValue(targetSheet, "E1", "5000 TO 10000");
                    SetCellValue(targetSheet, "H1", "MORE THAN 10000");
                    SetCellValue(targetSheet, "K1", "COMPLETE TOTAL");

                    Excel.Range Merge1 = targetSheet.get_Range("B1", "D1");
                    Merge1.Merge();

                    Excel.Range Merge2 = targetSheet.get_Range("E1", "G1");
                    Merge2.Merge();

                    Excel.Range Merge3 = targetSheet.get_Range("H1", "J1");
                    Merge3.Merge();

                    Excel.Range Heading = targetSheet.get_Range("A1", "K1");
                    Heading.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    Heading.Borders.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    Heading.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    targetSheet.Name = "BASED ON ECONOMIC STATUS";

                    // Select location of the Pivot Table.
                    pivotDestination = targetSheet.get_Range("A2", useDefault);
                    pivotDestination.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    // Add a Pivot Table to the Worksheet.
                    excelWorkBook.PivotTableWizard(
                        XlPivotTableSourceType.xlDatabase,
                        pivotData,
                        pivotDestination,
                        pivotTableName,
                        true,
                        true,
                        true,
                        true,
                        useDefault,
                        useDefault,
                        false,
                        false,
                        XlOrder.xlDownThenOver,
                        0,
                        useDefault,
                        useDefault
                        );

                    // Set variables for used to manipulate the Pivot Table.
                    pivotTable =
                        (PivotTable)targetSheet.PivotTables(pivotTableName);
                    CENTRE = ((PivotField)pivotTable.PivotFields("CENTRE NAME"));
                    MALE1 = (PivotField)pivotTable.PivotFields(2);
                    FEMALE1 = ((PivotField)pivotTable.PivotFields(3));
                    TOTAL1 = ((PivotField)pivotTable.PivotFields(4));
                    MALE2 = (PivotField)pivotTable.PivotFields(5);
                    FEMALE2 = ((PivotField)pivotTable.PivotFields(6));
                    TOTAL2 = ((PivotField)pivotTable.PivotFields(7));
                    MALE3 = (PivotField)pivotTable.PivotFields(8);
                    FEMALE3 = ((PivotField)pivotTable.PivotFields(9));
                    TOTAL3 = ((PivotField)pivotTable.PivotFields(10));
                    TOTAL_STUDENTS = ((PivotField)pivotTable.PivotFields("COMPLETE TOTAL"));

                    // Format the Pivot Table.
                    pivotTable.Format(XlPivotFormatType.xlReport2);
                    pivotTable.InGridDropZones = false;

                    // Set Centre as a Row Field.
                    CENTRE.Orientation =
                        XlPivotFieldOrientation.xlRowField;

                    // Set Value Field.
                    MALE1.Orientation =
                         XlPivotFieldOrientation.xlDataField;
                    FEMALE1.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL1.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    MALE2.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    FEMALE2.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL2.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    MALE3.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    FEMALE3.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL3.Orientation =
                   XlPivotFieldOrientation.xlDataField;
                    TOTAL_STUDENTS.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL_STUDENTS.Function = XlConsolidationFunction.xlSum;

                    // Save the Workbook.
                    excelWorkBook.SaveAs(workBookName, useDefault, useDefault,
                        useDefault, useDefault, useDefault,
                        XlSaveAsAccessMode.xlNoChange, useDefault, useDefault,
                        useDefault, useDefault, useDefault);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
                finally
                {
                    // Release the references to the Excel objects.
                    CENTRE = null;
                    TOTAL_STUDENTS = null;
                    MALE1 = null;
                    FEMALE1 = null;
                    MALE2 = null;
                    FEMALE2 = null;
                    MALE3 = null;
                    FEMALE3 = null;
                    TOTAL1 = null;
                    TOTAL2 = null;
                    TOTAL3 = null;
                    pivotDestination = null;
                    pivotData = null;
                    pivotTable = null;
                    targetSheet = null;

                    excelApplication.Application.DisplayAlerts = false;
                    ((Microsoft.Office.Interop.Excel.Worksheet)excelWorkBook.Worksheets[2]).Delete();
                    excelApplication.Application.DisplayAlerts = true;

                    // Release the Workbook object.
                    if (excelWorkBook != null)
                        excelWorkBook = null;

                    // Release the ApplicationClass object.
                    if (excelApplication != null)
                    {
                        excelApplication.Quit();
                        excelApplication = null;
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    MessageBox.Show("Report Generated");
                }
            }

            if (comboBox1.SelectedIndex == 3)
            {
                // Declare variables that hold references to excel objects.
                ApplicationClass excelApplication = null;
                Workbook excelWorkBook = null;
                Worksheet targetSheet = null;
                PivotTable pivotTable = null;
                Range pivotData = null;
                Range pivotDestination = null;
                PivotField CENTRE = null;
                PivotField MALE1 = null;
                PivotField FEMALE1 = null;
                PivotField MALE2 = null;
                PivotField FEMALE2 = null;
                PivotField MALE3 = null;
                PivotField FEMALE3 = null;
                PivotField MALE4 = null;
                PivotField FEMALE4 = null;
                PivotField MALE5 = null;
                PivotField FEMALE5 = null;
                PivotField MALE6 = null;
                PivotField FEMALE6 = null;
                PivotField TOTAL1 = null;
                PivotField TOTAL2 = null;
                PivotField TOTAL3 = null;
                PivotField TOTAL4 = null;
                PivotField TOTAL5 = null;
                PivotField TOTAL6 = null;
                PivotField TOTAL_STUDENTS = null;


                // Declare helper variables.
                string workBookName = @"F:\test\pivottablesample.xlsx";
                string pivotTableName = @"BASED ON EDUCATIONAL STATUS";
                string workSheetName = @"REPORT";


                try
                {
                    // Create an instance of Excel.
                    excelApplication = new ApplicationClass();

                    //Create a workbook and add a worksheet.
                    excelWorkBook = excelApplication.Workbooks.Add(
                        XlWBATemplate.xlWBATWorksheet);
                    targetSheet = (Worksheet)(excelWorkBook.Worksheets[1]);
                    targetSheet.Name = workSheetName;

                    // Add Data to the Worksheet.
                    SetCellValue(targetSheet, "A1", "BASED ON ECONOMIC STATUS");
                    SetCellValue(targetSheet, "A3", "CENTRE NAME");
                    SetCellValue(targetSheet, "B2", "BELOW 10TH");
                    SetCellValue(targetSheet, "C2", "10TH PASS");
                    SetCellValue(targetSheet, "D2", "12TH PASS");
                    SetCellValue(targetSheet, "E2", "GRADUATE");
                    SetCellValue(targetSheet, "F2", "POST GRADUATE");
                    SetCellValue(targetSheet, "G2", "ITI/DIPLOMA");
                    SetCellValue(targetSheet, "B3", "MALE");
                    SetCellValue(targetSheet, "C3", "FEMALE");
                    SetCellValue(targetSheet, "D3", "TOTAL");
                    SetCellValue(targetSheet, "E3", "MALE ");
                    SetCellValue(targetSheet, "F3", "FEMALE ");
                    SetCellValue(targetSheet, "G3", "TOTAL ");
                    SetCellValue(targetSheet, "H3", " MALE");
                    SetCellValue(targetSheet, "I3", " FEMALE");
                    SetCellValue(targetSheet, "J3", " TOTAL");
                    SetCellValue(targetSheet, "K3", "MALE  ");
                    SetCellValue(targetSheet, "L3", "FEMALE  ");
                    SetCellValue(targetSheet, "M3", "TOTAL  ");
                    SetCellValue(targetSheet, "N3", "  MALE");
                    SetCellValue(targetSheet, "O3", "  FEMALE");
                    SetCellValue(targetSheet, "P3", "  TOTAL");
                    SetCellValue(targetSheet, "Q3", " MALE ");
                    SetCellValue(targetSheet, "R3", " FEMALE ");
                    SetCellValue(targetSheet, "S3", " TOTAL ");
                    SetCellValue(targetSheet, "T3", "COMPLETE TOTAL");


                    var OrgCount = (from c in db.StudentProfiles
                                    group c.TrainingCenter by c.TrainingCenter into uniqueIds
                                    select uniqueIds.FirstOrDefault()).Count();





                    var OrgName = (from c in db.StudentProfiles
                                   select c.TrainingCenter).Distinct();


                    int i = 4;
                    foreach (var Org in OrgName)
                    {

                        SetCellValue(targetSheet, 'A' + (i).ToString(), Org);

                        var m1 = 0;
                        var m2 = 0;
                        var m3 = 0;
                        var m4 = 0;
                        var m5 = 0;
                        var m6 = 0;
                        var f1 = 0;
                        var f2 = 0;
                        var f3 = 0;
                        var f4 = 0;
                        var f5 = 0;
                        var f6 = 0;

                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount1 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.Education.Equals("Below 10th")
                                              select b).Count();
                            m1 = MaleCount1;
                            SetCellValue(targetSheet, 'B' + (i).ToString(), MaleCount1);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount1 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.Education.Equals("Below 10th")
                                                select b).Count();
                            f1 = FemaleCount1;
                            SetCellValue(targetSheet, 'C' + (i).ToString(), FemaleCount1);
                        }
                        var t1 = m1 + f1;

                        SetCellValue(targetSheet, 'D' + (i).ToString(), t1);


                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount2 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.Education.Equals("10th Pass")
                                              select b).Count();
                            m2 = MaleCount2;
                            SetCellValue(targetSheet, 'E' + (i).ToString(), MaleCount2);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount1 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.Education.Equals("10th Pass")
                                                select b).Count();
                            f2 = FemaleCount1;
                            SetCellValue(targetSheet, 'F' + (i).ToString(), FemaleCount1);
                        }
                        var t2 = m2 + f2;

                        SetCellValue(targetSheet, 'G' + (i).ToString(), t2);


                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount3 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.Education.Equals("12th Pass")
                                              select b).Count();
                            m3 = MaleCount3;
                            SetCellValue(targetSheet, 'H' + (i).ToString(), MaleCount3);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount3 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.Education.Equals("12th Pass")
                                                select b).Count();
                            f3 = FemaleCount3;
                            SetCellValue(targetSheet, 'I' + (i).ToString(), FemaleCount3);
                        }
                        var t3 = m3 + f3;

                        SetCellValue(targetSheet, 'J' + (i).ToString(), t3);

                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount4 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.Education.Equals("Graduate")
                                              select b).Count();
                            m4 = MaleCount4;
                            SetCellValue(targetSheet, 'K' + (i).ToString(), MaleCount4);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount4 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.Education.Equals("Graduate")
                                                select b).Count();
                            f4 = FemaleCount4;
                            SetCellValue(targetSheet, 'L' + (i).ToString(), FemaleCount4);
                        }
                        var t4 = m4 + f4;

                        SetCellValue(targetSheet, 'M' + (i).ToString(), t4);

                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount5 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.Education.Equals("Post Graduate")
                                              select b).Count();
                            m5 = MaleCount5;
                            SetCellValue(targetSheet, 'N' + (i).ToString(), MaleCount5);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount5 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.Education.Equals("Post Graduate")
                                                select b).Count();
                            f5 = FemaleCount5;
                            SetCellValue(targetSheet, 'O' + (i).ToString(), FemaleCount5);
                        }
                        var t5 = m5 + f5;

                        SetCellValue(targetSheet, 'P' + (i).ToString(), t5);

                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount6 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.Education.Equals("ITI/Diploma")
                                              select b).Count();
                            m6 = MaleCount6;
                            SetCellValue(targetSheet, 'Q' + (i).ToString(), MaleCount6);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount6 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.Education.Equals("ITI/Diploma")
                                                select b).Count();
                            f6 = FemaleCount6;
                            SetCellValue(targetSheet, 'R' + (i).ToString(), FemaleCount6);
                        }
                        var t6 = m6 + f6;

                        SetCellValue(targetSheet, 'S' + (i).ToString(), t6);

                        SetCellValue(targetSheet, 'T' + (i).ToString(), t1 + t2 + t3 + t4 + t5 + t6);

                        i++;
                    }

                    

                    // Select a range of data for the Pivot Table.
                    Excel.Range last = targetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    pivotData = targetSheet.get_Range("A3", last);
                    //pivotData.Merge(Type.Missing);

                    if (excelApplication.Application.Sheets.Count < 2)
                    {
                        targetSheet = (Excel.Worksheet)excelWorkBook.Worksheets.Add();
                    }
                    else
                    {
                        targetSheet = (Excel.Worksheet)(excelWorkBook.Worksheets[2]);
                    }

                    SetCellValue(targetSheet, "A1", "CENTRE NAME");
                    SetCellValue(targetSheet, "B1", "BELOW 10TH");
                    SetCellValue(targetSheet, "E1", "10TH PASS");
                    SetCellValue(targetSheet, "H1", "12TH PASS");
                    SetCellValue(targetSheet, "K1", "GRADUATE");
                    SetCellValue(targetSheet, "N1", "POST GRADUATE");
                    SetCellValue(targetSheet, "Q1", "ITI/DIPLOMA");
                    SetCellValue(targetSheet, "T1", "COMPLETE TOTAL");

                    Excel.Range Merge1 = targetSheet.get_Range("B1", "D1");
                    Merge1.Merge();

                    Excel.Range Merge2 = targetSheet.get_Range("E1", "G1");
                    Merge2.Merge();

                    Excel.Range Merge3 = targetSheet.get_Range("H1", "J1");
                    Merge3.Merge();

                    Excel.Range Merge4 = targetSheet.get_Range("K1", "M1");
                    Merge4.Merge();

                    Excel.Range Merge5 = targetSheet.get_Range("N1", "P1");
                    Merge5.Merge();

                    Excel.Range Merge6 = targetSheet.get_Range("Q1", "S1");
                    Merge6.Merge();

                    Excel.Range Heading = targetSheet.get_Range("A1", "T1");
                    Heading.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    Heading.Borders.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    Heading.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    targetSheet.Name = "BASED ON EDUCATIONAL STATUS";

                    // Select location of the Pivot Table.
                    pivotDestination = targetSheet.get_Range("A2", useDefault);
                    pivotDestination.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    // Add a Pivot Table to the Worksheet.
                    excelWorkBook.PivotTableWizard(
                        XlPivotTableSourceType.xlDatabase,
                        pivotData,
                        pivotDestination,
                        pivotTableName,
                        true,
                        true,
                        true,
                        true,
                        useDefault,
                        useDefault,
                        false,
                        false,
                        XlOrder.xlDownThenOver,
                        0,
                        useDefault,
                        useDefault
                        );

                    // Set variables for used to manipulate the Pivot Table.
                    pivotTable =
                        (PivotTable)targetSheet.PivotTables(pivotTableName);
                    CENTRE = ((PivotField)pivotTable.PivotFields("CENTRE NAME"));
                    MALE1 = (PivotField)pivotTable.PivotFields(2);
                    FEMALE1 = ((PivotField)pivotTable.PivotFields(3));
                    TOTAL1 = ((PivotField)pivotTable.PivotFields(4));
                    MALE2 = (PivotField)pivotTable.PivotFields(5);
                    FEMALE2 = ((PivotField)pivotTable.PivotFields(6));
                    TOTAL2 = ((PivotField)pivotTable.PivotFields(7));
                    MALE3 = (PivotField)pivotTable.PivotFields(8);
                    FEMALE3 = ((PivotField)pivotTable.PivotFields(9));
                    TOTAL3 = ((PivotField)pivotTable.PivotFields(10));
                    MALE4 = (PivotField)pivotTable.PivotFields(11);
                    FEMALE4 = ((PivotField)pivotTable.PivotFields(12));
                    TOTAL4 = ((PivotField)pivotTable.PivotFields(13));
                    MALE5 = (PivotField)pivotTable.PivotFields(14);
                    FEMALE5 = ((PivotField)pivotTable.PivotFields(15));
                    TOTAL5 = ((PivotField)pivotTable.PivotFields(16));
                    MALE6 = (PivotField)pivotTable.PivotFields(17);
                    FEMALE6 = ((PivotField)pivotTable.PivotFields(18));
                    TOTAL6 = ((PivotField)pivotTable.PivotFields(19));
                    TOTAL_STUDENTS = ((PivotField)pivotTable.PivotFields("COMPLETE TOTAL"));

                    // Format the Pivot Table.
                    pivotTable.Format(XlPivotFormatType.xlReport2);
                    pivotTable.InGridDropZones = false;

                    // Set Centre as a Row Field.
                    CENTRE.Orientation =
                        XlPivotFieldOrientation.xlRowField;

                    // Set Value Field.
                    MALE1.Orientation =
                         XlPivotFieldOrientation.xlDataField;
                    FEMALE1.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL1.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    MALE2.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    FEMALE2.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL2.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    MALE3.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    FEMALE3.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL3.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    MALE4.Orientation =
                         XlPivotFieldOrientation.xlDataField;
                    FEMALE4.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL4.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    MALE5.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    FEMALE5.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL5.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    MALE6.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    FEMALE6.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL6.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL_STUDENTS.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL_STUDENTS.Function = XlConsolidationFunction.xlSum;

                    // Save the Workbook.
                    excelWorkBook.SaveAs(workBookName, useDefault, useDefault,
                        useDefault, useDefault, useDefault,
                        XlSaveAsAccessMode.xlNoChange, useDefault, useDefault,
                        useDefault, useDefault, useDefault);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
                finally
                {
                    // Release the references to the Excel objects.
                    CENTRE = null;
                    TOTAL_STUDENTS = null;
                    MALE1 = null;
                    FEMALE1 = null;
                    MALE2 = null;
                    FEMALE2 = null;
                    MALE3 = null;
                    FEMALE3 = null;
                    MALE4 = null;
                    FEMALE4 = null;
                    MALE5 = null;
                    FEMALE5 = null;
                    MALE6 = null;
                    FEMALE6 = null;
                    TOTAL1 = null;
                    TOTAL2 = null;
                    TOTAL3 = null;
                    TOTAL4 = null;
                    TOTAL5 = null;
                    TOTAL6 = null;
                    pivotDestination = null;
                    pivotData = null;
                    pivotTable = null;
                    targetSheet = null;

                    excelApplication.Application.DisplayAlerts = false;
                    ((Microsoft.Office.Interop.Excel.Worksheet)excelWorkBook.Worksheets[2]).Delete();
                    excelApplication.Application.DisplayAlerts = true;

                    // Release the Workbook object.
                    if (excelWorkBook != null)
                        excelWorkBook = null;

                    // Release the ApplicationClass object.
                    if (excelApplication != null)
                    {
                        excelApplication.Quit();
                        excelApplication = null;
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    MessageBox.Show("Report Generated");
                }
            }

            if (comboBox1.SelectedIndex == 4)
            {

                // Declare variables that hold references to excel objects.
                ApplicationClass excelApplication = null;
                Workbook excelWorkBook = null;
                Worksheet targetSheet = null;
                PivotTable pivotTable = null;
                Range pivotData = null;
                Range pivotDestination = null;
                PivotField CENTRE = null;
                PivotField MALE1 = null;
                PivotField FEMALE1 = null;
                PivotField MALE2 = null;
                PivotField FEMALE2 = null;
                PivotField MALE3 = null;
                PivotField FEMALE3 = null;
                PivotField MALE4 = null;
                PivotField FEMALE4 = null;
                PivotField TOTAL1 = null;
                PivotField TOTAL2 = null;
                PivotField TOTAL3 = null;
                PivotField TOTAL4 = null;
                PivotField TOTAL_STUDENTS = null;


                // Declare helper variables.
                string workBookName = @"F:\test\pivottablesample.xlsx";
                string pivotTableName = @"BASED ON EMPLOYMENT STATUS";
                string workSheetName = @"REPORT";


                try
                {
                    // Create an instance of Excel.
                    excelApplication = new ApplicationClass();

                    //Create a workbook and add a worksheet.
                    excelWorkBook = excelApplication.Workbooks.Add(
                        XlWBATemplate.xlWBATWorksheet);
                    targetSheet = (Worksheet)(excelWorkBook.Worksheets[1]);
                    targetSheet.Name = workSheetName;

                    // Add Data to the Worksheet.
                    SetCellValue(targetSheet, "A1", "BASED ON EMPLOYMENT STATUS");
                    SetCellValue(targetSheet, "A3", "CENTRE NAME");
                    SetCellValue(targetSheet, "B2", "LESS THAN 5000");
                    SetCellValue(targetSheet, "C2", "5000 TO 10000");
                    SetCellValue(targetSheet, "D2", "MORE THAN 10000");
                    SetCellValue(targetSheet, "B3", "MALE");
                    SetCellValue(targetSheet, "C3", "FEMALE");
                    SetCellValue(targetSheet, "D3", "TOTAL");
                    SetCellValue(targetSheet, "E3", "MALE ");
                    SetCellValue(targetSheet, "F3", "FEMALE ");
                    SetCellValue(targetSheet, "G3", "TOTAL ");
                    SetCellValue(targetSheet, "H3", " MALE");
                    SetCellValue(targetSheet, "I3", " FEMALE");
                    SetCellValue(targetSheet, "J3", " TOTAL");
                    SetCellValue(targetSheet, "K3", " MALE ");
                    SetCellValue(targetSheet, "L3", " FEMALE ");
                    SetCellValue(targetSheet, "M3", " TOTAL ");
                    SetCellValue(targetSheet, "N3", "COMPLETE TOTAL");


                    var OrgCount = (from c in db.StudentProfiles
                                    group c.TrainingCenter by c.TrainingCenter into uniqueIds
                                    select uniqueIds.FirstOrDefault()).Count();





                    var OrgName = (from c in db.StudentProfiles
                                   select c.TrainingCenter).Distinct();


                    int i = 4;
                    foreach (var Org in OrgName)
                    {

                        SetCellValue(targetSheet, 'A' + (i).ToString(), Org);

                        var m1 = 0;
                        var m2 = 0;
                        var m3 = 0;
                        var m4 = 0;
                        var f1 = 0;
                        var f2 = 0;
                        var f3 = 0;
                        var f4 = 0;

                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount1 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.EmploymentStatus.Equals("Employed")
                                              select b).Count();
                            m1 = MaleCount1;
                            SetCellValue(targetSheet, 'B' + (i).ToString(), MaleCount1);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount1 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.EmploymentStatus.Equals("Employed")
                                                select b).Count();
                            f1 = FemaleCount1;
                            SetCellValue(targetSheet, 'C' + (i).ToString(), FemaleCount1);
                        }
                        var t1 = m1 + f1;

                        SetCellValue(targetSheet, 'D' + (i).ToString(), t1);


                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount2 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.EmploymentStatus.Equals("Not Employed")
                                              select b).Count();
                            m2 = MaleCount2;
                            SetCellValue(targetSheet, 'E' + (i).ToString(), MaleCount2);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount1 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.EmploymentStatus.Equals("Not Employed")
                                                select b).Count();
                            f2 = FemaleCount1;
                            SetCellValue(targetSheet, 'F' + (i).ToString(), FemaleCount1);
                        }
                        var t2 = m2 + f2;

                        SetCellValue(targetSheet, 'G' + (i).ToString(), t2);


                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount3 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.EmploymentStatus.Equals("Continuing Education")
                                              select b).Count();
                            m3 = MaleCount3;
                            SetCellValue(targetSheet, 'H' + (i).ToString(), MaleCount3);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount3 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.EmploymentStatus.Equals("Continuing Education")
                                                select b).Count();
                            f3 = FemaleCount3;
                            SetCellValue(targetSheet, 'I' + (i).ToString(), FemaleCount3);
                        }
                        var t3 = m3 + f3;

                        SetCellValue(targetSheet, 'J' + (i).ToString(), t3);

                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleCount4 = (from b in context.StudentProfiles
                                              where b.TrainingCenter.Equals(Org)
                                              && b.Gender.Equals("Male")
                                              && b.EmploymentStatus.Equals("Drop Out")
                                              select b).Count();
                            m4 = MaleCount4;
                            SetCellValue(targetSheet, 'K' + (i).ToString(), MaleCount4);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleCount4 = (from b in context.StudentProfiles
                                                where b.TrainingCenter.Equals(Org)
                                                && b.Gender.Equals("Female")
                                                && b.EmploymentStatus.Equals("Drop Out")
                                                select b).Count();
                            f4 = FemaleCount4;
                            SetCellValue(targetSheet, 'L' + (i).ToString(), FemaleCount4);
                        }
                        var t4 = m4 + f4;

                        SetCellValue(targetSheet, 'M' + (i).ToString(), t4);

                        SetCellValue(targetSheet, 'N' + (i).ToString(), t1 + t2 + t3 + t4);

                        i++;
                    }



                    // Select a range of data for the Pivot Table.
                    Excel.Range last = targetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    pivotData = targetSheet.get_Range("A3", last);
                    //pivotData.Merge(Type.Missing);

                    if (excelApplication.Application.Sheets.Count < 2)
                    {
                        targetSheet = (Excel.Worksheet)excelWorkBook.Worksheets.Add();
                    }
                    else
                    {
                        targetSheet = (Excel.Worksheet)(excelWorkBook.Worksheets[2]);
                    }

                    SetCellValue(targetSheet, "A1", "CENTRE NAME");
                    SetCellValue(targetSheet, "B1", "EMPLOYED");
                    SetCellValue(targetSheet, "E1", "NOT EMPLOYED");
                    SetCellValue(targetSheet, "H1", "CONTINUING EDUCATION");
                    SetCellValue(targetSheet, "K1", "DROP OUT");
                    SetCellValue(targetSheet, "N1", "COMPLETE TOTAL");

                    Excel.Range Merge1 = targetSheet.get_Range("B1", "D1");
                    Merge1.Merge();

                    Excel.Range Merge2 = targetSheet.get_Range("E1", "G1");
                    Merge2.Merge();

                    Excel.Range Merge3 = targetSheet.get_Range("H1", "J1");
                    Merge3.Merge();

                    Excel.Range Merge4 = targetSheet.get_Range("K1", "M1");
                    Merge4.Merge();

                    Excel.Range Heading = targetSheet.get_Range("A1", "N1");
                    Heading.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    Heading.Borders.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                    Heading.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    targetSheet.Name = "BASED ON EMPLOYMENT STATUS";

                    // Select location of the Pivot Table.
                    pivotDestination = targetSheet.get_Range("A2", useDefault);
                    pivotDestination.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    // Add a Pivot Table to the Worksheet.
                    excelWorkBook.PivotTableWizard(
                        XlPivotTableSourceType.xlDatabase,
                        pivotData,
                        pivotDestination,
                        pivotTableName,
                        true,
                        true,
                        true,
                        true,
                        useDefault,
                        useDefault,
                        false,
                        false,
                        XlOrder.xlDownThenOver,
                        0,
                        useDefault,
                        useDefault
                        );

                    // Set variables for used to manipulate the Pivot Table.
                    pivotTable =
                        (PivotTable)targetSheet.PivotTables(pivotTableName);
                    CENTRE = ((PivotField)pivotTable.PivotFields("CENTRE NAME"));
                    MALE1 = (PivotField)pivotTable.PivotFields(2);
                    FEMALE1 = ((PivotField)pivotTable.PivotFields(3));
                    TOTAL1 = ((PivotField)pivotTable.PivotFields(4));
                    MALE2 = (PivotField)pivotTable.PivotFields(5);
                    FEMALE2 = ((PivotField)pivotTable.PivotFields(6));
                    TOTAL2 = ((PivotField)pivotTable.PivotFields(7));
                    MALE3 = (PivotField)pivotTable.PivotFields(8);
                    FEMALE3 = ((PivotField)pivotTable.PivotFields(9));
                    TOTAL3 = ((PivotField)pivotTable.PivotFields(10));
                    MALE4 = (PivotField)pivotTable.PivotFields(11);
                    FEMALE4 = ((PivotField)pivotTable.PivotFields(12));
                    TOTAL4 = ((PivotField)pivotTable.PivotFields(13));
                    TOTAL_STUDENTS = ((PivotField)pivotTable.PivotFields("COMPLETE TOTAL"));

                    // Format the Pivot Table.
                    pivotTable.Format(XlPivotFormatType.xlReport2);
                    pivotTable.InGridDropZones = false;

                    // Set Centre as a Row Field.
                    CENTRE.Orientation =
                        XlPivotFieldOrientation.xlRowField;

                    // Set Value Field.
                    MALE1.Orientation =
                         XlPivotFieldOrientation.xlDataField;
                    FEMALE1.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL1.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    MALE2.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    FEMALE2.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL2.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    MALE3.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    FEMALE3.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL3.Orientation =
                   XlPivotFieldOrientation.xlDataField;
                    MALE4.Orientation =
                       XlPivotFieldOrientation.xlDataField;
                    FEMALE4.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL4.Orientation =
                   XlPivotFieldOrientation.xlDataField;
                    TOTAL_STUDENTS.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL_STUDENTS.Function = XlConsolidationFunction.xlSum;

                    // Save the Workbook.
                    excelWorkBook.SaveAs(workBookName, useDefault, useDefault,
                        useDefault, useDefault, useDefault,
                        XlSaveAsAccessMode.xlNoChange, useDefault, useDefault,
                        useDefault, useDefault, useDefault);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
                finally
                {
                    // Release the references to the Excel objects.
                    CENTRE = null;
                    TOTAL_STUDENTS = null;
                    MALE1 = null;
                    FEMALE1 = null;
                    MALE2 = null;
                    FEMALE2 = null;
                    MALE3 = null;
                    FEMALE3 = null;
                    MALE4 = null;
                    FEMALE4 = null;
                    TOTAL1 = null;
                    TOTAL2 = null;
                    TOTAL3 = null;
                    TOTAL4 = null;
                    pivotDestination = null;
                    pivotData = null;
                    pivotTable = null;
                    targetSheet = null;

                    excelApplication.Application.DisplayAlerts = false;
                    ((Microsoft.Office.Interop.Excel.Worksheet)excelWorkBook.Worksheets[2]).Delete();
                    excelApplication.Application.DisplayAlerts = true;

                    // Release the Workbook object.
                    if (excelWorkBook != null)
                        excelWorkBook = null;

                    // Release the ApplicationClass object.
                    if (excelApplication != null)
                    {
                        excelApplication.Quit();
                        excelApplication = null;
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    MessageBox.Show("Report Generated");
                }
            }

            if (comboBox1.SelectedIndex == 5)
            {
                // Declare variables that hold references to excel objects.
                ApplicationClass excelApplication = null;
                Workbook excelWorkBook = null;
                Worksheet targetSheet = null;
                PivotTable pivotTable = null;
                Range pivotData = null;
                Range pivotDestination = null;
                PivotField CENTRE = null;
                PivotField MALE = null;
                PivotField FEMALE = null;
                PivotField TOTAL_STUDENTS = null;


                // Declare helper variables.
                string workBookName = @"F:\test\pivottablesample.xlsx";
                string pivotTableName = @"BASED ON AVERAGE SALARY";
                string workSheetName = @"REPORT";


                try
                {
                    // Create an instance of Excel.
                    excelApplication = new ApplicationClass();

                    //Create a workbook and add a worksheet.
                    excelWorkBook = excelApplication.Workbooks.Add(
                        XlWBATemplate.xlWBATWorksheet);
                    targetSheet = (Worksheet)(excelWorkBook.Worksheets[1]);
                    targetSheet.Name = workSheetName;

                    // Add Data to the Worksheet.
                    SetCellValue(targetSheet, "A1", "BASED ON AVERAGE SALARY");
                    SetCellValue(targetSheet, "A2", "CENTRE NAME");
                    SetCellValue(targetSheet, "B2", "MALE");
                    SetCellValue(targetSheet, "C2", "FEMALE");
                    SetCellValue(targetSheet, "D2", "TOTAL");




                    var OrgCount = (from c in db.StudentProfiles
                                    group c.TrainingCenter by c.TrainingCenter into uniqueIds
                                    select uniqueIds.FirstOrDefault()).Count();





                    var OrgName = (from c in db.StudentProfiles
                                   select c.TrainingCenter).Distinct();


                    int i = 3;
                    foreach (var Org in OrgName)
                    {

                        SetCellValue(targetSheet, 'A' + (i).ToString(), Org);

                        Double AvgMaleSal = 0;
                        Double AvgFemaleSal = 0;
                        Double SumMaleSal=0;
                        Double SumFemaleSal = 0;

                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var MaleSalary = from b in context.StudentProfiles
                                             join c in context.Placements on b.Uid equals c.StudentUid
                                             where c.Salary != " " && b.TrainingCenter.Equals(Org) && b.Gender.Equals("Male")
                                             select c.Salary;
                          
                            var MaleSalaryCount = (from b in context.StudentProfiles
                                             join c in context.Placements on b.Uid equals c.StudentUid
                                             where c.Salary != " " && b.TrainingCenter.Equals(Org) && b.Gender.Equals("Male")
                                             select c.Salary).Count();
                            
                            foreach (var sal in MaleSalary)
                            {
                                SumMaleSal = Convert.ToDouble(sal) + SumMaleSal;
                            }

                            AvgMaleSal = Convert.ToDouble(SumMaleSal / MaleSalaryCount);
                            SetCellValue(targetSheet, 'B' + (i).ToString(), AvgMaleSal);
                        }
                        using (var context = new QSStagingDbContext())
                        {
                            // Query for all blogs with names starting with B 
                            var FemaleSalaryCount = (from b in context.StudentProfiles
                                             join c in context.Placements on b.Uid equals c.StudentUid
                                             where c.Salary != " " && b.TrainingCenter.Equals(Org) && b.Gender.Equals("Female")
                                             select c.Salary).Count();

                            var FemaleSalary = from b in context.StudentProfiles
                                                join c in context.Placements on b.Uid equals c.StudentUid
                                                where c.Salary != " " && b.TrainingCenter.Equals(Org) && b.Gender.Equals("Female")
                                                select c.Salary;

                            foreach (var sal in FemaleSalary)
                            {
                                SumFemaleSal = Convert.ToDouble(sal) + SumFemaleSal;
                            }
                            AvgFemaleSal = Convert.ToDouble(SumFemaleSal / FemaleSalaryCount);
                            SetCellValue(targetSheet, 'C' + (i).ToString(), AvgFemaleSal);
                        }
                        Double t = AvgMaleSal + AvgFemaleSal;

                        SetCellValue(targetSheet, 'D' + (i).ToString(), t);

                        i++;
                    }



                    // Select a range of data for the Pivot Table.
                    Excel.Range last = targetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    pivotData = targetSheet.get_Range("A2", last);

                    if (excelApplication.Application.Sheets.Count < 2)
                    {
                        targetSheet = (Excel.Worksheet)excelWorkBook.Worksheets.Add();
                    }
                    else
                    {
                        targetSheet = (Excel.Worksheet)(excelWorkBook.Worksheets[2]);
                    }

                    targetSheet.Name = "BASED ON AVERAGE SALARY";

                    // Select location of the Pivot Table.
                    pivotDestination = targetSheet.get_Range("A1", useDefault);

                    // Add a Pivot Table to the Worksheet.
                    excelWorkBook.PivotTableWizard(
                        XlPivotTableSourceType.xlDatabase,
                        pivotData,
                        pivotDestination,
                        pivotTableName,
                        true,
                        true,
                        true,
                        true,
                        useDefault,
                        useDefault,
                        false,
                        false,
                        XlOrder.xlDownThenOver,
                        0,
                        useDefault,
                        useDefault
                        );

                    // Set variables for used to manipulate the Pivot Table.
                    pivotTable =
                        (PivotTable)targetSheet.PivotTables(pivotTableName);
                    CENTRE = ((PivotField)pivotTable.PivotFields("CENTRE NAME"));
                    MALE = ((PivotField)pivotTable.PivotFields("MALE"));
                    FEMALE = ((PivotField)pivotTable.PivotFields("FEMALE"));
                    TOTAL_STUDENTS = ((PivotField)pivotTable.PivotFields("TOTAL"));

                    // Format the Pivot Table.
                    pivotTable.Format(XlPivotFormatType.xlReport2);
                    pivotTable.InGridDropZones = false;

                    // Set Centre as a Row Field.
                    CENTRE.Orientation =
                        XlPivotFieldOrientation.xlRowField;

                    // Set Value Field.
                    MALE.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    FEMALE.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL_STUDENTS.Orientation =
                        XlPivotFieldOrientation.xlDataField;
                    TOTAL_STUDENTS.Function = XlConsolidationFunction.xlSum;

                    // Save the Workbook.
                    excelWorkBook.SaveAs(workBookName, useDefault, useDefault,
                        useDefault, useDefault, useDefault,
                        XlSaveAsAccessMode.xlNoChange, useDefault, useDefault,
                        useDefault, useDefault, useDefault);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    // Release the references to the Excel objects.
                    CENTRE = null;
                    TOTAL_STUDENTS = null;
                    MALE = null;
                    FEMALE = null;
                    pivotDestination = null;
                    pivotData = null;
                    pivotTable = null;
                    targetSheet = null;

                    excelApplication.Application.DisplayAlerts = false;
                    ((Microsoft.Office.Interop.Excel.Worksheet)excelWorkBook.Worksheets[2]).Delete();
                    excelApplication.Application.DisplayAlerts = true;

                    // Release the Workbook object.
                    if (excelWorkBook != null)
                        excelWorkBook = null;

                    // Release the ApplicationClass object.
                    if (excelApplication != null)
                    {
                        excelApplication.Quit();
                        excelApplication = null;
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    MessageBox.Show("Report Generated");
                }
            }

        }
    }
}


