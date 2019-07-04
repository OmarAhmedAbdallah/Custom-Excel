using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace CreateExcelSheet.Models
{
    class Excel
    {
        public static bool CreateFile()
        {
            try
            {
                //try to create file with extension xlsx 
                //file name is Employees
                //file has workSheet called EmployeeTable
                //file will be saved in this path "F:\ITI_SD\EFG\ExcelSheets\"
                ExcelPackage excel = new ExcelPackage();

                //file has workSheet called EmployeeTable
                excel.Workbook.Worksheets.Add("EmployeeTable");

                //Create Table
                CreateTable(excel);

                //fill header
                FillTableHeader(excel);

                //Fill Data
                FillData(excel);
                
                //format Bonus
                //number with 2 decimal places and thousand separator and money symbol
                // FormatCellsToMoney(excel, Range, Format);
                FormatCellsToMoney(excel, "I8:I12", "$#,##0.00");


           

                var worksheet = excel.Workbook.Worksheets["EmployeeTable"];

                

                //Create File
                string FilePath = System.IO.File.ReadAllText(@"..\..\..\ExcelPath.txt");

                FileInfo excelFile = new FileInfo(FilePath);
                excel.SaveAs(excelFile);
            }
            catch
            {
                return false;
            }
           
            return true;
        }
        public static void CreateTable(ExcelPackage excel)
        {
            //Defining the tables parameters
            var worksheet = excel.Workbook.Worksheets["EmployeeTable"];

            ExcelRange rg = worksheet.Cells[7, 5, 12, 9];
            string tableName = "TableEmpolyee";

            //Ading a table to a Range
            ExcelTable tab = worksheet.Tables.Add(rg, tableName);

            //Formating the table style
            tab.TableStyle = TableStyles.Light8;
        }


        public static void FillTableHeader(ExcelPackage excel)
        {

            var worksheet = excel.Workbook.Worksheets["EmployeeTable"];
            //Table Header  Data               
            var headerRow = new List<string[]>()
                {
                    new string[] { "Employee ID", "Employee First Name", "Last Name", "Floor" , "Bonus" }
                };

            using (var HeaderCells = worksheet.Cells["E7:I7"])
            {
                // Popular header row data
                //range of first row Cells[Row Number,Col number]
                //Add header of table
                
                HeaderCells.LoadFromArrays(headerRow);



                //Make cell width Fit
                HeaderCells.AutoFitColumns();

                //change cell background color
                HeaderCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                HeaderCells.Style.Fill.BackgroundColor.SetColor(Color.Black);

                //change text color
                HeaderCells.Style.Font.Color.SetColor(Color.White);
                worksheet.Cells["F7"].Style.Font.Color.SetColor(Color.Red);

                //Make font bold
                HeaderCells.Style.Font.Bold = true;
            }

        }

        public static void FillData(ExcelPackage excel)
        {
            List<Employee> EmpList = Employee.GetEmployees();

            var worksheet = excel.Workbook.Worksheets["EmployeeTable"];
            //it write in excel from row 8 for 5 col
            //it take collection (List<Employee>)
            worksheet.Cells[8, 5].LoadFromCollection(EmpList);

        }

        public static void FormatCellsToMoney(ExcelPackage excel ,string Range,string Format)
        {
            var worksheet = excel.Workbook.Worksheets["EmployeeTable"];
            //Formate the Range
            worksheet.Cells[Range].Style.Numberformat.Format = Format;
        }
        
    }
}
