﻿using ClassLibrary;
using NetOffice.ExcelApi;
using System;
using System.Configuration;
using System.Threading.Tasks;

namespace WorkDiary
{
    /// <summary>
    /// Excel操作
    /// </summary>
    public class ExcelUnit : IDisposable
    {

        CellInfo cellInfo { get; set; }
        NetOffice.ExcelApi.Application ExcelApp;
        NetOffice.ExcelApi.Workbook WorkBook;
        Worksheet WorkSheet;
        public string ExcelFilename { get; set; }

        public ExcelUnit(string excelFilename)
        {
            ExcelApp = new NetOffice.ExcelApi.Application();
            ExcelApp.DisplayAlerts = false;

            ExcelFilename = excelFilename;
            cellInfo = new CellInfo();
        }

        public void Dispose()
        {
            try
            {
                WorkBook.Close();
                ExcelApp.Quit();
                ExcelApp.Dispose();
            }
            catch { }
        }

        internal string ReadCell(string cell)
        {
            string cellstr = string.Empty;
            WorkBook = ExcelApp.Workbooks.Open(ExcelFilename);
            WorkSheet = (Worksheet)WorkBook.Worksheets[1];

            try
            {
                cellstr = WorkSheet.Range(cell).Value2.ToString();
            }
            catch { }
            finally
            {
                WorkBook.Close();
            }
            return cellstr;
        }

        internal Person Read()
        {
            WorkBook = ExcelApp.Workbooks.Open(ExcelFilename);
            WorkSheet = (Worksheet)WorkBook.Worksheets[1];

            Person person = new Person();
            try
            {
                person.Id = WorkSheet.Range(cellInfo.IdCell).Value2.ToString();
                person.PersonName = WorkSheet.Range(cellInfo.NameCell).Value2.ToString();
                person.Department = WorkSheet.Range(cellInfo.DepartCell).Value2.ToString();
                person.Company = WorkSheet.Range(cellInfo.CompanyCell).Value2.ToString();
                person.Date = DateTime.FromOADate((double)WorkSheet.Range(cellInfo.DateCell).Value2).ToShortDateString();
                person.DiaryContent = WorkSheet.Range(cellInfo.ContentCell).Value2.ToString();
            }
            catch { }
            finally
            {
                WorkBook.Close();
            }
            return person;
        }

        internal async Task<Person> ReadAsync()
        {
            WorkBook = ExcelApp.Workbooks.Open(ExcelFilename);
            WorkSheet = (Worksheet)WorkBook.Worksheets[1];

            Person person = new Person();
            try
            {
                person.Id = WorkSheet.Range(cellInfo.IdCell).Value2.ToString();
                person.PersonName = WorkSheet.Range(cellInfo.NameCell).Value2.ToString();
                person.Department = WorkSheet.Range(cellInfo.DepartCell).Value2.ToString();
                person.Company = WorkSheet.Range(cellInfo.CompanyCell).Value2.ToString();
                person.Date = DateTime.FromOADate((double)WorkSheet.Range(cellInfo.DateCell).Value2).ToShortDateString();
                person.DiaryContent = WorkSheet.Range(cellInfo.ContentCell).Value2.ToString();
            }
            catch { }
            finally
            {
                WorkBook.Close();
            }
            return person;
        }




        internal bool SaveAs(string filename)
        {
            bool result = false;
            WorkBook = ExcelApp.Workbooks.Open(ExcelFilename);
            WorkSheet = (Worksheet)WorkBook.Worksheets[1];
            try
            {
                WorkSheet.SaveAs(filename);
                if (System.IO.File.Exists(filename))
                {
                    ExcelFilename = filename;
                    result = true;
                }
                return result;
            }
            finally
            {
                WorkBook.Close();
            }
        }

        internal bool SaveAs(Person p, string filename)
        {
            bool result = false;
            WorkBook = ExcelApp.Workbooks.Open(ExcelFilename);
            WorkSheet = (Worksheet)WorkBook.Worksheets[1];
            try
            {
                WorkSheet.Range(cellInfo.IdCell).Value2 = p.Id;
                WorkSheet.Range(cellInfo.NameCell).Value2 = p.PersonName;
                WorkSheet.Range(cellInfo.DateCell).Value2 = p.Date;
                WorkSheet.Range(cellInfo.CompanyCell).Value2 = p.Company;
                WorkSheet.Range(cellInfo.ContentCell).Value2 = p.DiaryContent;
                WorkSheet.Range(cellInfo.DepartCell).Value2 = p.Department;
                WorkSheet.SaveAs(filename);
                if (System.IO.File.Exists(filename))
                {
                    ExcelFilename = filename;
                    result = true;
                }
                return result;
            }
            finally
            {
                WorkBook.Close();
            }
        }

        internal void Write(Person p)
        {
            WorkBook = ExcelApp.Workbooks.Open(ExcelFilename);
            WorkSheet = (Worksheet)WorkBook.Worksheets[1];
            try
            {
                WorkSheet.Range(cellInfo.IdCell).Value2 = p.Id;
                WorkSheet.Range(cellInfo.NameCell).Value2 = p.PersonName;
                WorkSheet.Range(cellInfo.DateCell).Value2 = p.Date;
                WorkSheet.Range(cellInfo.CompanyCell).Value2 = p.Company;
                WorkSheet.Range(cellInfo.ContentCell).Value2 = p.DiaryContent;
                WorkSheet.Range(cellInfo.DepartCell).Value2 = p.Department;
            }
            finally
            {
                WorkBook.Close();
            }
        }
    }

    class CellInfo 
    {
        public string IdCell { get; set; }
        public string NameCell { get; set; }
        public string DepartCell { get; set; }
        public string CompanyCell { get; set; }
        public string DateCell { get; set; }
        public string ContentCell { get; set; }

        public CellInfo()
        {
            IdCell = ConfigurationManager.AppSettings.Get("idcell");
            NameCell = ConfigurationManager.AppSettings.Get("namecell");
            DepartCell = ConfigurationManager.AppSettings.Get("departcell");
            CompanyCell = ConfigurationManager.AppSettings.Get("companycell");
            ContentCell = ConfigurationManager.AppSettings.Get("contentcell");
            DateCell = ConfigurationManager.AppSettings.Get("datecell");
        }


    }
}
