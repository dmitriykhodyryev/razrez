using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
//using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.IO.Ports;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel; 
namespace ComplexObjects
{
    public class ExcelServ
    {
        Excel.Application XL = new Excel.Application();
        public Workbook wbook;
        public Range range;
        public string SheetName = "";
        public void Open(string WorkBookName)
        {
            XL.ShowStartupDialog = false;
            wbook = XL.Workbooks.Open(WorkBookName,
                false, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
        }
        public void Close()
        {
            XL.DisplayAlerts = false;
            wbook.Save();
            wbook.Close(false, Type.Missing, Type.Missing);
            XL.Quit();
        }
        public void findandcreateSheet(string sheetName)
        {
            Sheets excelsheets = wbook.Sheets;
            Worksheet shabsheet = (Worksheet)excelsheets.get_Item(1);
            Worksheet destsheet, somesheet;
            bool issheetexist = false;
            destsheet = (Worksheet)excelsheets.get_Item(1);
            for (int i = 1; i < excelsheets.Count + 1; i++)
            {
                somesheet = (Worksheet)excelsheets.get_Item(i);
                if ((string)somesheet.Name == sheetName)
                {
                    destsheet = (Worksheet)excelsheets.get_Item(i);
                    issheetexist = true;
                };
            };
            if (!issheetexist)
            {
                shabsheet.Copy(Type.Missing, excelsheets[excelsheets.Count]);
                destsheet = (Worksheet)excelsheets[excelsheets.Count];
                destsheet.Name = sheetName;
            };
            range = destsheet.Cells;
            Write(5, 6, sheetName);
            SheetName = sheetName;
        }
        public void Read(int row, int column, ref double value)
        {
            Range r = (Range)range.get_Item(row, column);
            if ((bool)r.HasFormula)
            {
              r.Calculate();
            };

            string str =(r.Value2==null)?"": r.Value2.ToString();
            if (str != "")
            {
                value = (double)r.Value2;
            }
            else value = 0.0;
            
            // value=(double)((Range)range.get_Item(row, column)).Value2;
        }
        public void Read(int row, int column, ref string value)
        {
            Range rng=((Range)range.get_Item(row, column));
            value =(rng.Value2==null)?null:((double) rng.Value2).ToString();
        }
        public void Read(int row, int column, ref int value)
        {
            value= (int)((Range)range.get_Item(row, column)).Value2;
        }
        public void UniWrite(int sheet, int row, int column, double value)
        {
            Sheets excelsheets = wbook.Sheets;
            Worksheet firstsheet = (Worksheet)excelsheets[sheet];
            Range frange = firstsheet.Cells;
            ((Range)frange.get_Item(row, column)).Value2 = value;
        }
        public void UniWrite(int sheet, int row, int column, string value)
        {
            Sheets excelsheets = wbook.Sheets;
            Worksheet firstsheet = (Worksheet)excelsheets[sheet];
            Range frange = firstsheet.Cells;
            ((Range)frange.get_Item(row, column)).Value2 = value;
        }
        public void UniRead( int sheet, int row, int column, ref string value)
        {
            Sheets excelsheets = wbook.Sheets;
            Worksheet firstsheet = (Worksheet)excelsheets[sheet];
            Range frange = firstsheet.Cells;
            Range rng = ((Range)frange.get_Item(row, column));
            value = (rng.Value2 == null) ? null : rng.Value2.ToString();
            value = (value == null) ? "" : value;
            value = value.Replace("ё", "е");
        }
        public void xldigitread(int sheet, int row, int column, ref string value)
        {
            Sheets excelsheets = wbook.Sheets;
            Worksheet firstsheet = (Worksheet)excelsheets[sheet];
            Range frange = firstsheet.Cells;
            Range rng = ((Range)frange.get_Item(row, column));
            value = (rng.Value2 == null) ? null : rng.Value2.ToString();
            value = (value == null) ? "" : value;
            value = value.Replace("ё", "е");
        }
        public void xlskwaread( int sheet, int row, int column, ref string value)
        {
            Sheets excelsheets = wbook.Sheets;
            Worksheet firstsheet = (Worksheet)excelsheets[sheet];
            Range frange = firstsheet.Cells;
            Range rng = ((Range)frange.get_Item(row, column));
            value = (rng.Value2 == null) ? null : rng.Value2.ToString();
            value = (value == null) ? "" : value;
            value = value.Replace(".", "999777999");
            value = value.Replace(",", "999777999");
            value = Regex.Replace(value, @"[^\d]", ";");
            value = value.Replace("999777999",".");
        }
        public void xlvodaread( int sheet, int row, int column, ref string value)
        {
            Sheets excelsheets = wbook.Sheets;
            Worksheet firstsheet = (Worksheet)excelsheets[sheet];
            Range frange = firstsheet.Cells;
            Range rng = ((Range)frange.get_Item(row, column));
            value = (rng.Value2 == null) ? null : rng.Value2.ToString();
            value = (value == null) ? "" : value;
            value = value.Replace(".", "999777999");
            value = value.Replace(",", "999777999");
            value = Regex.Replace(value, @"[^\d]", "");
            value = value.Replace("999777999", ".");
            if (!value.Contains(".")) { value = value.Trim(); }
            else 
            { 
                value = value.Trim(); 
                if (value.Length<3) value = "";
            };
        }
        public void UniRead(int sheet,int row, int column, ref double value)
        {
            Sheets excelsheets = wbook.Sheets;
            Worksheet firstsheet = (Worksheet)excelsheets[sheet];
            Range frange = firstsheet.Cells;
            Range rng = ((Range)frange.get_Item(row, column));
            string g = (rng.Value2 == null) ? "" : rng.Value2.ToString();
            double d = 0;
            bool b = double.TryParse(g, out d);
            value = b ? d : 999999;
        }
        public void FirstWrite(int row, int column, double value)
        {
            Sheets excelsheets = wbook.Sheets;
            Worksheet firstsheet = (Worksheet)excelsheets[1];
            Range frange = firstsheet.Cells;
            ((Range)frange.get_Item(row, column)).Value2 = value;
        }
        public void FirstRead(int row, int column, ref string value)
        {
            Sheets excelsheets = wbook.Sheets;
            Worksheet firstsheet = (Worksheet)excelsheets[1];
            Range frange = firstsheet.Cells;
            Range rng = ((Range)frange.get_Item(row, column));
            value = (rng.Value2 == null) ? null : rng.Value2.ToString();
        }
        public void FirstRead(int row, int column, ref double value)
        {
            Sheets excelsheets = wbook.Sheets;
            Worksheet firstsheet = (Worksheet)excelsheets[1];
            Range frange = firstsheet.Cells;
            Range rng = ((Range)frange.get_Item(row, column));
            string g = (rng.Value2==null)?"":rng.Value2.ToString();
            double d = 0;
            bool b = double.TryParse(g,out d);
            value = b ? d: 999999;
        }
        public bool FindSheet( ref Worksheet destsheet, string SheetName)
        {
            bool issheetexist = false;
            Sheets excelsheets = wbook.Sheets;
            Worksheet somesheet;
            for (int i = 1; i < excelsheets.Count + 1; i++)
            {
                somesheet = (Worksheet)excelsheets.get_Item(i);
                if ((string)somesheet.Name == SheetName)
                {
                    destsheet = (Worksheet)excelsheets.get_Item(i);
                    issheetexist = true;
                }
                break;
            };
            return issheetexist;
        }
        public void Write(int row, int column, double value, string SheetName)
        {
            Worksheet destsheet;
            Sheets excelsheets = wbook.Sheets;
            destsheet = (Worksheet)excelsheets.get_Item(1);
            if (FindSheet(ref destsheet,SheetName))
            {
                Range frange = destsheet.Cells;
                ((Range)frange.get_Item(row, column)).Value2 = value;
            };
        }
        public bool Read(int row, int column, ref string value, string SheetName)
        {
            Worksheet  destsheet;
            Sheets excelsheets = wbook.Sheets;
            destsheet = (Worksheet)excelsheets.get_Item(1);
            if (FindSheet(ref destsheet, SheetName))
            {
                Range frange = destsheet.Cells;
                Range rng = ((Range)frange.get_Item(row, column));
                value = (rng.Value2 == null) ? null : ((double)rng.Value2).ToString();
                return true;
            }
            else 
                return false;
        }
        public void Write(int row, int column, double value)
        {
            ((Range)range.get_Item(row, column)).Value2 = value;
        }
        public void Write(int row, int column, string value)
        {
            ((Range)range.get_Item(row, column)).Value2 = value;
        }
        public void Write(int row, int column, int value)
        {
            ((Range)range.get_Item(row, column)).Value2 = value;
        }
        public void WriteArray(int row, int column, ref List<double> xlist, ref List<double> ylist)
        {
            for (int i = 0; i < xlist.Count; i++)
            {
                Write(row + i, column, xlist[i]);
                Write(row + i, column+1, ylist[i]);
            };
        }
        public void WriteArray(int row, int column, ref List<double> tlist, ref List<double> flist, ref List<double> slist)
        {
            for (int i = 0; i < tlist.Count; i++)
            {
                Write(row + i, column, tlist[i]);
                Write(row + i, column + 1, flist[i]);
                Write(row + i, column + 2, slist[i]);
            };
        }
        public int rcount(int sheet)
        {
            Sheets excelsheets = wbook.Sheets;
            Worksheet firstsheet = (Worksheet)excelsheets[sheet];

            int rcount = firstsheet.Rows.Count;
            return rcount;
        }
        
       

    }
}
