using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel  = Microsoft.Office.Interop.Excel;

namespace IO
{
    class Excel
    {
        public string path { get; set; }
        public int sheet { get; set; }
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        public Excel()
        {
            this.path = @"E:\studia\V semestr\IO\Algorytmy genetyczne\genetykcs3.xls";
            this.sheet = 0;
            //this.ws = 1;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public Excel(string path, int sheet)
        {
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public string ReadCellAsString(int row, int col)
        {
            row++;
            col++;
            if (ws.Cells[row, col].Value2 != null)
            {
                return ws.Cells[row, col].Value2;
            }
            else
                return string.Empty;
        }

        public int ReadCellAsInt(int row, int col)
        {

            if (ws.Cells[row, col].Value2 != null)
            {
                return (int) ws.Cells[row, col].Value2;
            }
            else
                return 0;
        }

        public void WriteToCellInt(int row, int col, int value)
        {
            row++;
            col++;
            ws.Cells[row, col].Value2 = value.ToString();
        }

        public void WriteToCellString(int row, int col, string value)
        {
            row++;
            col++;
            ws.Cells[row, col].Value2 = value;
        }

        public void ReadColumnToArray(int startingRow, int endingRow,  int col, ArrayList array)
        {

            if(startingRow <= endingRow)
            {
                for (int i = 0; i < endingRow - startingRow + 1; i++)
                {
                    if (ws.Cells[i + startingRow - 1, col].Value2 != null)
                    {
                        array.Add(ws.Cells[i + startingRow, col].Value2);
                    }
                }
            }
            
        }

        public void WriteToColumnFromArray(int startingRow, int endingRow, int col, ArrayList array)
        {

            if (startingRow <= endingRow)
            {
                for (int i = 0; i < array.Count; i++)
                {
                    ws.Cells[i+startingRow, col].Value2 = Convert.ToString(array[i]);                   
                }
            }
        }

        public void WriteToColumnFromArrayInts(int startingRow, int endingRow, int col, int[] array, int sizeOfArray)
        {

            if (startingRow <= endingRow)
            {
                for (int i = 0; i < sizeOfArray; i++)
                {
                    ws.Cells[i + startingRow, col].Value2 = Convert.ToString(array[i]);
                }
            }
        }



        public void Save()
        {
            wb.Save();
        }

        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }

        public void Close()
        {
            wb.Close();
        }

        public void UnprotectSheet()
        {
            ws.Unprotect();
        }
    }
}
