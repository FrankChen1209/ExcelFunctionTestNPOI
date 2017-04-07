using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HPSF;
using System.IO;


namespace ConsoleApp1
{
    class Program
    {
        static XSSFWorkbook xwb = new XSSFWorkbook();
        static void Main(string[] args)
        {
            //ISheet sheet1 = xwb.CreateSheet("Sheet1");
            //XSSFRow row;
            //XSSFCell cell;
            //for(int rowIndex=0;rowIndex<9;rowIndex++)
            //{
            //    row = sheet1.CreateRow(rowIndex);
            //    for (int colIndex=0;colIndex<=rowIndex;colIndex++)
            //    {
            //        cell = row.CreateCell(colIndex);
            //        cell.SetCellValue(string.Format("{0}*{1}={2}", rowIndex + 1, colIndex + 1, (rowIndex + 1) * (colIndex + 1)));

            //    }
            //}
            //WriteToFile();
        }
        static void WriteToFile()
        {
            FileStream file = new FileStream(@"test.xlsx", FileMode.Create);
            xwb.Write(file);
            file.Close();
        }
    }
}
