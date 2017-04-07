using System;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.HPSF;
using System.IO;

namespace ExcelOutputTest
{
    class Program
    {
        static HSSFWorkbook hssfworkbook;
        static void Main(string[] args)
        {
            //Console.WriteLine("Hello World!");
            InitializeWorkbook();
            //XSSFSheet sheet1 = hssfworkbook.CreateSheet("Sheet1");

        }
        static void InitializeWorkbook()
        {
            hssfworkbook = new HSSFWorkbook();
            //DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            //dsi.Company = "Airsafe Company";
            //hssfworkbook.DocumentSummaryInformation = dsi;
            //SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            //si.Subject = "NPOI SDK Example";
            //hssfworkbook.SummaryInformation = si;
        }
    }
}