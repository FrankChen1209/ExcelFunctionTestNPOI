using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace ExcelFunctionTestNPOI
{
    public partial class Form1 : Form
    {
        static HSSFWorkbook mwk = new HSSFWorkbook();
        ISheet sheet = mwk.CreateSheet("例子");
        static HSSFWorkbook hssfworkbook;
        static int[,] IntensityDataValueIns = new int[100, 100];//存放插值后的光强数据
        static int[] Num = new int[100];

        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            for (int i = 0; i <100; i++)//Initialize array.
            {
                for (int j = 0; j < 100; j++)
                {
                    IntensityDataValueIns[i, j] = -1;
                }
                Num[i] = 0;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            StringBuilder temp = new StringBuilder();
            StringBuilder temp1 = new StringBuilder();
            StringBuilder temp2 = new StringBuilder();
            //创建工作薄
            HSSFWorkbook wk = new HSSFWorkbook();
            //创建一个名称为mySheet的表
            ISheet tb = wk.CreateSheet("mySheet");
            //创建一行，此行为第二行
            IRow row = tb.CreateRow(1);
            for (int i = 0; i < 200; i++)
            {
                ICell cell = row.CreateCell(i);  //在第二行中创建单元格
                cell.SetCellValue(i);//循环往第二行的单元格中添加数据
            }
            using (FileStream fs = File.OpenWrite(@"D:/myxls.xls")) //打开一个xls文件，如果没有则自行创建，如果存在myxls.xls文件则在创建是不要打开该文件！
            {
                wk.Write(fs);   //向打开的这个xls文件中写入mySheet表并保存。
                MessageBox.Show("提示：创建成功！");
            }

            for (int i=0;i<Num.Length;i++)
            {
                IntensityDataValueIns[0, i] = Num[i];
                temp.Append(Num[i]);
                temp1.Append(IntensityDataValueIns[0, i]);
                temp2.Append(IntensityDataValueIns[1, i]);
            }
            MoveDataUpDown(-1);

            temp2.Append("\r\n");
            for (int i = 0; i < Num.Length; i++)
            {
                temp2.Append(IntensityDataValueIns[1, i]);
            }
            textBox1.Text =Convert.ToString( temp)+"\r\n";
            textBox1.Text += Convert.ToString(temp1) + "\r\n";
            MoveDataUpDown(-1);
            textBox1.Text += Convert.ToString(temp2) + "\r\n";
        }
        //Move the Intensity data up or down. for example, -1 represent move down for 1 step, while 1 represent move up for 1 step.
        public void MoveDataUpDown(int StepDirection)
        {
            int[] temp = new int[100];
            if (StepDirection == -1)//Means move the data down for 1 step.
            {
                for (int i = 0; i < 100-1; i++)
                {
                    for (int j = 0; j < 100; j++)
                    {
                        IntensityDataValueIns[i + 1, j] = IntensityDataValueIns[i, j];
                    }
                }
                for(int i=0;i<100;i++)
                {
                    if ((IntensityDataValueIns[0, i] = 2 * IntensityDataValueIns[1, i] - IntensityDataValueIns[2, i]) < 0)
                    {
                        IntensityDataValueIns[0, i] = 2 * IntensityDataValueIns[1, i] - IntensityDataValueIns[2, i];
                    }
                    else
                        IntensityDataValueIns[0, i] = Convert.ToInt32(IntensityDataValueIns[1, i] * 0.75) ;
                }
            }
            else if (StepDirection == -2)//Means move the data down for 2 step.
            {

            }
            else if(StepDirection==1)//Means move the data up for 1 step.
            {

            }
            else if(StepDirection==2)//Means move the data up for 2 step.
            {

            }
        }
        //Use the textbox to display the 2 direction array.
        public string DisplayArray(int[,] array)
        {
            string temp=" ";
            for(int i=0;i<100;i++)
            {
                for(int j=0;j<100;j++)
                {
                    temp += array[i, j];
                }
                temp += "\r\n";
            }
            return temp;
        }
        //读取07版及以上Excel文件内容方法
        private void button2_Click(object sender, EventArgs e)
        {
            StringBuilder sbr = new StringBuilder();
            StringBuilder testarray = new StringBuilder();
            StringBuilder testarray1 = new StringBuilder();
            string CheckTestMethod;//check the test method.
            int CheckDataLocation = 0 ;//make sure the useful data's location. When the CheckDataLocation==2, we have found the location.
            int[,] IntensityData = new int[50, 50];

            using (FileStream fs = File.OpenRead(@"D:/进近灯ICAO A2-1 嵌入式 白色 20170310113752.xlsx"))   //打开myxls.xls文件
            {
                XSSFWorkbook wk = new XSSFWorkbook(fs);   //把xls文件中的数据写入wk中
                int IntensityDataRow = 0;
                int IntensityDataCol = 0;
                for(int i=0;i<50;i++)//Initialize array.
                {
                    for(int j=0;j<50;j++)
                    {
                        IntensityData[i, j] = -1;
                    }
                }
                for (int i = 0; i < wk.NumberOfSheets; i++)  //NumberOfSheets是myxls.xls中总共的表数
                {
                    ISheet sheet = wk.GetSheetAt(i);   //读取当前表数据
                    for (int j = 0; j <= sheet.LastRowNum; j++)  //LastRowNum 是当前表的总行数
                    {
                        IRow row = sheet.GetRow(j);  //读取当前行数据
                        if (row != null)
                        {
                            if(CheckDataLocation==2&&IntensityDataRow<50)
                            {
                                IntensityDataRow++;
                            }
                            sbr.Append("-------------------------------------\r\n"); //读取行与行之间的提示界限
                            for (int k = 0; k <= row.LastCellNum; k++)  //LastCellNum 是当前行的总列数
                            {
                                ICell cell = row.GetCell(k);  //当前表格
                                if (cell != null)
                                {
                                    sbr.Append(cell.ToString());   //获取表格中的数据并转换为字符串类型
                                    CheckTestMethod = cell.ToString();
                                    if (CheckTestMethod == "TT_INTEGRITY_TEST")
                                        CheckDataLocation++;
                                    if(CheckDataLocation==2)//Get the right location. Then, we begin to logo the data.
                                    {
                                        //CheckDataLocation = 0;
                                        IntensityDataCol = k-1;
                                        if (IntensityDataCol >= 0 && IntensityDataCol < 50)
                                        {
                                            IntensityData[IntensityDataRow, IntensityDataCol] = Convert.ToInt32(cell.ToString());
                                        }
                                    }
                                }


                            }
                        }
                    }
                }
            }
            for (int i = 0; i < 50; i++)
            {
                testarray.Append("\r\n");
                for (int j = 0; j < 50; j++)
                {
                    if (IntensityData[i, j] != -1)
                    {
                        testarray.Append(IntensityData[i, j].ToString());
                        testarray.Append("_");
                        Num[j] = IntensityData[i, j];
                    }
                    else
                        continue;
                }
                int[] temp = InsValue(Num);
                for(int k=0;k<temp.Length;k++)
                {
                    IntensityDataValueIns[i, k] = temp[k];
                }
            }

            testarray.ToString();
            using (StreamWriter wr = new StreamWriter(new FileStream(@"D:/TestArray.txt", FileMode.Append)))  //把读取xls文件的数据写入myText.txt文件中
            {
                wr.Write(testarray.ToString());
                wr.Flush();
            }

            for (int i = 0; i < 100; i++)
            {
                testarray1.Append("\r\n");
                for (int j = 0; j < 100; j++)
                {
                    if (IntensityDataValueIns[i, j] != -1)
                    {
                        testarray1.Append(IntensityDataValueIns[i, j].ToString());
                        testarray1.Append("_");
                    }
                    else
                        continue;
                }
            }

            testarray1.ToString();
            using (StreamWriter wr = new StreamWriter(new FileStream(@"D:/TestArray1.txt", FileMode.Append)))  //把读取xls文件的数据写入myText.txt文件中
            {
                wr.Write(testarray1.ToString());
                wr.Flush();
            }
            sbr.ToString();
            using (StreamWriter wr = new StreamWriter(new FileStream(@"D:/myText.txt", FileMode.Append)))  //把读取xls文件的数据写入myText.txt文件中
            {
                wr.Write(sbr.ToString());
                wr.Flush();
            }

        }
        //将获取到的光强值，插成0.5度一测的表
        public int[] InsValue(int[] array)
        {
            int arraylenth = 0;
            for(int i=0;i<array.Length;i++)//获取该数组中有效数值（光强值，而非初始值“-1”）的个数。
            {
                if (array[i] != -1)
                {
                    arraylenth++;
                }
                else
                    break;
            }
            int[] temparray = new int[arraylenth*2+1];//存放插值后的数据，插值后的数据总数是原来的2倍+1。
            for(int i=0;i<arraylenth;i++)//保存数组temparray脚码为1,3,5...的数据
            {
                temparray[i * 2+1] = array[i];
            }
            for(int i=1;i<arraylenth;i++)//保存数组temparray脚码为2,4,6...的数据
            {
                temparray[i * 2] = (temparray[i * 2 - 1] + temparray[i * 2 + 1]) / 2;
            }
            if ((temparray[0] = temparray[1] - (temparray[2] - temparray[1])) > 0)//保存数组temparray脚码为0的数据
            {
                temparray[0] = temparray[1] - (temparray[2] - temparray[1]);
            }
            else
                temparray[0] = (int)(temparray[1] * 0.75);
            return temparray;
        }
        //读取07版及以上Excel文件内容方法，并将读取到的内容存放在“myText.txt”中
        public void ReadExcel()
        {
            StringBuilder sbr = new StringBuilder();
            using (FileStream fs = File.OpenRead(@"D:/进近灯ICAO A2-1 嵌入式 白色 20170310113752.xlsx"))   //打开myxls.xls文件
            {
                XSSFWorkbook wk = new XSSFWorkbook(fs);   //把xls文件中的数据写入wk中
                for (int i = 0; i < wk.NumberOfSheets; i++)  //NumberOfSheets是myxls.xls中总共的表数
                {
                    ISheet sheet = wk.GetSheetAt(i);   //读取当前表数据
                    for (int j = 0; j <= sheet.LastRowNum; j++)  //LastRowNum 是当前表的总行数
                    {
                        IRow row = sheet.GetRow(j);  //读取当前行数据
                        if (row != null)
                        {
                            sbr.Append("-------------------------------------\r\n"); //读取行与行之间的提示界限
                            for (int k = 0; k <= row.LastCellNum; k++)  //LastCellNum 是当前行的总列数
                            {
                                ICell cell = row.GetCell(k);  //当前表格
                                if (cell != null)
                                {
                                    sbr.Append(cell.ToString());   //获取表格中的数据并转换为字符串类型
                                }
                            }
                        }
                    }
                }
            }
            sbr.ToString();
            using (StreamWriter wr = new StreamWriter(new FileStream(@"D:/myText.txt", FileMode.Append)))  //把读取xls文件的数据写入myText.txt文件中
            {
                wr.Write(sbr.ToString());
                wr.Flush();
            }
        }

       
    }
}

