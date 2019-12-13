using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.OpenXml4Net;
using NPOI.OpenXmlFormats;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace 编辑信息
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        static DataTable daty = new DataTable();

        private void button1_Click(object sender, EventArgs e) //导入数据按钮
        {
            OpenFileDialog of = new OpenFileDialog();
            if (of.ShowDialog() == DialogResult.OK)
            {
                string fp = of.FileName.ToString();
                DataTable dt = ExcelToDataTable(fp);
                dataGridView1.DataSource = dt;
                daty = dt;
            }

        }

        //导入外部数据，并将其转换成DataTable
        public static DataTable ExcelToDataTable(string fileName) 
        {
            ISheet sheet = null;  //空表
            DataTable data = new DataTable(); 
            FileStream fs = null; //空数据流
            IWorkbook workbook = null;
            int startRow = 0;

            fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook(fs);
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook(fs);

            sheet = workbook.GetSheetAt(0); //获取第一张数据表

            if (sheet != null)
            {
                IRow firstRow = sheet.GetRow(0);
                int cellCount = firstRow.LastCellNum;//一行最后一个cell的编号 即总的列数


                for (int i = firstRow.FirstCellNum; i < cellCount; ++i) //创建列
                {
                    DataColumn column = new DataColumn(firstRow.GetCell(i).ToString());
                    data.Columns.Add(column);

                }
                startRow = sheet.FirstRowNum + 1;


                //最后一列的标号
                int rowCount = sheet.LastRowNum;
                for (int i = startRow; i <= rowCount; ++i)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue; //没有数据的行默认是null　　　　　　　

                    DataRow dataRow = data.NewRow();
                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell != null) //同理，没有数据的单元格都默认是null
                        {
                            dataRow[j] = cell.ToString();
                        }
                    }
                    data.Rows.Add(dataRow);
                }
            }

            return data;
           
        }

        //气虚证
        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Dat(button2.Text);
            textBox1.Text = Dat(button2.Text).TableName;
        }

        //创建新表，并处理数据，求均值，方差，t值
        public static DataTable Dat(string st)
        {
            DataTable dat = new DataTable(st); //用于显示最后数据处理好的表
            //列内容
            dat.Columns.Add("西医指标");
            dat.Columns.Add("证候为1的均值");
            dat.Columns.Add("证候为1的方差");
            dat.Columns.Add("证候为0的均值");
            dat.Columns.Add("证候为0的方差");
            dat.Columns.Add("t值");
            dat.Columns.Add("p值");

            //行内容
            for (int i = 0; i < 12; i++)
            {
                dat.Rows.Add(daty.Columns[i + 4].Caption);
            }

            DataTable b = new DataTable();  //用于将证候为0和为1的数据分开
            
            //创建列
            for(int i = 0; i < 24; i++)
            {
                b.Columns.Add(i.ToString());
            }

            //创建行
            for(int i = 0; i < daty.Rows.Count; i++)
            {
                b.Rows.Add();
            }

            
            //统计数据
            for (int i = 0; i < daty.Rows.Count; i++)
            {
                for (int j = 0; j < dat.Rows.Count * 2; j += 2)
                    if (daty.Rows[i][st].ToString() == "1")
                        b.Rows[i][j] = daty.Rows[i][4 + j / 2];
                for (int j = 0; j < dat.Rows.Count * 2; j += 2)
                    if (daty.Rows[i][st].ToString() == "0")
                        b.Rows[i][1 + j] = daty.Rows[i][4 + j / 2];
            }

            //处理数据，方差
            int onelen = 0, zerolen = 0;
            for (int i = 0; i < b.Columns.Count; i++)
            {
                if (i % 2 == 0) //证候为1的方差
                {
                    List<double> m = new List<double>();  //存放数据
                    for (int j = 0; j < b.Rows.Count; j++)
                    {
                        string n = b.Rows[j][i].ToString();
                        if (n != "")
                        {
                            m.Add(double.Parse(n));
                        }
                    }

                    double sum1 = 0;
                    for (int j = 0; j < m.Count; j++)  //求和
                    {
                        sum1 += m[j];
                    }
                    double sum2 = 0;
                    for(int j = 0; j < m.Count; j++)  //分子
                    {
                        sum2 += Math.Pow((m[j] - (sum1 / m.Count)), 2);
                    }
                    onelen = m.Count;
                    dat.Rows[i / 2][2] = Math.Round((sum2 / (m.Count - 1)), 3).ToString();
                }
                else if (i % 2 == 1) //证候为0的方差
                {
                    List<double> m = new List<double>();
                    for (int j = 0; j < b.Rows.Count; j++)
                    {
                        string n = b.Rows[j][i].ToString();
                        if (n != "")
                        {
                            m.Add(double.Parse(n));
                        }
                    }

                    double sum1 = 0;
                    for (int j = 0; j < m.Count; j++)  //求和
                    {
                        sum1 += m[j];
                    }
                    double sum2 = 0;
                    for (int j = 0; j < m.Count; j++)  //分子
                    {
                        sum2 += Math.Pow((m[j] - (sum1 / m.Count)), 2);
                    }
                    zerolen = m.Count;
                    dat.Rows[i / 2][4] = Math.Round((sum2 / (m.Count - 1)), 3).ToString();
                }
            }

            //均值
            for (int i = 0; i < b.Columns.Count; i++)
            {
                if (i % 2 == 0)  //证候为1的均值
                {
                    List<double> m = new List<double>();
                    for (int j = 0; j < b.Rows.Count; j++)
                    {
                        string n = b.Rows[j][i].ToString();
                        if (n != "")
                        {
                            m.Add(double.Parse(n));
                        }
                    }
                    double sum1 = 0;
                    for (int j = 0; j < m.Count; j++)  //求和
                    {
                        sum1 += m[j];
                    }
                    dat.Rows[i / 2][1] = Math.Round((sum1 / m.Count), 3).ToString();
                }
                else if (i % 2 == 1)  //证候为0的均值
                {
                    List<double> m = new List<double>();
                    for (int j = 0; j < b.Rows.Count; j++)
                    {
                        string n = b.Rows[j][i].ToString();
                        if (n != "")
                        {
                            m.Add(double.Parse(n));
                        }
                    }

                    double sum1 = 0;
                    for (int j = 0; j < m.Count; j++)  //求和
                    {
                        sum1 += m[j];
                    }
                    dat.Rows[i / 2][3] = Math.Round((sum1 / m.Count), 3).ToString();
                }
            }

            //t值
            for (int i = 0; i < dat.Rows.Count; i++)
            {
                double b1 = double.Parse(dat.Rows[i][2].ToString()) / onelen;
                double b0 = double.Parse(dat.Rows[i][4].ToString()) / zerolen;
                double fm = Math.Sqrt(b1 + b0);
                double t = ((double.Parse(dat.Rows[i][1].ToString())) - (double.Parse(dat.Rows[i][3].ToString()))) / fm;
                dat.Rows[i][5] = Math.Round(t, 3).ToString();
            }

            //p值
            for( int i = 0; i < dat.Rows.Count; i++)
            {
                dat.Rows[i][6] = "不会算";
            }

            return dat;
        }

        //肾虚证
        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Dat(button3.Text);
            textBox1.Text = Dat(button3.Text).TableName;
        }

        //阳虚证
        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Dat(button4.Text);
            textBox1.Text = Dat(button4.Text).TableName;
        }

        //保存数据表的按钮
        private void button5_Click(object sender, EventArgs e)
        {
            SaveFileDialog sf = new SaveFileDialog();
            if (sf.ShowDialog() == DialogResult.OK)
            {
                string fp = sf.FileName.ToString();  //保存路径
                DataTable dt = dataGridView1.DataSource as DataTable;
                DataTableToExcel(dt, textBox1.Text, fp);
            }

        }


        //将DataTable表转换成Excel，传入的data是要保存的表,sheetName传入的是表名，最后传入路径
        public int DataTableToExcel(DataTable data, string sheetName, string fn)
        {
            //和建表很类似
            IWorkbook workbook = null; 
            FileStream fs = null;
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;
            fs = new FileStream(fn, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            if (fn.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook();
            else if (fn.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook();
            if (workbook != null)
            {
                sheet = workbook.CreateSheet(sheetName);
            }
            else
            {
                return -1;
            }

            //第一行，相当于列名
            IRow row = sheet.CreateRow(0);
            for (j = 0; j < data.Columns.Count; ++j)
            {
                row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
            }
            count = 1;

            
            //真正的内容
            for (i = 0; i < data.Rows.Count; ++i)
            {
                IRow row1 = sheet.CreateRow(count);
                for (j = 0; j < data.Columns.Count; ++j)
                {
                    row1.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                }
                ++count;
            }

            workbook.Write(fs); //写入到excel
            return count;
        }
    }
}
