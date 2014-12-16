using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;


namespace かわせみ検索v0
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();


        }
        class nx
        {
            //配列のサイズを固定で設定します。
            public static int maxx = 20;
            public static int maxy = 60;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            //読み込みファイル名を指定します
            string ExcelBookFileName = @"C:\Users\淳一\Documents\Visual Studio 2013\Projects\かわせみ検索v0\かわせみ検索v0\bin\Debug\test.xlsx";

            //2次元配列ar[row,columun]を設定します。
            string[,] ar = new string[nx.maxx, nx.maxy];
            
            //Excelアプリケーションオブジェクトを作成します。アプリケーションウィンドウは非表示にします。
            ExcelApp.Visible = false;
            
            //Excelを読み込みます
            Workbook wb = ExcelApp.Workbooks.Open(ExcelBookFileName);

            //ExcelのSheet1を指定
            Worksheet ws1 = wb.Sheets[1];
            ws1.Select(Type.Missing);

            for (int x = 1; x < nx.maxx; x++)
            {
                for (int y = 1; y < nx.maxy; y++)
                {
                    Range rgn = ws1.Cells[y, x];
                    dynamic val = rgn.Value2;
                    ar[x,y] += Convert.ToString(val);
                }
            }
            wb.Close(false); //ブッククローズ
            ExcelApp.Quit(); //Excel終了


            //String st_Result = "";
            //for (int i = 1; i < 10; i++)
            //{
            //    for (int j = 1; j < 10; j++)
            //    {
            //        st_Result += ar[i, j];
            //        st_Result += " ";
            //    }
            //    //st_Result += "\r\n";
            //}

            dataGridView1.ColumnCount = nx.maxx; //gridviewの行・列設定
            dataGridView1.RowCount = nx.maxy;
 
            
            //dataGridView1.Rows[1].Cells[1].Value = "TEST2";
                        
            for (int x = 1; x < nx.maxx; x++)
            {
                for (int y = 1; y < nx.maxy; y++)
                {
                    dataGridView1[x-1, y-1].Value = ar[x, y]; 
                 //   dataGridView1.Rows.Add(ar[j, i]);
                }
            }
            //dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;


            //   MessageBox.Show(st_Result + "を取得しました");


        }

    }
}