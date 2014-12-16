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

            //最大テーブル大きさ
            public static int maxx = 10;　//項目数
            public static int maxy = 250;　//件数
            //訂正後テーブル大きさ
            public static int maxfx = 10;　//項目数
            public static int maxfy = 250;　//件数        
        }
    //読み込みボタンよりExcelから読み込みます
        private void button1_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            //読み込みファイル名を指定します
            string ExcelBookFileName = @"C:\Users\淳一\Documents\Visual Studio 2013\Projects\かわせみ検索v0\かわせみ検索v0\bin\Debug\test.xlsx";

            //2次元配列ar[row,columun]を設定します。
            string[,] ar = new string[nx.maxx+20, nx.maxy+20];
            //string[,] arr = new string[nx.maxfx+10, nx.maxfy+10];
            //string[,] index = new string[nx.maxfy+10];

            
            //Excelアプリケーションオブジェクトを作成します。アプリケーションウィンドウは非表示にします。
            ExcelApp.Visible = false;
            
            //Excelを読み込みます
            Workbook wb = ExcelApp.Workbooks.Open(ExcelBookFileName);

            //ExcelのSheet1を指定
            Worksheet ws1 = wb.Sheets[1];
            ws1.Select(Type.Missing);

            int fake_y = -1;
            for (int y = 1; y < nx.maxy; y++)
            {
                Range rgn = ws1.Cells[y, 1];
                dynamic val1 = rgn.Value2;
                String t_tex = Convert.ToString(val1);
                //ar[0, y] = y.ToString();
                //ar[1, y] = "*"+fake_y.ToString();

                if (t_tex == null)
                {
                //    ar[2, y] = "*";
                }
                else
                {
                fake_y++;
                
                   for (int x = 1; x < nx.maxx; x++)
                    {
                     Range index = ws1.Cells[y, x];
                     dynamic val = index.Value2;
                     // arr[x, y] += x.ToString() + "-" + y.ToString() + Convert.ToString(val);
                     ar[x-1, fake_y] += Convert.ToString(val);

                    }
                }

                
            }
            wb.Close(false); //ブッククローズ
            ExcelApp.Quit(); //Excel終了


      //データグリッドビューに代入



            int rim_y = fake_y + 1;//読み込み位置限度の設定

            label1.Text = nx.maxfy.ToString()+"件読み込み、有効"+rim_y+"件";


            dataGridView1.ColumnCount = nx.maxx+3; //gridviewの行・列設定
            dataGridView1.RowCount = rim_y;

            for (int y = 0; y < rim_y; y++)
            {
            　　    for (int x = 0; x < nx.maxx; x++)
                    {
                    dataGridView1[x, y].Value =ar[x, y];//配列arをグリッドに並べる
                    }                     
                }
            }
           

        }

    }

