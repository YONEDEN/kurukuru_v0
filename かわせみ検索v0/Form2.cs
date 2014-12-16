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
            public static int maxx = 15;　//項目数
            public static int maxy = 10;　//件数
            //訂正後テーブル大きさ
            public static int maxfx = 10;　//項目数
            public static int maxfy = 10;　//件数        
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            //読み込みファイル名を指定します
            string ExcelBookFileName = @"C:\Users\淳一\Documents\Visual Studio 2013\Projects\かわせみ検索v0\かわせみ検索v0\bin\Debug\test.xlsx";

            //2次元配列ar[row,columun]を設定します。
            string[,] ar = new string[nx.maxx+10, nx.maxy+10];
            string[,] arr = new string[nx.maxfx+10, nx.maxfy+10];
            //string[,] index = new string[nx.maxfy+10];

            
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
                    arr[x, y] += x.ToString() + "-" + y.ToString() + Convert.ToString(val);
                    ar[x, y] += Convert.ToString(val);
                }
            }
            wb.Close(false); //ブッククローズ
            ExcelApp.Quit(); //Excel終了



      //データグリッドビューに代入

            dataGridView1.ColumnCount = nx.maxx+3; //gridviewの行・列設定
            dataGridView1.RowCount = nx.maxy+3;

            //定義
            int fake_y = 1;     //読み込み位置変数の初期設定
            int rim_y = nx.maxy;//読み込み位置限度の設定
            String s = "";      //空白の判定用変数

            for (int y = 1; y < nx.maxy; y++)
            {
                s = ar[1,y];//空白かどうかわからない配列を読み込み
                //fake_y = y;
                dataGridView1[0, y - 1].Value = fake_y;//左端に表示番号表示                
                //dataGridView1[1, y-1].Value = arr[1, fake_y];//
                //dataGridView1[1, y-1].Value = "*";//(成功)
                //dataGridView1[1, fake_y-1].Value = "*";//（成功）
                dataGridView1[1, fake_y-1].Value = "*"+arr[4,y];//(成功)

                if (s == "")
                //空白判定・空白なら
                    {
                    }
                else
                //空白でなかったら
                    fake_y++;
                    if (fake_y > rim_y)
                    {
                        break;
                    }
                    
            　　    for (int x = 1; x < nx.maxx; x++)
                    {
                        //label1.Text = fake_y.ToString();
                        
                    dataGridView1[x+1, fake_y-1].Value =arr[x, y];//配列arをグリッドx+1して2、0から並べる
                        //arr[x, y] = ar[x, fake_y];
                    }                     
                }
            }
 

        }

    }

