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
            public static int maxy = 50;　//件数
            //訂正後テーブル大きさ
            public static int maxfx = 10;　//項目数
            public static int maxfy = 50;　//件数        
        }
    //読み込みボタンよりExcelから読み込みます
        private void button1_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            //読み込みファイル名を指定します
            string ExcelBookFileName = @"C:\Users\淳一\Documents\Visual Studio 2013\Projects\かわせみ検索v0\かわせみ検索v0\bin\Debug\test.xlsx";

            //2次元配列ar[row,columun]を設定します。
            string[,] ar = new string[nx.maxx+20, nx.maxy+20];
            string[,] col = new string[nx.maxfx+10, 2];
            string[] skip_n = new string[]{
                "1",  //A
                "1",  //B
                "0",  //C
                "1",  //D
                "1",  //E
                "1",  //F
                "1",  //G
                "1",  //H
                "1",  //I
                "1",  //J
                "1",  //K
                "1",  //L
                "1",  //M
                "1",  //N
                "1",  //O
                "1",  //P
                "1",  //Q
                };


            //Excelアプリケーションオブジェクトを作成します。アプリケーションウィンドウは非表示にします。
            ExcelApp.Visible = false;
            
            //Excelを読み込みます
            Workbook wb = ExcelApp.Workbooks.Open(ExcelBookFileName);

            //ExcelのSheet1を指定
            Worksheet ws1 = wb.Sheets[1];
            ws1.Select(Type.Missing);

            int fake_x = 0;
            int fake_y = 0;
            for (int x = 1; x < nx.maxx; x++)
                {
                        //Range index = ws1.Cells[1, x];
                        //dynamic val = index.Value2;
                        // arr[x, y] += x.ToString() + "-" + y.ToString() + Convert.ToString(val);
                        //col[x - 1, 0] += Convert.ToString(val);

                        if (skip_n[x] == "0")
                        //非表示スイッチが0の場合
                        {
                        //    Range index = ws1.Cells[1, x];
                        //    dynamic val = index.Value2;
                        }
                        else
                        //非表示スイッチが1の場合
                        {
                            Range index = ws1.Cells[1, x];
                            dynamic val = index.Value2;
                            //ar[x, fake_y - 1] += fake_x+Convert.ToString(val) + skip_n[x];
                            col[fake_x, 0] += Convert.ToString(val);
                            fake_x++; //スキップ列番号を増やす
                        }

                }
            

            for (int y = 2; y < nx.maxy; y++)
            {
                Range rgn = ws1.Cells[y, 1];
                dynamic val1 = rgn.Value2;
                String t_tex = Convert.ToString(val1);
                //ar[0, y] = y.ToString();
                //ar[1, y] = "*"+fake_y.ToString();

                if (t_tex == null)
                {
                //    なにもしない
                }
                else
                {
                fake_y++; //スキップ行番号を増やす
                fake_x = 0;
                   for (int x = 0; x < nx.maxx; x++)
                    {
                       if (skip_n[x] == "0")
                           //非表示スイッチが0の場合
                       {
                           //Range index = ws1.Cells[y, x+1];
                           //dynamic val = index.Value2;
                           // arr[x, y] += x.ToString() + "-" + y.ToString() + Convert.ToString(val);
                           //ar[x, fake_y - 1] += fake_x + Convert.ToString(val) + skip_n[x] + fake_x;
                           //ar[x, fake_y - 1] += fake_x + Convert.ToString(val) + skip_n[x] + fake_x;
 
                       }
                       else
                           //非表示スイッチが1の場合
                       {
                           Range index = ws1.Cells[y, x+1];
                           dynamic val = index.Value2; 
                           //ar[x, fake_y - 1] += fake_x+Convert.ToString(val) + skip_n[x];
                           ar[fake_x, fake_y - 1] += Convert.ToString(val);
                           fake_x++; //スキップ列番号を増やす
                       }
                    }
                }

                
            }
            wb.Close(false); //ブッククローズ
            ExcelApp.Quit(); //Excel終了


      //データグリッドビューに代入



            int rim_y = fake_y + 0;//読み込み位置限度の設定

            label1.Text = nx.maxfy.ToString()+"件読み込み、有効"+rim_y+"件";





            dataGridView1.ColumnCount = nx.maxx+3; //gridviewの行・列設定
            dataGridView1.RowCount = fake_y;



            //DataGridView1の行ヘッダーに行番号を表示する
            for (int i = 0; i < nx.maxfx; i++)
            {
                dataGridView1.Columns[i].HeaderCell.Value = col[i, 0];
            }

            //行ヘッダーの幅を自動調節する
            dataGridView1.AutoResizeRowHeadersWidth(
                DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);


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

