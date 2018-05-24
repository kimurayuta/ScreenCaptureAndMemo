using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;

// 参照の追加 -> COM -> Microsoft Office 14.0 Object Library
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace MemoToExcel
{
    class Program
    {
        // 実行方法
        // コマンドライン実行
        // 画像とメモが保存されているディレクトリを渡すと、エクセルにまとめます。
        // C:\Users\yu-kimura\MemoToExcel.exe C:\Users\yu-kimura\Documents\xxx
        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            //エクセルを非表示
            ExcelApp.Visible = false;

            //エクセルファイルのオープン
            Microsoft.Office.Interop.Excel.Workbook WorkBook = ExcelApp.Workbooks.Add();

            //1シート目の選択
            Microsoft.Office.Interop.Excel.Worksheet sheet = WorkBook.Worksheets[1];

            string[] pngFiles = Directory.GetFiles(args[0], "*.png");
            int i = 0;
            //string[,] data = new string[pngFiles.Length, 1];

            int imageMaxW = 640, imageMaxH = 480, shapeMaxW = 320, interval = imageMaxH + 10, imageX = shapeMaxW + 20;
            foreach (string png in pngFiles)
            {
                Bitmap bmpSrc = new Bitmap(png);
                //元画像の縦横サイズを調べる
                int width = bmpSrc.Width;
                int height = bmpSrc.Height;

                if (width < imageMaxW || height < imageMaxH)
                {
                    sheet.Shapes.AddPicture(png, MsoTriState.msoFalse, MsoTriState.msoTrue, imageX, interval * i, width, height);
                }
                else
                {
                    sheet.Shapes.AddPicture(png, MsoTriState.msoFalse, MsoTriState.msoTrue, imageX, interval * i, imageMaxW, imageMaxH);
                }
                Microsoft.Office.Interop.Excel.Shape s = sheet.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, interval * i, shapeMaxW, imageMaxH);
                string memo = File.ReadAllText(png + ".txt", Encoding.Unicode);

                s.TextFrame.Characters(0, 0).Text = memo;

                if (memo.StartsWith("*"))
                {
                    s.TextFrame.Characters(0, memo.Length).Font.Color = Color.Red;
                }

                i++;
                //data[i++, 0] = File.ReadAllText(png + ".txt", Encoding.Unicode);
            }
            //Microsoft.Office.Interop.Excel.Range range = sheet.Range[sheet.Cells[1,1], sheet.Cells[i,1]];
            // range.Value2 = data;

            //workbookを閉じる
            WorkBook.SaveAs(args[0] + ".xlsx");
            WorkBook.Close();
            //エクセルを閉じる
            ExcelApp.Quit();
        }
    }
}

// https://msdn.microsoft.com/ja-jp/library/07wt70x2.aspx
// https://msdn.microsoft.com/ja-jp/library/6yk7a1b0.aspx
// http://wannabe-note.com/1160
// http://www.ipentec.com/document/document.aspx?page=csharp-save-excel-new-file-and-write
// http://migelnanai.blog.so-net.ne.jp/2007-04-09
// http://ameblo.jp/okya-tec/entry-10622959020.html
// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.textframe.characters.aspx
// https://msdn.microsoft.com/ja-jp/library/microsoft.office.interop.excel.shape_members%28v=office.11%29.aspx
// http://www.moug.net/tech/exvba/0120020.html
// https://msdn.microsoft.com/ja-jp/library/microsoft.office.interop.excel.shapes.addpicture%28v=office.11%29.aspx
// https://social.msdn.microsoft.com/Forums/vstudio/ja-JP/c68f4658-1537-4859-a05d-fbb46ddf28f3?forum=csharpgeneralja
