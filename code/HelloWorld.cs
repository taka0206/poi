using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace HelloNoPOI
{
    /// <summary>
    /// C# - COM連携のサンプル
    /// </summary>
    class HelloWorld
    {
        static void Main(string[] args)
        {
            // Excelオブジェクトの生成
            Excel.Application xls = new Excel.Application();
            // ワークブックの生成
            Excel.Workbook workBook = xls.Workbooks.Add();
            // シートの生成
            Excel.Worksheet sheet = workBook.Sheets.Add();
            sheet.Name = "HelloWorld";
            // 文字列を設定するRangeを取得
            Excel.Range aRange = sheet.get_Range("A1");
            aRange.Font.Name = "ＭＳ 明朝";
            aRange.Font.Size = 48;
            aRange.Font.Color = System.Drawing.Color.Aqua;
            // Rangeに値設定
            aRange.Value = "Hello World On C# Only♪";
            // ワークブックの書き出し
            workBook.SaveAs(@".\HelloWorld-CSharp.xls");
            workBook.Close(true);
            xls.Quit();
            Console.WriteLine("done!");
        }
    }
}
