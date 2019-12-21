using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace helloWorldCSAndJava
{
    /// <summary>
    /// C# ; POIによる”こんにちは。世界”
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            // Webサービス参照の生成
            CsWrapper.CSWrapperClient wrapper = new CsWrapper.CSWrapperClient();
            // ワークブックのオープンとRow、Cellの作成
            wrapper.createWorkSheetAndRowAndCell("2003", "HelloWorld");
            // 文字フォント、ポイント、色の設定
            wrapper.setFontAndStyle("ＭＳ 明朝", 48, 49);
            // Cellに値設定
            wrapper.setCellValue("Hello World On C#♪");
            // ワークブック書き出し
            wrapper.write("./HelloWorld-CSharp.xls");

            Console.WriteLine("Done!");
        }
    }
}
