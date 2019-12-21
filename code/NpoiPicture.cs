using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util.IO;

namespace NPOIPicture
{
    /// <summary>
    /// シートに画像を貼り付ける。
    /// C# + NPOIバージョン
    /// </summary>
    class NpoiPicture
    {
        // Patriarchオブジェクト 2003の場合のみ
        protected HSSFPatriarch _patr2003 = null;

        /// <summary>
        /// 処理の実行
        /// </summary>
        public void Run()
        {
            // ワークブックの生成
            Workbook workBook = new HSSFWorkbook(); 
 
            // ワークシート生成
            Sheet sheet = workBook.CreateSheet("Sheet1");
            // 画像ファイルを読み込む
            byte[] bytes;
            try
            {
                using (FileStream ios = new FileStream(@"c:\poi\work\npoi-logo.jpg", System.IO.FileMode.Open))
                {
                    bytes = IOUtils.ToByteArray(ios);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("画像ファイル読込エラー\n" + e.ToString());
                return;
            }
            int picIdx = workBook.AddPicture(bytes, PictureType.JPEG);
            ClientAnchor anchor;
            // 画像の貼り付け
            _patr2003 = (HSSFPatriarch)((HSSFSheet)sheet).CreateDrawingPatriarch();
            anchor = new HSSFClientAnchor(0, 0, 0, 0, (short)1, 1, (short)4, 13);
            anchor.AnchorType = 0; // Cellに併せて移動・リサイズ
            // 画像の貼り付け
            _patr2003.CreatePicture(anchor, picIdx);

            // ワークブック書き出し
            try
            {
                using (FileStream fot = new FileStream(@"c:\poi\work\Book1.xls", FileMode.OpenOrCreate))
                {
                    workBook.Write(fot);
                }
            }
            catch(IOException e)
            {
                Console.WriteLine("ワークブック書き出しに失敗しました。\n" + e.ToString());
            }
            Console.Write("Done!");
        }
        /// <summary>
        /// エントリーポイント
        /// </summary>
        /// <param name="args"></param>
        static public void Main(string[] args)
        {
            new NpoiPicture().Run();
        }

    }
}
