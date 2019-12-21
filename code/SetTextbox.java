import java.io.*;
import org.apache.poi.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * シートにテキストボックスを貼り付ける
 */ 
public class SetTextbox {

  // Patriarchオブジェクト 2003の場合のみ
  protected HSSFPatriarch _patr2003 = null;
  // Drawingオブジェクト 2007の場合のみ
  protected XSSFDrawing _patr2007 = null;

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // ワークシート生成
    Sheet sheet = workBook.createSheet("Sheet1");
    // テキストボックスを作る
    if (mode.equals("2003")) {
      _patr2003 = ((HSSFSheet)sheet).createDrawingPatriarch();
      HSSFClientAnchor anchor = new HSSFClientAnchor(
            0, 0, 0, 0,
            (short)1, 1, (short)8, 6);
      anchor.setAnchorType(0); // Cellに併せて移動・リサイズ
      // Textbox作成
      HSSFTextbox box = _patr2003.createTextbox(anchor);
      // Textboxに書式設定
      // 水平中央揃え
      box.setHorizontalAlignment(
          HSSFTextbox.HORIZONTAL_ALIGNMENT_CENTERED); 
      // 垂直中央揃え
      box.setVerticalAlignment(
          HSSFTextbox.VERTICAL_ALIGNMENT_CENTER);
      // テキストボックスに設定するHSSFRichTextStringインスタンス生成
      HSSFRichTextString rst = 
              new HSSFRichTextString("Apache POI");
      // Fontを指定
      Font fnt = workBook.createFont();
      fnt.setFontName("ＭＳ 明朝");
      fnt.setFontHeightInPoints((short)48);
      fnt.setColor((short)HSSFColor.BLUE.index);
      // FontをHSSFRichTextStringに適用
      rst.applyFont(fnt);
      // TextboxにHSSFRichTextStringを指定
      box.setString(rst);
    }
    else {
      // XSSFはPOI API未実装である。
    }
   // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? this.getClass().getName() + "_Book1.xls" : 
                      this.getClass().getName() + "_Book1.xlsx");
      workBook.write(out);
    }catch(IOException e){
      System.out.println(e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println(e.toString());
      }
    }
    System.out.println("done!");
  }
  /** エントリーポイント */

  public static void main(String[] args) {

    if (args.length != 1) {
      System.out.println("エラー：モードを指定して下さい。");
      return;
    }
    else if ( !args[0].equals("2003") ) {
      System.out.println("エラー：モードは今のところ2003のみ指定して下さい。");
      return;
    }
    // 処理の実行
    new SetTextbox().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
