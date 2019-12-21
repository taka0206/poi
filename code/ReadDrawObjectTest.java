import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 
import org.apache.poi.hssf.record.crypto.*;
/**
 * 描画オブジェクトつきブックの読み込みテスト
 */
class ReadDrawObjectTest {
  /** 処理の実行
   * @param モード
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? 
          "./input/ReadDrawObject_in.xls" : 
          "./input/ReadDrawObject_in.xlsx");
      workBook = mode.equals("2003") ? 
                  new HSSFWorkbook(fis) : 
                  new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + 
                      e.toString());
      return;
    }
    // シートの取得
    Sheet sheet = workBook.getSheetAt(0);
    if (mode.equals("2003")) {
      // 描画元締めの取得
      HSSFPatriarch _patr2003 =
        ((HSSFSheet)sheet).createDrawingPatriarch();
    }
    else {
      XSSFDrawing _patr2007 = 
        ((XSSFSheet)sheet).createDrawingPatriarch();
    }
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? 
        "./" + this.getClass().getName() + "_Book1.xls" : 
        "./" + this.getClass().getName() + "_Book1.xlsx");
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
      System.out.println("エラー：モードを指定してください。");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007")) {
      System.out.println("エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new ReadDrawObjectTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
    
  }
}
