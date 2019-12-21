import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * 印刷時余白設定テスト
 */ 
public class SetMarginTest {
  /** 
   * cm - インチ変換
   * @param cm 長さ(センチメートル)
   */
  protected double getInch(double cm) {
    return cm * 0.3937;
  }
  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
 
    // ワークシート生成
    Sheet sheet = workBook.createSheet();
    // 印刷時の余白を設定
    // 上下1.5cm
    sheet.setMargin(Sheet.TopMargin, getInch(1.5));
    sheet.setMargin(Sheet.BottomMargin, getInch(1.5));
    // 左右2cm
    sheet.setMargin(Sheet.LeftMargin, getInch(2.0));
    sheet.setMargin(Sheet.RightMargin, getInch(2.0));
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./" + this.getClass().getName() + "_Book1.xls" : 
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
      System.out.println("エラー：モードを指定して下さい。");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new SetMarginTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
