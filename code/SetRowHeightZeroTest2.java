import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * Row単位の標準セルスタイル設定テスト
 */
public class SetRowHeightZeroTest2 {
  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    Workbook workBook = null;
    // ワークブックの読み込み
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./zero.xls" : "./zero.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println(e.toString());
    }
    // シートの生成 
    Sheet sheet = workBook.getSheetAt(0);
    // 3行目を非表示に設定
    sheet.getRow(2).setZeroHeight(false);
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./Book1.xls" : 
                      "./Book1.xlsx");
      workBook.write(out);
    }catch(IOException e){
      System.out.println("ブックの書き込みに失敗しました。\n" + e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println("ブックの書き込みに失敗しました。\n" + e.toString());
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
    new SetRowHeightZeroTest2().Run(args[0]);
  }
}
