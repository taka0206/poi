import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * 強制改ページ解除テスト
 */ 
public class RemovePageBreakTest {

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

    // Rowを30行Cellを15列生成
    for (int i=0; i<30; i++) {
      Row row = sheet.createRow(i);
      for (int j=0; j<15; j++) {
        row.createCell(j).setCellValue(i + "-" + j);
      }
    }
    // 改ページ位置を10行、20行と5列、10列に設定
    sheet.setRowBreak(9);
    sheet.setRowBreak(19);
    sheet.setColumnBreak(4);
    sheet.setColumnBreak(9);
    // 改ページ位置(行)をすべて解除
    for(int breakLine : sheet.getRowBreaks()) {
      sheet.removeRowBreak(breakLine);
    }
    // 改ページ位置(列)をすべて削除
    for(int breakCol : sheet.getColumnBreaks()) {
      sheet.removeColumnBreak(breakCol);
    }
    // 全行解除してみる - IllegalArgumentExceptionが発生
    /*
    for (int i=0; i<=sheet.getLastRowNum(); i++) {
      sheet.removeRowBreak(i);
    }
    */
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
    new RemovePageBreakTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
