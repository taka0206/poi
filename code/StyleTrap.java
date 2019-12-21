import java.io.*;
import org.apache.poi.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * セルスタイルの設定　問題ありバージョン
 */ 
public class StyleTrap {

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
    // Rowを生成
    Row row = sheet.createRow(5);
    // セルスタイル生成
    CellStyle style = workBook.createCellStyle();
    for (int i=1; i<=30; i++) {
      // セルを生成
      Cell cell = row.createCell(i);
      cell.setCellValue("●");
      int mod = i % 3;
      switch (mod) {
        case 1:
          style.setAlignment(CellStyle.ALIGN_LEFT);
          break;
        case 2:
          style.setAlignment(CellStyle.ALIGN_CENTER);
          break;
        default:
          style.setAlignment(CellStyle.ALIGN_RIGHT);
          break;
      }
      cell.setCellStyle(style);
    }
   // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./Book1.xls" : 
                      "./Book1.xlsx");
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
    new StyleTrap().Run(args[0]);
  }
}
