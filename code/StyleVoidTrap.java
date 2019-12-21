import java.io.*;
import org.apache.poi.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;
/**
 * セルスタイルの設定 修正バージョン
 */ 
public class StyleVoidTrap {
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
    // セルスタイルを3個生成
    for (int i=0; i<3; i++) {
      workBook.createCellStyle();
    }
    // 作成したセルに書式を設定する。
    CellStyle style1 = workBook.getCellStyleAt((short)21);
    style1.setAlignment(CellStyle.ALIGN_LEFT);
    CellStyle style2 = workBook.getCellStyleAt((short)22);
    style2.setAlignment(CellStyle.ALIGN_CENTER);
    CellStyle style3 = workBook.getCellStyleAt((short)23);
    style3.setAlignment(CellStyle.ALIGN_RIGHT);

    for (int i=1; i<=30; i++) {
      // セルを生成
      Cell cell = row.createCell(i);
      cell.setCellValue("●");
      int mod = i % 3;
      switch (mod) {
        case 1:
          cell.setCellStyle(style1);
          break;
        case 2:
          cell.setCellStyle(style2);
          break;
        default:
          cell.setCellStyle(style3);
          break;
      }
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
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new StyleVoidTrap().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
