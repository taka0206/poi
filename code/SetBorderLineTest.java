import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

/**
 * 罫線のテスト
 */
public class SetBorderLineTest {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {

    // 線種名
    String[] linePatNames = {
                             "BORDER_NONE"
                            ,"BORDER_THIN"
                            ,"BORDER_MEDIUM"
                            ,"BORDER_DASHED"
                            ,"BORDER_HAIR"
                            ,"BORDER_THICK"
                            ,"BORDER_DOUBLE"
                            ,"BORDER_DOTTED"
                            ,"BORDER_MEDIUM_DASHED"
                            ,"BORDER_DASH_DOT"
                            ,"BORDER_MEDIUM_DASH_DOT"
                            ,"BORDER_DASH_DOT_DOT"
                            ,"BORDER_MEDIUM_DASH_DOT_DOT"
                            ,"BORDER_SLANTED_DASH_DOT"
    };
    // 線種値
    short[] linePatValues = {
                             CellStyle.BORDER_NONE
                            ,CellStyle.BORDER_THIN
                            ,CellStyle.BORDER_MEDIUM
                            ,CellStyle.BORDER_DASHED
                            ,CellStyle.BORDER_HAIR
                            ,CellStyle.BORDER_THICK
                            ,CellStyle.BORDER_DOUBLE
                            ,CellStyle.BORDER_DOTTED
                            ,CellStyle.BORDER_MEDIUM_DASHED
                            ,CellStyle.BORDER_DASH_DOT
                            ,CellStyle.BORDER_MEDIUM_DASH_DOT
                            ,CellStyle.BORDER_DASH_DOT_DOT
                            ,CellStyle.BORDER_MEDIUM_DASH_DOT_DOT
                            ,CellStyle.BORDER_SLANTED_DASH_DOT
    };

    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // シートの生成 
    Sheet sheet = workBook.createSheet("罫線");
    int rowNo = 0;
    for (int i=0; i<14; i++) {
      Row row = sheet.createRow(rowNo);
      Cell cell0 = row.createCell(0);
      cell0.setCellValue(linePatNames[i]);
      Cell cell1 = row.createCell(1);
      cell1.setCellValue("(" + linePatValues[i] + ")");
      // セルを結合する。
      sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo + 1, 0,0));
      sheet.addMergedRegion(new CellRangeAddress(rowNo, rowNo + 1, 1,1));
      // 結合セルのスタイル
      CellStyle styleM = workBook.createCellStyle();
      styleM.setAlignment(CellStyle.ALIGN_RIGHT);
      styleM.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
      cell0.setCellStyle(styleM);
      cell1.setCellStyle(styleM);
      Cell cell = sheet.createRow(rowNo+1).createCell(2);
      // CellStyle生成
      CellStyle styleLine = workBook.createCellStyle();
      // 線種を設定
      styleLine.setBorderTop(linePatValues[i]);
      cell.setCellStyle(styleLine);
      rowNo += 2;
    }
    sheet.autoSizeColumn(0,true);  // 1列目を自動幅設定に(マージ対象)
    sheet.autoSizeColumn(1,true);  // 2列目を自動幅設定に(マージ対象)
    sheet.setColumnWidth(2, 12800); // 3列目を広く
    sheet.setDisplayGridlines(false); // 罫線が見易いようシート枠線を消す。

    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? this.getClass().getName() + "_Book1.xls" : 
                      this.getClass().getName() + "_Book1.xlsx");
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
    // 処理の実行
    new SetBorderLineTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
