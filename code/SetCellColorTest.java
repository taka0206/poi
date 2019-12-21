import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * 前景色・背景色・パターンのテスト
 */
public class SetCellColorTest {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {

    // 塗り潰しパターン名
    String[] fillPatNames = {
                             "NO_FILL"
                            ,"SOLID_FOREGROUND"
                            ,"FINE_DOTS"
                            ,"ALT_BARS"
                            ,"SPARSE_DOTS"
                            ,"THICK_HORZ_BANDS"
                            ,"THICK_VERT_BANDS"
                            ,"THICK_BACKWARD_DIAG"
                            ,"THICK_FORWARD_DIAG"
                            ,"BIG_SPOTS"
                            ,"BRICKS"
                            ,"THIN_HORZ_BANDS"
                            ,"THIN_VERT_BANDS"
                            ,"THIN_BACKWARD_DIAG"
                            ,"THIN_FORWARD_DIAG"
                            ,"SQUARES"
                            ,"DIAMONDS"
                            ,"LESS_DOTS"
                            ,"LEAST_DOTS"
    };
    // 塗り潰しパターン値
    short[] fillPatValues = {
                             CellStyle.NO_FILL
                            ,CellStyle.SOLID_FOREGROUND
                            ,CellStyle.FINE_DOTS
                            ,CellStyle.ALT_BARS
                            ,CellStyle.SPARSE_DOTS
                            ,CellStyle.THICK_HORZ_BANDS
                            ,CellStyle.THICK_VERT_BANDS
                            ,CellStyle.THICK_BACKWARD_DIAG
                            ,CellStyle.THICK_FORWARD_DIAG
                            ,CellStyle.BIG_SPOTS
                            ,CellStyle.BRICKS
                            ,CellStyle.THIN_HORZ_BANDS
                            ,CellStyle.THIN_VERT_BANDS
                            ,CellStyle.THIN_BACKWARD_DIAG
                            ,CellStyle.THIN_FORWARD_DIAG
                            ,CellStyle.SQUARES
                            ,CellStyle.DIAMONDS
                            ,CellStyle.LESS_DOTS
                            ,CellStyle.LEAST_DOTS
    };

    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // シートの生成 
    Sheet sheet = workBook.createSheet("Cell塗り潰しパターン");
    
    for (int i=0; i<19;i++) {
      Row row = sheet.createRow(i);
      row.createCell(0).setCellValue(fillPatNames[i]);
      row.createCell(1).setCellValue("(" + fillPatValues[i] +")");
      Cell cell = row.createCell(2);
      CellStyle style = workBook.createCellStyle();
      style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
      style.setFillBackgroundColor(IndexedColors.WHITE.getIndex());
      style.setFillPattern(fillPatValues[i]);
      cell.setCellStyle(style);
    }
    sheet.autoSizeColumn(0);  // 1列目を自動幅設定に
    sheet.setZoom(2,1);       // 200%に拡大

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
    new SetCellColorTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
