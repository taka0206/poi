import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * 親のRowのそのまた親のシート取得テスト
 */
public class GetParentSheetByCellTest {

  /**
   * Cell固有の処理
   *@param cell セルの参照
   */
  public void cellProc(Cell cell) {
    Sheet pSheet = cell.getSheet();
    System.out.println("親のSheetは" + 
                pSheet.getSheetName() + "です。");
    System.out.println(pSheet.getSheetName() + 
                "にはRowが" + 
                pSheet.getLastRowNum() + "行あり、");
    System.out.println("私は" + 
                cell.getRowIndex() +
                "行目のRowの" + 
                cell.getColumnIndex() + 
                "個めのセルです。");
  }
  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    Random rand = new Random(); // 乱数発生の準備
    int lim;
    // シートを5枚、各シートにRowを最大10行Cellを最大20個生成
    for (int i=0; i<5; i++) {
      Sheet sheet = workBook.createSheet();
      lim = rand.nextInt(10) + 1;
      for (int j=0; j<lim; j++) {
        Row row = sheet.createRow(j);
        lim = rand.nextInt(20) + 1;
        for (int k=0; k<lim; k++) {
          row.createCell(k).setCellValue(i + "-" + j + "-" + k);
        }
      }
    }
    // 確認用にWorkbookを出力
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
        return;
      }
    }
    // 処理するCellを一つ決定する。
    Sheet curSheet = workBook.getSheetAt(rand.nextInt(5));
    if (curSheet == null) {
      System.out.println("Sheet取得失敗");
      return;
    }
    Row curRow = curSheet.getRow(rand.nextInt(curSheet.getLastRowNum()));
    if (curRow == null) {
      System.out.println("Row取得失敗");
      return;
    }
    Cell curCell = curRow.getCell(rand.nextInt(curRow.getLastCellNum()));
    // Cell固有処理呼び出し
    cellProc(curCell);
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
    new GetParentSheetByCellTest().Run(args[0]);
    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
