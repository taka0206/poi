import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * Cell移動のテスト
 */
public class MoveCellTest {
  /**
   * Cellの状態表示
   * @param row 処理対象のRow
   */
  protected void PrintCellState(Row row)
  {
    for (int i=0; i<10; i++) {
      Cell cell = row.getCell(i);
      if (cell != null) {
        System.out.print("○");
      }
      else {
        System.out.print("×");
      }
    }
    System.out.println("");
  }
  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // シートの生成 
    Sheet sheet = workBook.createSheet();
    // Rowを生成
    Row row = sheet.createRow(0);
    // Cellを10個生成
    for (int i=0; i<10; i++) {
      row.createCell(i).setCellValue(i);
    }
    // 真ん中あたりのCellを削除
    row.removeCell(row.getCell(4));
    // Cellの状態を出力
    PrintCellState(row);
    // Cellを移動
    for (int i=5; i<10; i++) {
      if (mode.equals("2003")) {
        HSSFCell hCell = (HSSFCell)row.getCell(i);
        ((HSSFRow)row).moveCell(hCell, (short)(i-1));
      }
    }
    // Cellの状態を出力
    PrintCellState(row);
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
    new MoveCellTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
