import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 

/**
 * 疑わしきを検証するクラス
 * 生成しなかった行やセルはどうなっているのか。
 */
class UtagawaTest {
  protected String _mode;

  /**
   * ブックの書き出し処理
   */
  protected void writeBook() {
    // ワークブックの生成
    Workbook workBook = _mode.equals("2003") ? new HSSFWorkbook() : new XSSFWorkbook();
    // ワークシートを生成
    Sheet sheet = workBook.createSheet("Sheet1");
    // Rowを10、Cellを20生成
    for (int i=0; i<10; i++) {
      Row row = sheet.createRow(i);
      for (int j=0; j<20; j++) {
        row.createCell(j);
      }
    }
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( _mode.equals("2003") ? this.getClass().getName() + "_Book1.xls" : 
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
    System.out.println("WriteBook done!");
  }
  /**
   * ブックの読み込み
   */
  protected void readBook() {
    FileInputStream fis = null;
    Workbook workBook = null;
    // ワークブックの読み込み
    try {
      fis = new FileInputStream( _mode.equals("2003") ? 
              this.getClass().getName() + "_Book1.xls" : this.getClass().getName() + "_Book1.xlsx");
      workBook = _mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println(e.toString());
    }
    Sheet sheet = workBook.getSheetAt(0);
    // 任意のセルに値設定
    // もし行とセルが存在しなければここで落ちるはずである。
    try {
      sheet.getRow(1).getCell(0).setCellValue("セルはあるか");
    }
    catch (Exception e) {
      System.out.println("やっぱりセルはなかった！！" + e.toString());
      return;
    }
    System.out.println("セルは存在しました。");
  }
  /** 処理の実行
   * @param mode モード
   * @param rw   書き出し、読み込み
   */
  public void Run(String mode, String rw) {

    _mode = mode;

    if (rw.equals("w")) {
      writeBook();
    }
    else {
      readBook();
    }
  }
  /** エントリーポイント */
  public static void main(String[] args) {
    if (args.length == 0) {
      System.out.println("エラー：使い方-> CalcTest モード rwフラグ");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    String inputValue;
    if (args.length == 1) {
      while (true) {
        System.out.print("Read(r)/write(w)のいずれかを指定してください。中止(X) ->");
        BufferedReader buf =
                new BufferedReader(
                       new InputStreamReader(System.in),1);
        try {
          inputValue = buf.readLine().toLowerCase();
        }
        catch (Exception e)
        {
          System.out.println("rwフラグ入力でエラーが発生しました。" + e.toString());
          return;
        }
        if (inputValue.equals("x")) {
          return;
        } 
        if ( !inputValue.equals("r") && !inputValue.equals("w") ) {
          System.out.println("エラー：rwフラグはrまたはwを指定して下さい。");
        }
        else {
          break;
        }
      }
    }
    else {
      if ( !args[1].equals("r") && !args[1].equals("w") ) {
        System.out.println("エラー：rwフラグはrまたはwを指定して下さい。");
        return;
      }
      inputValue = args[1];
    }
    new UtagawaTest().Run(args[0], inputValue);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
