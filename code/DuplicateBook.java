import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 * ワークブックインスタンスの複製
 */
class DuplicateBook {
  /** 処理の実行
   * @param モード
   */
  public void Run(String mode) {
    FileInputStream fisA = null;
    FileInputStream fisB = null;
    Workbook workBookA = null;
    Workbook workBookB = null;
    // ワークブックを読み込む
    try {
      fisA = new FileInputStream( 
          mode.equals("2003") ? 
            "./input/SampleLauncherORG.xls" : 
            "./input/SampleLauncherORG.xlsm");
      workBookA = mode.equals("2003") ? 
            new HSSFWorkbook(fisA) : 
            new XSSFWorkbook(fisA);
      fisA.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + 
                        e.toString());
      return;
    }
    // 名前を変更してワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(
          mode.equals("2003") ? 
          this.getClass().getName() + "_Book1.xls" : 
          this.getClass().getName() + "_Book1.xlsm");
      workBookA.write(out);
    }catch(IOException e){
      System.out.println(e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println(e.toString());
      }
    }
    // 複製されたワークブックを読み込む
    try {
      fisB = new FileInputStream(
          mode.equals("2003") ? 
          this.getClass().getName() + "_Book1.xls" : 
          this.getClass().getName() + "_Book1.xlsm");
      workBookB = mode.equals("2003") ?
            new HSSFWorkbook(fisB) : 
            new XSSFWorkbook(fisB);
      fisB.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" +
                        e.toString());
      return;
    }
    System.out.println("ワークブックの複製が完了しました。");
  }
  /** エントリーポイント */
  public static void main(String[] args) {
    if (args.length != 1) {
      System.out.println("エラー：モードを指定してください。");
      return;
    }
    else if ( !args[0].equals("2003") &&
              !args[0].equals("2007") ) {
      System.out.println(
      "エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new DuplicateBook().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
