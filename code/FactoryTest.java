import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

/**
 * WorkbookFactoryのテスト
 */
class FactoryTest {
  /** 処理の実行
   * @param fName 読み込むファイル名
   */
  public void Run(String fName) {
    // ワークブックを読み込む
    Workbook workBook = null;
    try {
      workBook = WorkbookFactory.create(
          new FileInputStream(fName));
    }
    catch(Exception e) {
      System.out.println(e.toString());
    }
    // Excelドキュメント形式を判定する。
    if (workBook instanceof HSSFWorkbook) {
      System.out.println("Excel2003以前の形式です。");
    }
    else if(workBook instanceof XSSFWorkbook) {
      System.out.println("Excel2007以降の形式です。");
    }
    else {
      System.out.println("不明な形式です。");
    }
  }
  /** エントリーポイント */
  public static void main(String[] args) {
    String inputValue;
    // 読み込み対象ファイル入力
    while (true) {
      System.out.print("読み込むExcelファイル名を入力してください(フルパス)。中止(x) ->");
      BufferedReader buf =
              new BufferedReader(
                     new InputStreamReader(System.in),1);
      try {
        inputValue = buf.readLine().toLowerCase();
      }
      catch (Exception e)
      {
        System.out.println("ファイル名入力でエラーが発生しました。" + e.toString());
        return;
      }
      if (inputValue.equals("x")) {
        return;
      }
      if (inputValue.length() != 0) {
        break;
      }
    }
    // 処理の実行
    new FactoryTest().Run(inputValue);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
