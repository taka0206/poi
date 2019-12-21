import java.io.*;
import org.apache.poi.ss.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.*;
/**
 * バッチファイル作成2
 */
class LancheBat2 {

  protected String _mode;   // 動作モード
  protected Workbook _workBook; // ランチャーワークブックのインスタンス
  protected int _classPos;    // クラス名の桁位置
  protected String[] _breakKeys; // キーブレーク見出退避領域
  protected int[] _breakPos; // キーブレーク行番号退避領域
  /** 
   * 初期処理
   */
  protected boolean init() {
    FileInputStream fis = null;
    // ワークブックを読み込む
    _workBook = null;
    try {
      fis = new FileInputStream("./input/SampleLauncherORG.xls");
      _workBook = new HSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + e.toString());
      return false;
    }
    return true;
  }
  /** 
   * バッチファイル作成処理
   */
  protected void buildBat() {
    // データシートの取得
    Sheet dSheet = _workBook.getSheet("データシート");
    // 2行目から最終行まで処理
    for (int i=2; i<dSheet.getLastRowNum(); i++) {
      Row row = dSheet.getRow(i);
      String className = row.getCell(4).getStringCellValue(); // クラス名
      if (className.equals("なし") == false) {
        if (row.getCell(8).getBooleanCellValue() == false) {
          // 特殊入力のやつは処理しない。
          if (row.getCell(7).getBooleanCellValue() == false) {
            // Bookを生成するやつは処理しない。
            if (row.getCell(5).getBooleanCellValue() == false) {
              // ビルドコマンド出力
              System.out.println("javac " + className + ".java");
              // 実行コマンド出力(2003)
              System.out.println("java " + className + " 2003");
              if (row.getCell(6).getBooleanCellValue() == true) {
                // 実行コマンド出力(2007)
                System.out.println("java " + className + " 2007");
              }
            }
          }
        }
      }
    }
  }
  /** 処理の実行
   * @param モード
   */
  public void Run() {

    // 初期処理
    if (init() == false) {
      return;
    }
    // バッチファイル作成
    buildBat();
  }
  /** エントリーポイント */
  public static void main(String[] args) {

    new LancheBat2().Run();
  }
}
