import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

/**
 * アクティブセル、アクティブシート設定クラス
 */
class SetActive {
  /**
   * アクティブセル、アクティブシートの設定処理メイン
   * @param fName ワークブックファイル名
   */
  public void setActiveMain(String fName) {
    // ワークブックを読み込む
    Workbook workBook = null;
    try {
      workBook = WorkbookFactory.create(
                new FileInputStream(fName));
    }
    catch(Exception e) {
      System.out.println(e.toString());
      return;
    }
    // まず全シートの選択状態を解除
    for (int i = 0; 
        i < workBook.getNumberOfSheets(); i++) {
      workBook.getSheetAt(i).setSelected(false);
    }
    // ワークブックに存在するシートを順に処理し、
    // A1セルをアクティブにする。
    for (int i=0;
          i<workBook.getNumberOfSheets(); i++) {
      Sheet sheet = workBook.getSheetAt(i);
      // 1行目のRowを取得
      Row row = sheet.getRow(0);
      if (row == null) {
        // 1行目のRowが存在しない場合。
        Row nrow = sheet.createRow(0);
        Cell cell = nrow.createCell(0);
        // A1セルをアクティブに。
        cell.setAsActiveCell();
      }
      else {
        // 1行目のRowが存在する場合。
        Cell cell = row.getCell(
              0, Row.CREATE_NULL_AS_BLANK);
        // A1セルをアクティブに。
        cell.setAsActiveCell();
      }
    }
    // 最後に第1シートをアクティブに。
    workBook.setActiveSheet(0);
    workBook.getSheetAt(0).setSelected(true);
    // ワークブックを書き出す
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(fName);
      workBook.write(out);
    }catch(IOException e){
      System.out.println(e.toString());
      return;
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println(e.toString());
        return;
      }
    }
    System.out.println(fName + "を処理しました。");
  }
  /** 処理の実行
   * @param path 処理するフォルダ(ディレクトリ)名
   */
  public void Run(String path) {
    File dir = new File(path);
    if (!dir.exists()) {
      System.out.println("指定されたパスは存在しません。");
      return;
    }
    File[] files = dir.listFiles();
    for (int i = 0; i < files.length; i++) {
      File file = files[i];
      if (file.isFile() && file.canRead()) {
        if (file.getPath().toLowerCase().endsWith(".xls") ||
            file.getPath().toLowerCase().endsWith(".xlsx")) {
          setActiveMain(file.getPath());
        }
      }
    }
  }
  /** エントリーポイント */
  public static void main(String[] args) {
    if (args.length != 1) {
      System.out.println(
        "エラー：使い方-> SetActive フォルダ名");
      return;
    }
    new SetActive().Run(args[0]);
  }
}
