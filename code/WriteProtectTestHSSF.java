import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * 書込み保護設定のテスト
 */
public class WriteProtectTestHSSF {

  /** 
   * 処理の実行
   */
  public void Run() {
    // ワークブックの生成
    HSSFWorkbook workBook = new HSSFWorkbook();
    // シート生成 
    HSSFSheet sheet = workBook.createSheet();
    // 行を10行セルを10個作成して値設定
    for (int i=0; i<10; i++) {
      HSSFRow row = sheet.createRow(i);
      for (int j=0; j<10; j++) {
        row.createCell(j).setCellValue(i+"-"+j);
      }
    }
		// 書込み保護を設定
		workBook.writeProtectWorkbook("POI", "POIAPI");
		// ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( "./Book1.xls");
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

    new WriteProtectTestHSSF().Run();
  }
}
