import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

/**
 * 標準行高さ設定のテスト
 */
public class SetDefRowHeightTest {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // シートの生成 
    Sheet sheet1 = workBook.createSheet();
    Sheet sheet2 = workBook.createSheet();
		// Rowを10行生成
		for (int i=0; i<10; i++) {
			Row row1 = sheet1.createRow(i);
			row1.createCell(0).setCellValue(i);
			row1.setHeight((short)1000);
			Row row2 = sheet2.createRow(i);
			row2.createCell(0).setCellValue(i);
			row2.setHeight((short)2000);
		}
		// Sheet1の標準行高さ設定
//		sheet1.setDefaultRowHeight((short)400);
		// sheet2の標準行高さ設定
//		sheet2.setDefaultRowHeightInPoints((float)40.0);
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./Book1.xls" : 
                      "./Book1.xlsx");
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
    new SetDefRowHeightTest().Run(args[0]);
  }
}
