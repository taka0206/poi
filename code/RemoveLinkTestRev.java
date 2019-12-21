import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * ハイパーリンク解除のテスト(リベンジ版)
 */
public class RemoveLinkTestRev {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックを読み込む
		FileInputStream fis = null;
		Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./poilink.xls" : "./poilink.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + e.toString());
      return;
    }
    // シートの取得
    Sheet sheet = workBook.getSheetAt(0);
		// 2行目から順番に処理
		for (int i=1; i<sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			Cell cellOrg = row.getCell(1);	// Bセルを取得
			// 情報を退避
			String sVal = cellOrg.getStringCellValue();	// 値
			Comment com = cellOrg.getCellComment();			// セルコメント
			CellStyle style = cellOrg.getCellStyle();		// セルスタイル
			int type = cellOrg.getCellType();						// セルタイプ
			// 元のCellを削除
			row.removeCell(cellOrg);
			// 同じ場所にCellを生成
			/*
			Cell cellNew = row.createCell(1);
			// Cellの情報を設定しなおす。
			cellNew.setCellValue(sVal);
			cellNew.setCellComment(com);
			cellNew.setCellStyle(style);
			cellNew.setCellType(type);
			*/
		}
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./hlink.xls" : 
                      "./hlink.xlsx");
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
    new RemoveLinkTestRev().Run(args[0]);
  }
}
