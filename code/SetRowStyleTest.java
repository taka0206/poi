import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * Row単位の標準セルスタイル設定テスト
 */
public class SetRowStyleTest {

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
    // セルスタイル生成
    CellStyle style = workBook.createCellStyle();
    // ＭＳ明朝 11ポイントのフォントを生成
    Font fnt = workBook.createFont();
    fnt.setFontName("ＭＳ 明朝");
    fnt.setFontHeightInPoints((short)11);
    // セルスタイルにフォントを設定
    style.setFont(fnt);
    // Rowを5行生成
		for (int i=0; i<5; i++) {
    	sheet.createRow(i);
    }
		// 3行目だけにCellStyle設定
		//if (mode.equals("2003")) ((HSSFRow)sheet.getRow(2)).setRowStyle((HSSFCellStyle)style);
    // 各RowにCellを10個生成し文字を設定
		for (int i=0; i<5; i++) {
    	for (int j=0; j<10; j++) {
      	sheet.getRow(i).createCell(j).setCellValue(i + "-" + j);
			}
		}
		// Row側から標準セルスタイルを取得
		if (mode.equals("2003")) {
			HSSFCellStyle rstyle = ((HSSFRow)sheet.getRow(2)).getRowStyle();
			if (rstyle != null) {
				HSSFFont rfnt = rstyle.getFont((HSSFWorkbook)workBook);
				System.out.println(rfnt.getFontName() + rfnt.getFontHeightInPoints() + "ポイント");
			}
			else {
				System.out.println("標準CellStyleはnullです");
			}
		}

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
    new SetRowStyleTest().Run(args[0]);
  }
}
