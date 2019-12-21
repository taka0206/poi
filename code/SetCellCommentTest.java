import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;
/**
 * Cellコメン設定テスト
 */
public class SetCellCommentTest {

  // Patriarchオブジェクト 2003の場合のみ
  protected HSSFPatriarch _patr2003 = null;
  // Drawingオブジェクト 2007の場合のみ
  protected XSSFDrawing _patr2007 = null;

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
    // コメントの生成と貼り付け
    if (mode.equals("2003")) {
      _patr2003 = 
        ((HSSFSheet)sheet).createDrawingPatriarch();
      HSSFClientAnchor anchor = new HSSFClientAnchor(
        0, 0, 0, 0, (short)6, 4, (short)8, 9);
      anchor.setAnchorType(0); // Cellに併せて移動・リサイズ
      // コメントの生成
      HSSFComment cmt = _patr2003.createComment(anchor);
      // コメントに文字設定
      HSSFRichTextString rt = 
        new HSSFRichTextString("コメント");
      Font fnt = workBook.createFont();
      fnt.setFontName("ＭＳ Ｐゴシック");
      fnt.setFontHeightInPoints((short)14);
      fnt.setColor((short)HSSFColor.RED.index);
      fnt.setItalic(true);
      fnt.setBoldweight(Font.BOLDWEIGHT_BOLD);
      rt.applyFont(fnt);
      cmt.setString(rt); 
      cmt.setAuthor(new String("丸岡 孝司"));
      Cell cell = sheet.createRow(5).createCell(5);
      cell.setCellComment(cmt);
    }
    else {
      _patr2007 = 
        ((XSSFSheet)sheet).createDrawingPatriarch();
      XSSFClientAnchor anchor = new XSSFClientAnchor(
        0, 0, 0, 0, (short)6, 4, (short)8, 9);
      anchor.setAnchorType(0); // Cellに併せて移動・リサイズ
      // コメントの生成
      XSSFComment cmt = 
        _patr2007.createCellComment(anchor);
      // コメントに文字設定
      XSSFRichTextString rt = 
        new XSSFRichTextString("コメント");
      Font fnt = workBook.createFont();
      fnt.setFontName("ＭＳ Ｐゴシック");
      fnt.setFontHeightInPoints((short)14);
      fnt.setColor((short)HSSFColor.RED.index);
      fnt.setItalic(true);
      fnt.setBoldweight((short)10);
      rt.applyFont(fnt);
      cmt.setString(rt); 
      cmt.setAuthor(new String("丸岡 孝司"));
      Cell cell = sheet.createRow(5).createCell(5);
      cell.setCellComment(cmt);
    }
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
    new SetCellCommentTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
