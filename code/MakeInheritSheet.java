import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * 継承図シート作成クラス
 */
class MakeInheritSheet {
  HSSFWorkbook _workBook = null;
  HSSFSheet _sheet = null;
  HSSFPatriarch _patr = null;

  /**
   * 修飾済み文字列取得
   * @param text 設定する文字列
   */
  protected HSSFRichTextString getContent(
                              String text) {
    HSSFRichTextString st = 
              new HSSFRichTextString(text);
    HSSFFont fnt = _workBook.createFont();
    fnt.setFontName("ＭＳ Ｐゴシック");
    fnt.setFontHeightInPoints((short)12);
    st.applyFont(fnt);
    return st;
  }
  /** 処理の実行 */
  public void Run() {
    // ワークブックの生成
    _workBook = new HSSFWorkbook();
    // ワークシートの生成
    _sheet = _workBook.createSheet(
              "SSインターフェース継承図");
    _patr = _sheet.createDrawingPatriarch();
    // テキストボックスの生成
    HSSFTextbox box1 = _patr.createTextbox(
          new HSSFClientAnchor(0, 0, 0, 0, 
              (short)3, 3, (short) 8, 6));
    box1.setString(getContent(
      "org.apache.poi.ss.usermodel.Workbook" +
      "\nインターフェース"));
    box1.setVerticalAlignment(
      HSSFTextbox.HORIZONTAL_ALIGNMENT_CENTERED);
    box1.setHorizontalAlignment(
      HSSFTextbox.VERTICAL_ALIGNMENT_CENTER);

    HSSFTextbox box2 = _patr.createTextbox(
                new HSSFClientAnchor(0, 0, 0, 0, 
                (short)1, 10, (short) 5, 13));
    box2.setString(getContent(
      "org.apache.poi.hssf.usermodel." +
      "\nHSSFWorkbookクラス"));
    box2.setVerticalAlignment(
      HSSFTextbox.HORIZONTAL_ALIGNMENT_CENTERED);
    box2.setHorizontalAlignment(
      HSSFTextbox.VERTICAL_ALIGNMENT_CENTER);

    HSSFTextbox box3 = _patr.createTextbox(
                new HSSFClientAnchor(0, 0, 0, 0, 
                (short)6, 10, (short) 10, 13));
    box3.setString(getContent(
      "org.apache.poi.xssf.usermodel.\n" +
      "XSSFWorkbookクラス"));
    box3.setVerticalAlignment(
      HSSFTextbox.HORIZONTAL_ALIGNMENT_CENTERED);
    box3.setHorizontalAlignment(
      HSSFTextbox.VERTICAL_ALIGNMENT_CENTER);
    // ラインの生成
    HSSFSimpleShape shape1 = _patr.createSimpleShape(
                new HSSFClientAnchor(0, 0, 0, 0,
                (short)5, 6, (short)3,10));
    shape1.setShapeType(HSSFSimpleShape.OBJECT_TYPE_LINE);
    shape1.setLineStyle(HSSFShape.LINESTYLE_LONGDASHGEL);
    HSSFSimpleShape shape2 = _patr.createSimpleShape(
                new HSSFClientAnchor(0, 0, 0, 0,
                (short)6, 6, (short)8,10));
    shape2.setShapeType(HSSFSimpleShape.OBJECT_TYPE_LINE);
    shape2.setLineStyle(HSSFShape.LINESTYLE_LONGDASHGEL);
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(
            this.getClass().getName() + "_Book1.xls");
      _workBook.write(out);
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
  /**
   * エントリーポイント
   */
  public static void main(String args[]) {
    new MakeInheritSheet().Run();
    System.out.print("リターンキーで終了……");

    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
