import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * POI HSSF ->Excelドキュメント操作ライブラリ
 * クラス構造概要解説ワークシートシート作成
 */
class MakeSummarySheet {
  
  HSSFWorkbook _workBook = null;
  HSSFSheet _sheet = null;
  HSSFPatriarch _patr = null;
  
  /** コンストラクター */
  public MakeSummarySheet() { 
  }
  /**
   * フォント生成サブ
   *@param point 文字のポイント数
   *@param center センタリングするかどうかのフラグ
   */
  protected HSSFCellStyle getCellStyle(short point, 
                          boolean center) {
    try {
      HSSFCellStyle st = _workBook.createCellStyle();
      if (center == true){
        st.setAlignment(HSSFCellStyle.ALIGN_CENTER);
      }
      HSSFFont fnt = _workBook.createFont();
      fnt.setFontName("ＭＳ Ｐゴシック");
      fnt.setFontHeightInPoints(point);
      st.setFont(fnt);
      return st;
    } catch (Exception e) {
      System.out.println(e.toString());
    }
    return null;

  }
  /**
   * コメントオブジェクト生成サブ
   *@param comment コメントに設定したい文字列
   */
  protected HSSFComment getComment(String comment) {
    HSSFComment cmt = 
      _patr.createComment(new HSSFClientAnchor(
                0, 0, 0, 0, 
                (short)1, 1, (short) 8, 6));
    HSSFRichTextString rt = 
              new HSSFRichTextString(comment);
    HSSFFont fnt = _workBook.createFont();
    fnt.setFontName("ＭＳ Ｐゴシック");
    fnt.setFontHeightInPoints((short)14);
    fnt.setColor((short)HSSFColor.BLUE.index);
    fnt.setItalic(true);
    fnt.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
    rt.applyFont(fnt);
    cmt.setString(rt); 
    cmt.setAuthor(new String("丸岡 孝司"));
    return cmt;
  }
  /** 一括罫線描画処理
   *@param stRow 開始Row
   *@param edRow 終了Row
   *@param stCell 開始Cell
   *@param edCell 終了Cell
   */
  protected void drawLines(int stRow, int edRow, 
                            int stCell, int edCell) {
    // 上罫線は開始行のセルのみ
    HSSFRegionUtil.setBorderTop(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT ,
        new Region(stRow, (short)stCell, 
                    stRow, (short)edCell),
        _sheet, _workBook);
    // 下罫線は終了行のセルのみ
    HSSFRegionUtil.setBorderBottom(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT ,
        new Region(edRow, (short)stCell,
                  edRow, (short)edCell),
        _sheet, _workBook);
    // 左罫線は各行の開始セルのみ
    HSSFRegionUtil.setBorderLeft(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT,
        new Region(stRow, (short)stCell, 
                  edRow, (short)stCell),
        _sheet, _workBook);
    // 右罫線は各行の終了セルのみ
    HSSFRegionUtil.setBorderRight(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT,
        new Region(stRow, (short)edCell,
                  edRow, (short)edCell), 
        _sheet, _workBook);
  }
  /** 処理の実行 */
  public void Run() {
    // ワークブックの生成
    _workBook = new HSSFWorkbook();
    // ワークシートの生成
    _sheet = _workBook.createSheet("POIライブラリ概要");
    _patr = _sheet.createDrawingPatriarch();

    // Rowの一括生成
    for (int i=0;i<25;i++) {
      HSSFRow row = _sheet.createRow(i);
      //cellの一括生成
      for (int j=0;j<20;j++) {
        row.createCell((short)j);
      }
    }
    _sheet.getRow(22).getCell(3).setCellValue(
            "ワークブック全体");
    _sheet.getRow(22).getCell(3).setCellStyle(
            getCellStyle((short)36,false));
    // ワークブックコメント
    HSSFComment cmtBook = getComment("HSSFWorkbook");
    _sheet.getRow(22).getCell(3).setCellComment(
              cmtBook);

    // 大きく罫線を引く
    drawLines(0,19,0,9);

    // シートの文字とコメントの設定
    _sheet.getRow(3).getCell(3).setCellValue(
              "ワークシート");
    _sheet.getRow(3).getCell(3).setCellStyle(
                  getCellStyle((short)24,false));
    HSSFComment cmtSheet = getComment("HSSFSheet");
    _sheet.getRow(3).getCell(3).setCellComment(
              cmtSheet);

    // Rowの文字と罫線描画とコメントの設定
    _sheet.addMergedRegion(
          new Region(8, (short)0, 8, (short)9));
    _sheet.getRow(8).getCell(0).setCellValue(
              "行(Row)");
    _sheet.getRow(8).getCell(0).setCellStyle(
          getCellStyle((short)12,true));
    drawLines(8,8,0,9);
    HSSFComment cmtRow = getComment("HSSFRow");
    _sheet.getRow(8).getCell(0).setCellComment(
          cmtRow);

    // Cellの文字列と罫線描画とコメントの設定
    _sheet.getRow(16).getCell(1).setCellValue(
          "セル(Cell)");
    _sheet.getRow(16).getCell(1).setCellStyle(
                getCellStyle((short)9,false));
    // 任意のセルに罫線を引く
    drawLines(16,16,1,1);
    // Cellコメント
    HSSFComment cmtCell = getComment("HSSFCell");
    _sheet.getRow(16).getCell(1).setCellComment(
                cmtCell);

    // リージョンの文字列と罫線描画とコメントの設定
    _sheet.addMergedRegion(
      new Region(16, (short)5, 17, (short)7));
    _sheet.getRow(16).getCell(5).setCellValue(
          "マージドリージョン(MergedRegion)");
    _sheet.getRow(16).getCell(5).setCellStyle(
          getCellStyle((short)9,true));
    drawLines(16,17,5,7);
    // リージョンコメント
    HSSFComment cmtRegion = 
            getComment("マージされたHSSFCell");
    _sheet.getRow(16).getCell(5).setCellComment(
        cmtRegion);

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
  public static void main(String[] args){
    new MakeSummarySheet().Run();

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
