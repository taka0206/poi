import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

/**
 * POI HSSF/XSSF共通 ->Excelドキュメント操作ライブラリ
 * クラス構造概要解説ワークシートシート作成
 */
class MakeSummarySheetSS {
  
  protected String _mode;
  // ワークブックインターフェース
  protected Workbook _workBook = null;
  // シートインターフェース
  protected Sheet _sheet = null;
  // Patriarchオブジェクト 2003の場合のみ
  protected HSSFPatriarch _patr2003 = null;
  // Drawingオブジェクト 2007の場合のみ
  protected XSSFDrawing _patr2007 = null;

  /** 
   * コンストラクター
   *@param mode 動作モード
  */
  public MakeSummarySheetSS(String mode) { 
    _mode = mode;
  }
  /**
   * フォント生成サブ
   *@param point 文字のポイント数
   *@param center センタリングするかどうかのフラグ
   */
  protected void setUserCellStyle(Cell cel, 
                short point, boolean center) {
    try {
      CellStyle st = cel.getCellStyle();
      if (center == true){
        st.setAlignment(CellStyle.ALIGN_CENTER);
      }
      Font fnt = _workBook.createFont();
      fnt.setFontName("ＭＳ Ｐゴシック");
      fnt.setFontHeightInPoints(point);
      st.setFont(fnt);
    } catch (Exception e) {
      System.out.println(e.toString());
    }
  }
  /**
   * コメントオブジェクト生成サブ
   *@param comment コメントに設定したい文字列
   */
  protected Comment getComment(String comment) {

    Comment cmt = null;
    RichTextString rt = null; 

    if (_mode.equals("2003")) {
      cmt = _patr2003.createComment(
              new HSSFClientAnchor(0, 0, 0, 0, 
                (short)1, 1, (short) 8, 6));
      rt = new HSSFRichTextString(comment);
    }
    else {
      cmt = _patr2007.createCellComment(
              new XSSFClientAnchor(0, 0, 0, 0,
                (short)1, 1, (short) 8, 6));
      rt = new XSSFRichTextString(comment);
    }
    Font fnt = _workBook.createFont();
    fnt.setFontName("ＭＳ Ｐゴシック");
    fnt.setFontHeightInPoints((short)14);
    fnt.setColor((short)HSSFColor.BLUE.index);
    fnt.setItalic(true);
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
    for(int i=stCell; i<=edCell;i++) {
      Cell cel = _sheet.getRow(stRow).getCell(i);
      CellStyle styleU = cel.getCellStyle();
      styleU.setBorderTop(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT);
      cel.setCellStyle(styleU);
    }
    // 下罫線は終了行のセルのみ
    for(int i=stCell; i<=edCell;i++) {
      Cell cel = _sheet.getRow(edRow).getCell(i);
      CellStyle styleD = cel.getCellStyle();
      styleD.setBorderBottom(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT);
      cel.setCellStyle(styleD);
    }
    // 左罫線は各行の開始セルのみ
    for(int i=stRow; i<=edRow;i++) {
      Cell cel = _sheet.getRow(i).getCell(stCell);
      CellStyle styleL = cel.getCellStyle();
      styleL.setBorderLeft(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT);
      cel.setCellStyle(styleL);
    }
    // 右罫線は各行の終了セルのみ
    for(int i=stRow; i<=edRow;i++) {
      Cell cel = _sheet.getRow(i).getCell(edCell);
      CellStyle styleR = cel.getCellStyle();
      styleR.setBorderRight(
        HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT);
      cel.setCellStyle(styleR);
    }
  }
  /** 処理の実行 */
  public void Run() {
    Cell wCell = null;
    // ワークブックの生成
    if (_mode.equals("2003")) {
      _workBook = new HSSFWorkbook();
    }
    else if (_mode.equals("2007")) {
      _workBook = new XSSFWorkbook();
    }
    else {
      System.out.println(
            "モードは2003か2007を指定します。");
      return;
    }
    // ワークシートの生成
    _sheet = _workBook.createSheet("POIライブラリ概要");
    if (_mode.equals("2003")) {
      _patr2003 = 
        ((HSSFSheet)_sheet).createDrawingPatriarch();
    }
    else {
      _patr2007 = 
        ((XSSFSheet)_sheet).createDrawingPatriarch();
    }
    // Rowの一括生成
    for (int i=0;i<25;i++) {
      Row row = _sheet.createRow(i);
      //cellの一括生成
      //styleも作成しておく
      for (int j=0;j<20;j++) {
        Cell cel = row.createCell((short)j);
        cel.setCellStyle(_workBook.createCellStyle());
      }
    }
    wCell = _sheet.getRow(22).getCell(3);
    wCell.setCellValue("ワークブック全体");
    setUserCellStyle(wCell,(short)36,false);
    // ワークブックコメント
    Comment cmtBook = getComment("Workbook");
    wCell.setCellComment(cmtBook);

    // 大きく罫線を引く
    drawLines(0,19,0,9);

    // 罫線を引く
    drawLines(8,8,0,9);

    // シートの文字とコメントの設定
    wCell = _sheet.getRow(3).getCell(3);
    wCell.setCellValue("ワークシート");
    setUserCellStyle(wCell,(short)24,false);
    Comment cmtSheet = getComment("Sheet");
    wCell.setCellComment(cmtSheet);

    // Rowの文字と罫線描画とコメントの設定
    _sheet.addMergedRegion(
        new org.apache.poi.ss.util.CellRangeAddress(
              8, 8, 0, 9));
    wCell = _sheet.getRow(8).getCell(0);
    wCell.setCellValue("行(Row)");
    setUserCellStyle(wCell,(short)12,true);
    Comment cmtRow = getComment("Row");
    wCell.setCellComment(cmtRow);

    // 罫線を引く
    drawLines(16,16,1,1);

    // Cellの文字列と罫線描画とコメントの設定
    wCell = _sheet.getRow(16).getCell(1);
    wCell.setCellValue("セル(Cell)");
    setUserCellStyle(wCell,(short)9,false);
    // Cellコメント
    Comment cmtCell = getComment("Cell");
    wCell.setCellComment(cmtCell);

    // 罫線を引く
    drawLines(16,17,5,7);

    // リージョンの文字列と罫線描画とコメントの設定
    _sheet.addMergedRegion(
        new org.apache.poi.ss.util.CellRangeAddress(
              16, 17, 5, 7));
    wCell = _sheet.getRow(16).getCell(5);
    wCell.setCellValue("マージドリージョン(MergedRegion)");
    setUserCellStyle(wCell,(short)9,true);
    // リージョンコメント
    Comment cmtRegion = getComment("マージされたCell");
    wCell.setCellComment(cmtRegion);

    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      if (_mode.equals("2003")) {
        out = new FileOutputStream(
          this.getClass().getName() + "_Book1.xls");
      }
      else {
        out = new FileOutputStream(
          this.getClass().getName() + "_Book1.xlsx");
      }
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
    if (args.length != 1) {
      System.out.println("エラー：モードを指定して下さい。");
      return;
    }
    else if ( !args[0].equals("2003") && 
              !args[0].equals("2007") ) {
      System.out.println(
        "エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    new MakeSummarySheetSS(args[0]).Run();

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
