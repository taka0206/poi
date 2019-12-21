import javax.jws.*;
import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 * POIラッパーWebサービス
 */
@WebService(targetNamespace="http://example.org")

public class CSWrapper {
  private Workbook _workBook = null;
  private Sheet _sheet = null;
  private Font _font = null;
  private CellStyle _style = null;
  private Row _row = null;
  private Cell _cell = null;
  /** コンストラクタ */
  public CSWrapper(){
  }
  /** 
   * ワークシート、行、セル生成
   *@param mode 動作モード 2003 or 2007
   *@param sName ワークシートの名前
   */ 
  public boolean createWorkSheetAndRowAndCell(
    @WebParam(name="mode") String mode,
    @WebParam(name="sName") String sName) {
    if ( !mode.equals("2003") && !mode.equals("2007")) {
      return false;
    }
    if (mode.equals("2003")) {
      _workBook = new HSSFWorkbook();
    }
    else {
      _workBook = new XSSFWorkbook();
    }
    _font = _workBook.createFont();
    _style = _workBook.createCellStyle();
    _sheet = _workBook.createSheet(sName);
    _row = _sheet.createRow(0);
    _cell = _row.createCell((short)0);
    return true;
  }
  /**
   * フォントの指定とスタイル設定
   *@param fontName フォント名
   *@param po 文字のポイント
   *@param col 文字色
   */
  public void setFontAndStyle(@WebParam(name="fontName") String fontName, 
                        @WebParam(name="po") int po, 
                        @WebParam(name="col") int col) {
    _font.setFontName(fontName);
    _font.setFontHeightInPoints((short)po);
    _font.setColor((short)col);
    _style.setFont(_font);
    _cell.setCellStyle(_style);
  }
  /**
   * セルに文字設定
   *@param sVal セルに設定したい文字
   */
  public void setCellValue(@WebParam(name="sVal") String sVal) {
    _cell.setCellValue(sVal);
  }
  /**
   * ブックファイル出力
   *@param fName 出力ファイル名
   */
  public void write(@WebParam(name="fname") String fName) {
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(fName);
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
  }
}
