import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
/**
 * JBasic向け型にうるさくないラッパークラス
 */
public class JBasWrapper {
  private Workbook _workBook = null;
  private Sheet _sheet = null;
  private Font _font = null;
  private CellStyle _style = null;
  private Row _row = null;
  private Cell _cell = null;
  private String _mode = "";
  /** コンストラクタ */
  public JBasWrapper(){
  }
  /** 
   * ワークシート、行、セル生成
   *@param sMode モード
   *@param sName ワークシートの名前
   */ 
  public void createWorkSheetAndRowAndCell(String sMode, String sName) {
    _mode = sMode;
    _workBook = _mode.equals("2003") ? new HSSFWorkbook() : new XSSFWorkbook();
    _font = _workBook.createFont();
    _style = _workBook.createCellStyle();
    _sheet = _workBook.createSheet(sName);
    _row = _sheet.createRow(0);
    _cell = _row.createCell(0);
  }
  /**
   * フォントの指定とスタイル設定
   *@param fontName フォント名
   *@param po 文字のポイント
   *@param col 文字色
   */
  public void setFontAndStyle(String fontName, int po, int col) {
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
  public void setCellValue(String sVal) {
    _cell.setCellValue(sVal);
  }
  /**
   * ブックファイル出力
   *@param fName 出力ファイル名
   */
  public void write(String fName) {
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(fName + (_mode.equals("2003") ? ".xls" : ".xlsx"));
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
