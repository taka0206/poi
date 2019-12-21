import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 

/**
 * テンプレートを元にExcel帳票を作成するクラス
 */
public class PpoiTmplSS {
  // テンプレートブック名
  protected String _tmplBookName;
  // 出力ブック名
  protected String _outBookName;
  // 出力用データ配列
  protected ArrayList _itemAry;
  // ワークブックインターフェース
  protected Workbook _workBook = null;
  // 位置情報配列(ハッシュ)
  protected Hashtable<String,PosInfo> _posTbl = 
        new Hashtable<String,PosInfo>();
  // 関数情報配列
  protected ArrayList _funcTbl = new ArrayList();
  // 動作モード "2003" or "2007"
  protected String _mode;
  /**
   * インナークラス 位置情報
   */
  protected class PosInfo {
    public String _itemName;  // 項目名称
    public int _row;          // 行
    public int _col;          // 桁
    public String _type;      // タイプ
    public boolean _isArray;  // 配列フラグ
    public int _arrayMax;     // 配列最大値
    public int _incValue;     // 行増分
    /** コンストラクター
     *@param itemName 項目名称
     *@param row      行
     *@param col      桁
     *@param type     タイプ
     *@param isArray  配列？
     *@param arrayMax 配列最大値
     *@param incValue 行増分
     */
    public PosInfo( String itemName
                    ,int row
                    ,int col
                    ,String type
                    ,boolean isArray
                    ,int arrayMax
                    ,int incValue
                  ) {
      this._itemName = itemName;
      this._row = row;
      this._col = col;
      this._type = type;
      this._isArray = isArray;
      this._arrayMax = arrayMax;
      this._incValue = incValue;
    }
  }
  /**
   * インナークラス　関数情報
   */
  protected class FuncInfo {
    public String _funcName;  // 項目名称(使わない)
    public int _row;          // 行
    public int _col;          // 桁
    public boolean _isArray;  // 配列フラグ
    public int _arrayMax;     // 配列最大値
    public int _incValue;     // 行増分
    /**
     * コンストラクター
     *@param funcName 項目名称
     *@param row      行
     *@param col      桁
     *@param isArray  配列？
     *@param arrayMax 配列最大値
     *@param incValue 行増分
     */
    public FuncInfo( String funcName
                    ,int row
                    ,int col
                    ,boolean isArray
                    ,int arrayMax
                    ,int incValue
                  ) {
      this._funcName = funcName;
      this._row = row;
      this._col = col;
      this._isArray = isArray;
      this._arrayMax = arrayMax;
      this._incValue = incValue;
    }
  }
  /**
   * コンストラクター
   *@param mode         動作モード
   *@param tmplBookName テンプレートブックファイル名
   *@param outBookName  出力ブックファイル名
   *@param itemAry      出力用データ配列
   */
  public PpoiTmplSS(String mode, 
                    String tmplBookName,
                    String outBookName,
                    ArrayList itemAry){
    this._mode = mode;
    this._tmplBookName = tmplBookName;
    this._outBookName = outBookName;
    this._itemAry = itemAry;
  }
  /**
   * テンプレート読み込み処理
   *@return 読み込みに成功した場合はTrue 
   */
  protected boolean readTemplate() {
    System.out.println("テンプレートファイルを読み込みます。");
    // テンプレート読み込み
    try {
      if (_mode.equals("2003") ) {
        _workBook = new HSSFWorkbook(
            new FileInputStream(_tmplBookName));
      }
      else {
        _workBook = new XSSFWorkbook(
            new FileInputStream(_tmplBookName));
      }
    }
    catch( FileNotFoundException e ) {
      System.out.println(
        "テンプレートブックファイルが存在しません(" + 
        _tmplBookName + ")。");
      return false;
    }
    catch( IOException e ) {
      System.out.println(
        "テンプレートブックファイルの読み込みに失敗しました(" + 
        _tmplBookName + ")。" + e.toString());
      return false;
    }
    catch( Exception e ) {
      System.out.println(
        "テンプレート読み込みでエラーが発生しました。" + 
        e.toString());
      return false;
    }
    return true;
  }
  /** 位置情報テーブル構築 */
  protected boolean buildPosTable() {
    System.out.println("位置情報を構築します。");
    try {
      Sheet sheet = _workBook.getSheetAt(1);
      _posTbl.clear();
      for (int i=1;; i++) {
        // 項目数が変動する可能性があるので、
        // Null行が出現するまでループ
        Row row = sheet.getRow(i);
        if (row == null) break;
        PosInfo info = new PosInfo(
            row.getCell(0).getStringCellValue()
            ,(int)(row.getCell(1).getNumericCellValue())
            ,(int)(row.getCell(2).getNumericCellValue())
            ,row.getCell(3).getStringCellValue()
            ,row.getCell(4).getBooleanCellValue()
            ,(int)(row.getCell(5).getNumericCellValue())
            ,(int)(row.getCell(6).getNumericCellValue())
            );
        _posTbl.put(info._itemName,info);
      }
    }
    catch (Exception e) {
      System.out.println("位置情報の構築でエラーが発生しました。" + 
                    e.toString());
      return false;
    }
    return true;
  }
  /** 
   * 関数情報テーブル構築
   */
  protected boolean buildFuncTable() {
    System.out.println("関数情報を構築します。");
    Sheet sheet = _workBook.getSheetAt(1);
    _funcTbl.clear();
    try{
        // 項目数が変動する可能性があるので、
        // Null行が出現するまでループ
      for (int i=12;;i++) {
        Row row = sheet.getRow(i);
        if (row == null) break;
        FuncInfo info = new FuncInfo(
            row.getCell(0).getStringCellValue()
            ,(int)(row.getCell(1).getNumericCellValue())
            ,(int)(row.getCell(2).getNumericCellValue())
            ,row.getCell(3).getBooleanCellValue()
            ,(int)(row.getCell(4).getNumericCellValue())
            ,(int)(row.getCell(5).getNumericCellValue())
            );
        _funcTbl.add(info);
      }
    }
    catch (Exception e) {
      System.out.println("関数情報の構築でエラーが発生しました。" + 
              e.toString());
      return false;
    }
    return true;
  }
  /**
   * Excel帳票の作成
   */
  protected boolean makeExcelDocument() {
    System.out.println("Excel帳票の作成を行います。");
    Sheet sheet = _workBook.getSheetAt(0);
    try {
      for (int i=0; i<_itemAry.size(); i++) {
        String line = (String)_itemAry.get(i);
        String[] items = line.split("\t");
        PosInfo info = _posTbl.get(items[0]);
        if (info != null) {
          // 単一のときと配列のときで処理を分ける
          if (info._isArray == false) {
            // 単一のとき
            Cell cell = 
              sheet.getRow(info._row).getCell(info._col);
            if (setCellValue(
                  cell,info._type,items[1]) == false ) {
              return false;
            }
          }
          else {
            for (int j=1; j<items.length; j++) {
              // 繰り返し最大値よりデータが多い場合は棄てる。
              if (j>info._arrayMax) {
                break;
              }
              Cell cell = sheet.getRow(
                info._row + 
                (j*info._incValue) - 1).getCell(info._col);
              if (setCellValue(cell,
                    info._type,items[j]) == false ) {
                return false;
              }
            }
          }
        }
      }
      // ワークシートの内容を再集計
      if (resetFuncs() == false) return false;
    }
    catch (Exception e) {
      System.out.println("Excel帳票作成でエラーが発生しました。" + 
                e.toString());
      return false;
    }
    return true;
  }
  /**
   * Cellに値設定
   *@param cell 値を設定するセル
   *@param type タイプ
   *@param value 設定する値
   */
  protected boolean setCellValue(Cell cell, 
                        String type, String value) {
    try {
      if (type.equals("string")) {
        cell.setCellValue(value);
      }
      else if(type.equals("nummber")) {
        cell.setCellValue(Double.parseDouble(value));
      }
      else if(type.equals("date")) {
        cell.setCellValue(
          DateFormat.getDateInstance().parse(value));
      }
    }
    catch (Exception e) {
      System.out.println("Cellへの値設定でエラーが発生しました。" +
                e.toString());
      return false;
    }
    return true;
  }
  /**
   * 埋め込み関数の再設定処理
   * 2007モードの場合のみ必要
   */
  protected boolean resetFuncs() {
    System.out.println("埋め込み関数の再設定を行います。");
    try {
      for(int i=0; i<_funcTbl.size();i++) {
        FuncInfo tbl = (FuncInfo)_funcTbl.get(i);
        if (tbl._isArray == false) {
          Cell cell = 
            _workBook.getSheetAt(0).getRow(
                tbl._row).getCell(tbl._col);
          String func = cell.getCellFormula();
          cell.setCellFormula(func);
        }
        else {
          int rpos = tbl._row;
          for (int j=0;j<tbl._arrayMax;j++) {
            Cell cell = 
              _workBook.getSheetAt(0).getRow(rpos).getCell(tbl._col);
            String func = cell.getCellFormula();
            cell.setCellFormula(func);
            rpos += tbl._incValue;
          }
        }
      }
    }
    catch(Exception e) {
      System.out.println("組み込み関数の再設定でエラーが発生しました。" + 
                  e.toString());
    }
    return true;
  }
  /**
   * ExcelBook書き出し
   */
  protected boolean write() {
    System.out.println("ブックの書き出しを行います。");
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(_outBookName);
      _workBook.removeSheetAt(1);
      _workBook.write(out);
    }catch(IOException e){
      System.out.println(e.toString());
      return false;
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println(e.toString());
        return false;
      }
    }
    return true;
  }
  /**
   * 処理のコントロール
   */
  public boolean Run() {
    // テンプレートブック読み込み
    if (readTemplate() == false) {
      return false;
    }
    // 位置情報テーブル構築
    if (buildPosTable() == false) {
      return false;
    }
    // 関数情報テーブル構築 2007モードのときのみ
    //if (_mode.equals("2007")){
      if (buildFuncTable() == false) {
        return false;
      }
    //}
    // Excel帳票の作成
    if (makeExcelDocument() == false) {
      return false;
    }
    // ExcelBook書き出し
    if (write() == false) {
      return false;
    }
    return true;
  }
  /**
  * テスト用ルーチン
  *@param args [0]->動作モード、
  *            [1]->データファイル、
  *            [2]->テンプレートブック、
  *            [3]->出力ブック
  */
  public static void main(String[] args){

    if (args.length != 1) {
      System.out.println("パラメーターエラーです。");
    }
    else if (!args[0].equals("2003") && 
             !args[0].equals("2007")) {
      System.out.println(
        "動作モードは2003または2007を指定してください。");
    }
    else {
      try {
        ArrayList itemArray = new ArrayList();
        itemArray.clear();
        FileReader fr = new FileReader("./input/PpoiTmpl.dat");
        BufferedReader br = new BufferedReader(fr);
        String line;
        while ((line = br.readLine()) != null) {
          itemArray.add(line);
        }
        br.close();
        fr.close();
        if (new PpoiTmplSS(args[0],
                args[0].equals("2003") ? 
                "./input/nohinsyo_tmpl.xls" :
                "./input/nohinsyo_tmpl.xlsx",
                args[0].equals("2003") ? 
                "./PpoiTmplSS_Book1.xls" :
                "./PpoiTmplSS_Book1.xlsx",
                itemArray).Run()==true) {
          System.out.println("正常終了");
        }
        else {
          System.out.println("異常終了");
        }
      }
      catch (Exception e) {
        System.out.println("エラーが発生しました。" + 
                e.toString());
      } 
    }
  }
}
