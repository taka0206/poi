import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 * Cellの値取得テスト
 */
class GetCellValueTest {
  /** 処理の実行
   * @param モード
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    // ワークブックを読み込む
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/GetCellValue.xls" : "./input/GetCellValue.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("ブックの読み込みに失敗しました。\n" + e.toString());
      return;
    }
    // 1番目のシートの取得
    Sheet sheet = workBook.getSheetAt(0);
    int j;
    for(int i=1; i<13; i++) {
      j = 3;
      Row row = sheet.getRow(i);
      Cell cell = row.getCell(1, row.CREATE_NULL_AS_BLANK);
      Cell cellDst = null;

      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      // getStringCellvalue
      try {
        String s = cell.getStringCellValue();
        if (s==null) {
          cellDst.setCellValue("null");
        }
        else if (s.equals("")) {
          cellDst.setCellValue("空文字");
        }
        else {
          cellDst.setCellValue("○");
        }
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("×");
      }
      // getRitchStringCellValue
      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      try {
        RichTextString rs = cell.getRichStringCellValue();
        cellDst.setCellValue("○");
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("×");
      }
      // getDateCellValue
      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      try {
        if(DateUtil.isCellDateFormatted(cell)) {
          Date dt = cell.getDateCellValue();
          cellDst.setCellValue("日付" + dt.toString());
        }
        else {
          cellDst.setCellValue("日付でない");
        }
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("×");
      }
      // getNumericCellValue
      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      try {
        double db = cell.getNumericCellValue();
        cellDst.setCellValue("○(" + db + ")");
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("×");
      }
      // getBooleanCellValue
      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      try {
        boolean b = cell.getBooleanCellValue();
        cellDst.setCellValue("○(" + b + ")" );
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("×");
      }
      // getCellFormula
      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      try {
        String cf = cell.getCellFormula();
        cellDst.setCellValue("○");
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("×");
      }
      // getHyperlink() 
      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      try {
        Hyperlink hl = cell.getHyperlink();
        if (hl != null) {
          cellDst.setCellValue("○");
        }
        else {
          cellDst.setCellValue("null");
        }
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("×");
      }

      // getErrorCellValue
      cellDst = row.getCell(j++, row.CREATE_NULL_AS_BLANK);
      try {
        byte bt = cell.getErrorCellValue();
        cellDst.setCellValue("○");
      }
      catch(Exception e) {
        System.out.println(e.toString());
        cellDst.setCellValue("×");
      }
    }
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? this.getClass().getName() + "_Book1.xls" : 
                      this.getClass().getName() + "_Book1.xlsx");
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
      System.out.println("エラー：モードを指定してください。");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new GetCellValueTest().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
