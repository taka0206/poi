import java.io.*;
import org.apache.poi.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * シートに画像を貼り付ける
 */ 
public class SetPicture {

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
    Workbook workBook = mode.equals("2003") ?
                  new HSSFWorkbook() : 
                  new XSSFWorkbook();
 
    // ワークシート生成
    Sheet sheet = workBook.createSheet("Sheet1");
    // 画像ファイルを読み込む
    byte bytes[];
    try {
      bytes =  IOUtils.toByteArray(
        new FileInputStream("./project-logo.jpg"));
    }
    catch (Exception e) {
      System.out.println("画像ファイル読込エラー" + 
                      e.toString());
      return;
    }
    int picIdx = workBook.addPicture(bytes, 
                Workbook.PICTURE_TYPE_JPEG);
    ClientAnchor anchor;
    // 画像の貼り付け
    if (mode.equals("2003")) {
      _patr2003 = (
        (HSSFSheet)sheet).createDrawingPatriarch();
      anchor = new HSSFClientAnchor(40, 40, 980, 220, 
                (short)1, 1, (short)3, 12);
      // Cellに併せて移動・リサイズ
      anchor.setAnchorType(0); 
      // 画像の貼り付け
      _patr2003.createPicture(anchor, picIdx);
    }
    else {
      _patr2007 = (
        (XSSFSheet)sheet).createDrawingPatriarch();
      anchor = new XSSFClientAnchor(40, 40, 980, 220, 
                (short)1, 1, (short)3, 12);
      // Cellに併せて移動・リサイズ
      anchor.setAnchorType(0); 
      // 画像の貼り付け
      _patr2007.createPicture(anchor, picIdx);
    }

   // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? 
                      this.getClass().getName() + "_Book1.xls" : 
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
    else if ( !args[0].equals("2003") && 
              !args[0].equals("2007") ) {
      System.out.println(
        "エラー：モードは2003または2007を指定して下さい。");
      return;
    }
    // 処理の実行
    new SetPicture().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
