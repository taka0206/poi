import java.io.*;
import org.apache.poi.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * シートにその他図形を貼り付ける(hssfのみ)
 */ 
public class SetShapeOther {

  /** 
   * 処理の実行
   * @param mode 動作モード
   */
  public void Run(String mode) {
    // ワークブックの生成
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // ワークシート生成
    Sheet sheet = workBook.createSheet("シェイプ");
    // 各種シェイプを作る
    if (mode.equals("2003")) {
      HSSFPatriarch _patr2003 = 
        ((HSSFSheet)sheet).createDrawingPatriarch();
      // COMBOBOX → 使用する意味なし：
      // 出力Workbookを開くときにメッセージ
      // (このファイルを開こうとしたときに、Office ファイル検証機能によって問題が検出されました。
      // このファイルを開くのはセキュリティ上危険である可能性があります。)
      // かつ、なにも動作しない。
      HSSFClientAnchor anchorRectCombo = 
        new HSSFClientAnchor(0, 0, 0, 0, 
                  (short)1, 1, (short)3, 4);
      // Cellに併せて移動・リサイズ
      anchorRectCombo.setAnchorType(0); 
      HSSFSimpleShape rShapeCombo = 
        _patr2003.createSimpleShape(anchorRectCombo);
      rShapeCombo.setShapeType(
        HSSFSimpleShape.OBJECT_TYPE_COMBO_BOX);
/*
      // PICTURE →　使用不可:書き込み時ClassCastException
      HSSFClientAnchor anchorRectPic = new HSSFClientAnchor(0, 0, 0, 0, 
                                  (short)1, 1, (short)3, 4);
      anchorRectPic.setAnchorType(0); // Cellに併せて移動・リサイズ
      HSSFSimpleShape rShapePic = _patr2003.createSimpleShape(anchorRectPic);
      rShapePic.setShapeType(HSSFSimpleShape.OBJECT_TYPE_PICTURE);
      // COMMENT → 使用不可：書き込み時IllegalArgumentException
      HSSFClientAnchor anchorComm = new HSSFClientAnchor(0, 0, 0, 0, 
                                  (short)1, 1, (short)3, 4);
      anchorComm.setAnchorType(0); // Cellに併せて移動・リサイズ
      HSSFSimpleShape rShapeComm = _patr2003.createSimpleShape(anchorComm);
      rShapeComm.setShapeType(HSSFSimpleShape.OBJECT_TYPE_COMMENT);
*/
    }
    else {
      // ここでは処理しない。
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
    else if ( !args[0].equals("2003") ) {
      System.out.println("エラー：モードは2003のみ指定して下さい。");
      return;
    }
    // 処理の実行
    new SetShapeOther().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
