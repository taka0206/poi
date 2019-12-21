import java.io.*;
import org.apache.poi.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * シートに図形を貼り付ける(hssfのみ)
 */ 
public class SetShape {

  // Patriarchオブジェクト シェイプ用
  protected HSSFPatriarch _patr2003 = null;
  // Patriarchオブジェクト ポリゴン用
  protected HSSFPatriarch _patr2003P = null;
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
      _patr2003 = ((HSSFSheet)sheet).createDrawingPatriarch();
      // 四角形-1
      HSSFClientAnchor anchorRect1 = 
                  new HSSFClientAnchor(0, 0, 0, 0, 
                      (short)1, 1, (short)3, 9);
      // Cellに併せて移動・リサイズ
      anchorRect1.setAnchorType(0); 
      HSSFSimpleShape rShape1 = 
          _patr2003.createSimpleShape(anchorRect1);
      rShape1.setShapeType(
          HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
      // 線の色を青にする。
      rShape1.setLineStyleColor(0x00, 0x00, 0xff);
      // 枠線をちょっと太め(2pt)に
      rShape1.setLineWidth(
          HSSFSimpleShape.LINEWIDTH_ONE_PT * 2);
      // 四角の中をシアンに塗り潰す
      rShape1.setFillColor(0x00, 0xff, 0xff);
      // 四角形-2
      HSSFClientAnchor anchorRect2 = 
                  new HSSFClientAnchor(0, 0, 0, 0, 
                      (short)4, 1, (short)6, 9);
      // Cellに併せて移動・リサイズ
      anchorRect2.setAnchorType(0); 
      HSSFSimpleShape rShape2 = 
          _patr2003.createSimpleShape(anchorRect2);
      rShape2.setShapeType(
          HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
      // 枠線をちょっと太め(2pt)に
      rShape2.setLineWidth(
          HSSFSimpleShape.LINEWIDTH_ONE_PT * 2);
      // 塗り潰しなしに設定
      rShape2.setNoFill(true);
      sheet.createRow(4).createCell(4).setCellValue(
          "透けてます");
      // 楕円の描画
      HSSFClientAnchor anchorOval = 
                new HSSFClientAnchor(0, 0, 0, 0, 
                (short)7, 1, (short)10, 9);
      // Cellに併せて移動・リサイズ
      anchorOval.setAnchorType(0);
      HSSFSimpleShape ovalShape = 
        _patr2003.createSimpleShape(anchorOval);
      ovalShape.setShapeType(
          HSSFSimpleShape.OBJECT_TYPE_OVAL);
      // 線の色を赤にする。
      ovalShape.setLineStyleColor(0xff, 0x00, 0x00);
      // 枠線を太め(5pt)に
      ovalShape.setLineWidth(
          HSSFSimpleShape.LINEWIDTH_ONE_PT * 5);
      // 四角の中をオレンジに塗り潰す
      ovalShape.setFillColor(0xff, 0xa5, 0x00);
      // 直線を描く
      // 形状名テーブルの定義
      String lineNames[] = {
        "実線"
       ,"破線"
       ,"点線"
       ,"一点鎖線"
       ,"二点差線"
       ,"粗い点線"
       ,"粗い破線"
       ,"粗い一点鎖線"
       ,"直線の長い一点鎖線"
       ,"直線の長い二点鎖線"
      };
      // 形状テーブルの定義
      int lineStyles[] = {  
        HSSFShape.LINESTYLE_SOLID
       ,HSSFShape.LINESTYLE_DASHSYS
       ,HSSFShape.LINESTYLE_DOTSYS
       ,HSSFShape.LINESTYLE_DASHDOTSYS
       ,HSSFShape.LINESTYLE_DASHDOTDOTSYS
       ,HSSFShape.LINESTYLE_DOTGEL
       ,HSSFShape.LINESTYLE_LONGDASHGEL
       ,HSSFShape.LINESTYLE_DASHDOTGEL
       ,HSSFShape.LINESTYLE_LONGDASHDOTGEL
       ,HSSFShape.LINESTYLE_LONGDASHDOTDOTGEL
      };
      int line = 11;
      for (int i=0; i<10; i++) {
        // 直線の場合は、Cellの真ん中あたりにくるようにマージンを取る。
        HSSFClientAnchor anchorLine = 
              new HSSFClientAnchor(0, 128, 0, 128, 
                (short)1, line, (short)4, line);
        // Cellに併せて移動・リサイズ
        anchorLine.setAnchorType(0); 
        // SimpleShapeの生成
        HSSFSimpleShape lShape = 
          _patr2003.createSimpleShape(anchorLine);
        // 直線を指定
        lShape.setShapeType(
          HSSFSimpleShape.OBJECT_TYPE_LINE);
        // 直線の形状を指定
        lShape.setLineStyle(lineStyles[i]);
        // 線の形状を書く
        Cell cell = sheet.createRow(line).createCell(4);
        cell.setCellValue(lineNames[i]);
        line++;
      }
    }
    else {
      // ここでは処理しない。
    }
    // ポリゴンを描く
    // シート2を作成
    Sheet sheet2 = workBook.createSheet("ポリゴン");
    if (mode.equals("2003")) {
      _patr2003P = 
        ((HSSFSheet)sheet2).createDrawingPatriarch();
      // ClientAnchorの生成
      HSSFClientAnchor anchorPol = 
          new HSSFClientAnchor(0, 0, 0, 0, 
            (short)1, 1, (short)6, 9);
      // Cellに併せて移動・リサイズ
      anchorPol.setAnchorType(0); 
      // Polygonインスタンスを生成
      HSSFPolygon pol = 
        _patr2003P.createPolygon(anchorPol);
      // 描画領域指定
      pol.setPolygonDrawArea(100, 100);
      // 各点のX、Y座標を設定する。
      pol.setPoints(
        // x座標の配列
        new int[]{10,20,30,40,50,60,70,80,90},
        // y座標の配列
        new int[]{10,20,30,20,80,10,50,90,40}); 
      // 線の色をマルーンにする
      pol.setLineStyleColor(0x80, 0x00, 0x00);
      // ミディアムパープルで塗り潰し
      pol.setFillColor(0x93,0x70, 0xdb);
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
    new SetShape().Run(args[0]);

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
