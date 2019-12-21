import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * Java+POIで世界に挨拶するプログラム
 *
 */
public class HelloWorld{
  public static void main(String[] args){

    // ワークブックの生成
    HSSFWorkbook workBook = new HSSFWorkbook();
    // ワークシートの生成
    HSSFSheet sheet = 
        workBook.createSheet("HelloWorld");

    // Rowの生成
    HSSFRow row = sheet.createRow(0);

    // cellの生成
    HSSFCell cell = row.createCell((short)0);

    // cellスタイルの生成
    HSSFCellStyle st = workBook.createCellStyle();

    // フォントの生成
    HSSFFont fnt = workBook.createFont();
    fnt.setFontName("ＭＳ 明朝");
    fnt.setFontHeightInPoints((short)48);
    fnt.setColor((short)HSSFColor.AQUA.index);

    // cellスタイルにフォント設定
    st.setFont(fnt);

    // cellにスタイル設定
    cell.setCellStyle(st);

    // cellに値設定
    cell.setCellValue("Hello World♪");

    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(
              "HelloWorld_Book1.xls");
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

    System.out.print("リターンキーで終了……");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
