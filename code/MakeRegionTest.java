import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.*;

public class MakeRegionTest {

  public static void main(String[] args) {
    // ワークブックの生成
    HSSFWorkbook workBook = new HSSFWorkbook();

    // ワークシート生成
    HSSFSheet sheet = workBook.createSheet("Sheet1");

		// 行とセルを一括で作る
		for (int i=0; i<5;i++) {
			HSSFRow row = sheet.createRow(i);
			for (int j=0;j<10;j++){
				row.createCell(j);
			}
		}
    // リージョン作成のテスト
		sheet.addMergedRegion(new CellRangeAddress(1,2,3,5));
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream("./Book1.xls");
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
}
