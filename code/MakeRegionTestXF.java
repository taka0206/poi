import java.io.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.*;

public class MakeRegionTestXF {

  public static void main(String[] args) {
    // ワークブックの生成
    XSSFWorkbook workBook = new XSSFWorkbook();

    // ワークシート生成
    XSSFSheet sheet = workBook.createSheet("Sheet1");

		// 行とセルを一括で作る
		for (int i=0; i<5;i++) {
			XSSFRow row = sheet.createRow(i);
			for (int j=0;j<10;j++){
				row.createCell(j);
			}
		}
    // リージョン作成のテスト
		sheet.addMergedRegion(new CellRangeAddress(1,2,3,5));
    // ワークブック書き出し
    FileOutputStream out = null;
    try{
      out = new FileOutputStream("./Book1.xlsx");
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
