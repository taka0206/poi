import java.io.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.*;

public class MakeRegionTestXF {

  public static void main(String[] args) {
    // ���[�N�u�b�N�̐���
    XSSFWorkbook workBook = new XSSFWorkbook();

    // ���[�N�V�[�g����
    XSSFSheet sheet = workBook.createSheet("Sheet1");

		// �s�ƃZ�����ꊇ�ō��
		for (int i=0; i<5;i++) {
			XSSFRow row = sheet.createRow(i);
			for (int j=0;j<10;j++){
				row.createCell(j);
			}
		}
    // ���[�W�����쐬�̃e�X�g
		sheet.addMergedRegion(new CellRangeAddress(1,2,3,5));
    // ���[�N�u�b�N�����o��
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
