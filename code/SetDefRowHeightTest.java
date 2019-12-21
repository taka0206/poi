import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

/**
 * �W���s�����ݒ�̃e�X�g
 */
public class SetDefRowHeightTest {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // �V�[�g�̐��� 
    Sheet sheet1 = workBook.createSheet();
    Sheet sheet2 = workBook.createSheet();
		// Row��10�s����
		for (int i=0; i<10; i++) {
			Row row1 = sheet1.createRow(i);
			row1.createCell(0).setCellValue(i);
			row1.setHeight((short)1000);
			Row row2 = sheet2.createRow(i);
			row2.createCell(0).setCellValue(i);
			row2.setHeight((short)2000);
		}
		// Sheet1�̕W���s�����ݒ�
//		sheet1.setDefaultRowHeight((short)400);
		// sheet2�̕W���s�����ݒ�
//		sheet2.setDefaultRowHeightInPoints((float)40.0);
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./Book1.xls" : 
                      "./Book1.xlsx");
      workBook.write(out);
    }catch(IOException e){
      System.out.println("�u�b�N�̏������݂Ɏ��s���܂����B\n" + e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println("�u�b�N�̏������݂Ɏ��s���܂����B\n" + e.toString());
      }
    }
    System.out.println("done!");
  }
  /** �G���g���[�|�C���g */
  public static void main(String[] args) {
    if (args.length != 1) {
      System.out.println("�G���[�F���[�h���w�肵�ĉ������B");
      return;
    }
    else if ( !args[0].equals("2003") && !args[0].equals("2007") ) {
      System.out.println("�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    new SetDefRowHeightTest().Run(args[0]);
  }
}
