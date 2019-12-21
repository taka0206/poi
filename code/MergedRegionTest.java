import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

/**
 * Cell�̌����e�X�g
 */
public class MergedRegionTest {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // �V�[�g�̐���
    Sheet sheet = workBook.createSheet();
    // �K���ɕ�����ݒ肷��
    sheet.createRow(0).createCell(0).setCellValue(
          "�c����");
    sheet.getRow(0).createCell(2).setCellValue(
          "������");
    sheet.createRow(2).createCell(2).setCellValue(
          "�c������");
    // Cell����������
    // �c����
    sheet.addMergedRegion(
        new CellRangeAddress(0, 6, 0, 0));
    // ������
    sheet.addMergedRegion(
        new CellRangeAddress(0, 0, 2, 5));
    // �c������
    sheet.addMergedRegion(
        new CellRangeAddress(2, 6, 2, 5));
    // ���[�N�u�b�N�����o��
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
    // �����̎��s
    new MergedRegionTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
