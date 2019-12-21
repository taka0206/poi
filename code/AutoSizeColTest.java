import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * �񕝎����ݒ�̃e�X�g
 */
public class AutoSizeColTest {

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
    // Cell�ɐݒ肷�镶����e�[�u��
    String dat[] = {"1234567890",
                    "123456789012345",
                    "12345",
                    "123456789012",
                    "12345678901234567890",
                    "123"};
    // Row��Cell�̍쐬
    Row row1 = sheet.createRow(0);
    for( int i=0; i<6; i++) {
      Cell cell = row1.createCell(i);
      cell.setCellValue(dat[i]);
    }
    // 2�s�� �l�͋t���܂ɐݒ�
    Row row2 = sheet.createRow(1);
    for( int i=0; i<6; i++) {
      Cell cell = row2.createCell(i);
      cell.setCellValue(dat[5-i]);
    }
    // �񕝎����ݒ胂�[�h��
    for (int i=0; i<6; i++) {
      sheet.autoSizeColumn(i);
    }
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? this.getClass().getName() + "_Book1.xls" : 
                      this.getClass().getName() + "_Book1.xlsx");
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
    // �����̎��s
    new AutoSizeColTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
