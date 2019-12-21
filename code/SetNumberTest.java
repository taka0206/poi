import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * ���l�ݒ�̃e�X�g
 */
public class SetNumberTest {

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
    // Row��Cell�̐���
    Row row = sheet.createRow(0);
    // Cell�𐶐������l��ݒ�
    for (int i=0; i<10; i++) {
      row.createCell(i).setCellValue((double)i);
    }
    // Cell�𐶐���"����"��ݒ�
    Row row2 = sheet.createRow(1);
    for (int i=0; i<10; i++) {
      row2.createCell(i).setCellValue(
        Integer.toString(i));
    }
    // �l��ǂݏo��
    for (int i=0; i<10; i++) {
      try {
        double d = row2.getCell(i).getNumericCellValue();
        System.out.println("Cell(" + i + ")" + d);
      }
      catch( Exception e) {
        System.out.println(
          "���l�Ƃ��Ēl���擾�ł��܂���ł����B Cell(" +
           i + ")");
      }
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
    new SetNumberTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
