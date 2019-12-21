import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * �^�U�l�ݒ�̃e�X�g
 */
public class SetBooleanTest {

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
    // Row�̐���
    Row row = sheet.createRow(0);
    // Cell�𐶐����^�U�l��ݒ�
    for (int i=0; i<10; i++) {
      row.createCell(i).setCellValue((i%2)==1);
    }
    // ������Ƃ��Đ^�U�l���ǂ���ݒ�
    for (int i=0; i<10; i++) {
      row.createCell(i).setCellValue((i%2)==1 ? 
        "TRUE" : "FALSE");
      row.createCell(i).setCellType(Cell.CELL_TYPE_BOOLEAN);

    }
    // �^�U�l�Ƃ��Ĉ����邩�𔻒�
    for (int i=0; i<10; i++) {
      try {
        if (row.getCell(i).getBooleanCellValue() == true) {
          System.out.println("Cell[" + i + "] = �^"); 
        }
        else {
          System.out.println("Cell[" + i + "] = �U"); 
        }
      }
      catch (Exception e) {
        System.out.println("�^�U�l�Ƃ��Ď擾�ł��܂���ł����B"); 
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
    new SetBooleanTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
