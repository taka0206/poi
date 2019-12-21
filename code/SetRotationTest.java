import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * ������]�̃e�X�g
 */
public class SetRotationTest {

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
    // ��]�����\���prow(3�s��)�̐���
    Row row = sheet.createRow(2);
    // row������40�s�N�Z���ɁB
    row.setHeightInPoints((float)40);
    // �p�x�\���prow(4�s��)�̐���
    Row row2 = sheet.createRow(3);
    short angle = -90;  // -90������J�n
    for (int i=0; i<13; i++) {
      // ������]CellStyle����
      CellStyle style = workBook.createCellStyle();
      // �������񂹂�ݒ�B
      style.setAlignment(CellStyle.ALIGN_CENTER);
      // �c�����񂹂�ݒ�B
      style.setVerticalAlignment(
          CellStyle.VERTICAL_CENTER);
      // �p�x���w��
      style.setRotation(angle);
      // Cell�̐����ƕ�����ݒ�
      Cell cell = row.createCell(i);
      cell.setCellValue("POI");
      // CellStyle�̓K�p
      cell.setCellStyle(style);
      // �p�x�\���pCellStyle�̐���
      CellStyle style2 = workBook.createCellStyle();
      // ������������
      style2.setAlignment(CellStyle.ALIGN_CENTER);
      // Cell�̐���
      Cell cell2 = row2.createCell(i);
      // �p�x�������ݒ�
      cell2.setCellValue(angle + "��");
      // CellStyle�̓K�p
      cell2.setCellStyle(style2);
      angle += 15;
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
    new SetRotationTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
