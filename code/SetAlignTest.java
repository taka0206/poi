import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * �����z�u�̃e�X�g
 */
public class SetAlignTest {

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
    //-----------------
    // �������̔z�u
    //-----------------
    // �z�u�`�����e�[�u��
    String[] captions = { 
      "������"
     ,"�I��͈͓��Œ���"
     ,"������"
     ,"����"
     ,"���������񂪒����Ƃ��܂�Ԃ��đS�̕\������w��"
     ,"���l��"
     ,"�E�l��"
    };
    // �z�u�`���e�[�u��
    short[] alignKinds = { CellStyle.ALIGN_CENTER
                      ,CellStyle.ALIGN_CENTER_SELECTION
                      ,CellStyle.ALIGN_FILL
                      ,CellStyle.ALIGN_GENERAL
                      ,CellStyle.ALIGN_JUSTIFY
                      ,CellStyle.ALIGN_LEFT
                      ,CellStyle.ALIGN_RIGHT
                     };

    // Style��7��ސ������ACell�ɐݒ�
    for (int i=0; i<7; i++) {
      // CellStyle����
      CellStyle style = workBook.createCellStyle();
      style.setAlignment(alignKinds[i]);
      // Row��Cell�𐶐����A������Style��ݒ�
      Cell cell = sheet.createRow(i + 1).createCell(1);
      cell.setCellValue(captions[i]);
      // Cell��CellSytle��K�p
      cell.setCellStyle(style);
    }
    // �񕝐ݒ�
    sheet.setColumnWidth(1, 5120);
    //-----------------
    // �c�����̔z�u
    //-----------------
    // �z�u�`�����e�[�u��
    String[] captionsV = { 
      "��l��"
     ,"������"
     ,"���l��"
     ,"����́A�܂�Ԃ��đS�̕\���Ɠ��l�̌��ʂ�����"
                     };
    // �z�u�`���e�[�u��
    short[] alignKindsV = { CellStyle.VERTICAL_TOP
                      ,CellStyle.VERTICAL_CENTER
                      ,CellStyle.VERTICAL_BOTTOM
                      ,CellStyle.VERTICAL_JUSTIFY
                     };

    // Style��5��ސ������ACell�ɐݒ�
    for (int i=0; i<4; i++) {
      // CellStyle����
      CellStyle style = workBook.createCellStyle();
      style.setVerticalAlignment(alignKindsV[i]);
      // Row��Cell�𐶐����A������Style��ݒ�
      Row row = sheet.createRow(i+9);
      Cell cell = row.createCell(1);
      cell.setCellValue(captionsV[i]);
      // Cell��CellSytle��K�p
      cell.setCellStyle(style);
      // �s�̍�����ݒ�-40�s�N�Z��
      row.setHeightInPoints((float)40);
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
    new SetAlignTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
