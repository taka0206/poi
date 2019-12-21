import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * �����t�H���g�ݒ�̃e�X�g
 */
public class SetFontTest {

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
    // �ݒ蕶���e�[�u��
    String[] captions = { "�����E��d����"
                      ,"��d����(��v)�E���������E�ΆE��t��"
                      ,"����"
                      ,"�����E���������E�Α�"
                      ,"�����E����(��v)�A���t��"
                     };
    // �����F�e�[�u��
    short[] colors = { (short)HSSFColor.GREEN.index
                      ,(short)HSSFColor.BLUE.index
                      ,(short)HSSFColor.RED.index
                      ,(short)HSSFColor.MAROON.index
                      ,(short)HSSFColor.VIOLET.index
                     };
    // �A���_�[���C����ʃe�[�u��
    byte[] ulines = { Font.U_DOUBLE
                     ,Font.U_DOUBLE_ACCOUNTING
                     ,Font.U_NONE
                     ,Font.U_SINGLE
                     ,Font.U_SINGLE_ACCOUNTING
                    };
    // ��t��/���t���e�[�u��
    short[] offset = { Font.SS_NONE 
                      ,Font.SS_SUPER
                      ,Font.SS_NONE
                      ,Font.SS_NONE
                      ,Font.SS_SUB
                     };
    // Style��Font��5��ސ������ACell�ɐݒ�
    for (int i=0; i<5; i++) {
      // 1.CellStyle�C���X�^���X����
      CellStyle style = workBook.createCellStyle();
      // 2.�t�H���g�C���X�^���X�𐶐��B
      Font fnt = workBook.createFont();
      // 3.�t�H���g�C���X�^���X�ɂ��܂��܂Ȑݒ���s���B
      // �t�H���g���
      fnt.setFontName("�l�r�@�S�V�b�N");
      // �|�C���g
      fnt.setFontHeightInPoints((short)(12+(i*2)));
      // �����F
      fnt.setColor(colors[i]);
      // �Α�
      fnt.setItalic(((i % 2) == 1) ? true : false);
      // �ʏ�܂��͑���
      fnt.setBoldweight(((i % 2) == 1) ? 
                  Font.BOLDWEIGHT_NORMAL : 
                  Font.BOLDWEIGHT_BOLD);
      // ����
      fnt.setUnderline(ulines[i]);
      // ��������
      fnt.setStrikeout(((i % 2) == 1) ? true : false);
      // ��t���܂��͉��t��
      fnt.setTypeOffset(offset[i]);
      // 4.CellStyle��Font��K�p
      style.setFont(fnt);
      // Row��Cell�𐶐����A������Style��ݒ�
      Cell cell = sheet.createRow(i + 1).createCell(1);
      cell.setCellValue(captions[i]);
      // 5.Cell��CellSytle��K�p
      cell.setCellStyle(style);
    }
    // �J�����������ݒ�
    sheet.autoSizeColumn(1);
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
    new SetFontTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
