import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * �����I�ɕ����̐F��ύX����T���v��
 */ 
public class SetPartColor {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
 
    // ���[�N�V�[�g����
    Sheet sheet = workBook.createSheet("Sheet1");
    // Row��1�s��������B
    Row row = sheet.createRow(0);
    // Cell���ЂƂ��
    Cell cell = row.createCell(0);
    // RichTextString�̃C���X�^���X�𐶐�����B
    RichTextString rt = mode.equals("2003") ? 
      new HSSFRichTextString("Hello POI World��") :
      new XSSFRichTextString("Hello POI World��");
    // 2��ނ̃t�H���g�𐶐�
    Font fnt1 = workBook.createFont();
    fnt1.setFontName("�l�r ����");
    fnt1.setFontHeightInPoints((short)48);
    fnt1.setColor((short)HSSFColor.AQUA.index);
    Font fnt2 = workBook.createFont();
    fnt2.setFontName("�l�r ����");
    fnt2.setFontHeightInPoints((short)48);
    fnt2.setColor((short)HSSFColor.RED.index);
    // �����S�̂�1�̃t�H���g��ݒ�
    rt.applyFont(0, rt.length(), fnt1);
    // POI�̕�����2�̃t�H���g��ݒ�
    rt.applyFont(6, 9, fnt2);
    // �Z���ɒl�ݒ�
    cell.setCellValue(rt);
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
    new SetPartColor().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
