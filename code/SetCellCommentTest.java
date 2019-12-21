import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;
/**
 * Cell�R�����ݒ�e�X�g
 */
public class SetCellCommentTest {

  // Patriarch�I�u�W�F�N�g 2003�̏ꍇ�̂�
  protected HSSFPatriarch _patr2003 = null;
  // Drawing�I�u�W�F�N�g 2007�̏ꍇ�̂�
  protected XSSFDrawing _patr2007 = null;

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
    // �R�����g�̐����Ɠ\��t��
    if (mode.equals("2003")) {
      _patr2003 = 
        ((HSSFSheet)sheet).createDrawingPatriarch();
      HSSFClientAnchor anchor = new HSSFClientAnchor(
        0, 0, 0, 0, (short)6, 4, (short)8, 9);
      anchor.setAnchorType(0); // Cell�ɕ����Ĉړ��E���T�C�Y
      // �R�����g�̐���
      HSSFComment cmt = _patr2003.createComment(anchor);
      // �R�����g�ɕ����ݒ�
      HSSFRichTextString rt = 
        new HSSFRichTextString("�R�����g");
      Font fnt = workBook.createFont();
      fnt.setFontName("�l�r �o�S�V�b�N");
      fnt.setFontHeightInPoints((short)14);
      fnt.setColor((short)HSSFColor.RED.index);
      fnt.setItalic(true);
      fnt.setBoldweight(Font.BOLDWEIGHT_BOLD);
      rt.applyFont(fnt);
      cmt.setString(rt); 
      cmt.setAuthor(new String("�ۉ� �F�i"));
      Cell cell = sheet.createRow(5).createCell(5);
      cell.setCellComment(cmt);
    }
    else {
      _patr2007 = 
        ((XSSFSheet)sheet).createDrawingPatriarch();
      XSSFClientAnchor anchor = new XSSFClientAnchor(
        0, 0, 0, 0, (short)6, 4, (short)8, 9);
      anchor.setAnchorType(0); // Cell�ɕ����Ĉړ��E���T�C�Y
      // �R�����g�̐���
      XSSFComment cmt = 
        _patr2007.createCellComment(anchor);
      // �R�����g�ɕ����ݒ�
      XSSFRichTextString rt = 
        new XSSFRichTextString("�R�����g");
      Font fnt = workBook.createFont();
      fnt.setFontName("�l�r �o�S�V�b�N");
      fnt.setFontHeightInPoints((short)14);
      fnt.setColor((short)HSSFColor.RED.index);
      fnt.setItalic(true);
      fnt.setBoldweight((short)10);
      rt.applyFont(fnt);
      cmt.setString(rt); 
      cmt.setAuthor(new String("�ۉ� �F�i"));
      Cell cell = sheet.createRow(5).createCell(5);
      cell.setCellComment(cmt);
    }
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
    new SetCellCommentTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
