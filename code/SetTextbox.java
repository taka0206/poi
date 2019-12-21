import java.io.*;
import org.apache.poi.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * �V�[�g�Ƀe�L�X�g�{�b�N�X��\��t����
 */ 
public class SetTextbox {

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
    // ���[�N�V�[�g����
    Sheet sheet = workBook.createSheet("Sheet1");
    // �e�L�X�g�{�b�N�X�����
    if (mode.equals("2003")) {
      _patr2003 = ((HSSFSheet)sheet).createDrawingPatriarch();
      HSSFClientAnchor anchor = new HSSFClientAnchor(
            0, 0, 0, 0,
            (short)1, 1, (short)8, 6);
      anchor.setAnchorType(0); // Cell�ɕ����Ĉړ��E���T�C�Y
      // Textbox�쐬
      HSSFTextbox box = _patr2003.createTextbox(anchor);
      // Textbox�ɏ����ݒ�
      // ������������
      box.setHorizontalAlignment(
          HSSFTextbox.HORIZONTAL_ALIGNMENT_CENTERED); 
      // ������������
      box.setVerticalAlignment(
          HSSFTextbox.VERTICAL_ALIGNMENT_CENTER);
      // �e�L�X�g�{�b�N�X�ɐݒ肷��HSSFRichTextString�C���X�^���X����
      HSSFRichTextString rst = 
              new HSSFRichTextString("Apache POI");
      // Font���w��
      Font fnt = workBook.createFont();
      fnt.setFontName("�l�r ����");
      fnt.setFontHeightInPoints((short)48);
      fnt.setColor((short)HSSFColor.BLUE.index);
      // Font��HSSFRichTextString�ɓK�p
      rst.applyFont(fnt);
      // Textbox��HSSFRichTextString���w��
      box.setString(rst);
    }
    else {
      // XSSF��POI API�������ł���B
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
    else if ( !args[0].equals("2003") ) {
      System.out.println("�G���[�F���[�h�͍��̂Ƃ���2003�̂ݎw�肵�ĉ������B");
      return;
    }
    // �����̎��s
    new SetTextbox().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
