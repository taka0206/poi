import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * �p���}�V�[�g�쐬�N���X
 */
class MakeInheritSheet {
  HSSFWorkbook _workBook = null;
  HSSFSheet _sheet = null;
  HSSFPatriarch _patr = null;

  /**
   * �C���ςݕ�����擾
   * @param text �ݒ肷�镶����
   */
  protected HSSFRichTextString getContent(
                              String text) {
    HSSFRichTextString st = 
              new HSSFRichTextString(text);
    HSSFFont fnt = _workBook.createFont();
    fnt.setFontName("�l�r �o�S�V�b�N");
    fnt.setFontHeightInPoints((short)12);
    st.applyFont(fnt);
    return st;
  }
  /** �����̎��s */
  public void Run() {
    // ���[�N�u�b�N�̐���
    _workBook = new HSSFWorkbook();
    // ���[�N�V�[�g�̐���
    _sheet = _workBook.createSheet(
              "SS�C���^�[�t�F�[�X�p���}");
    _patr = _sheet.createDrawingPatriarch();
    // �e�L�X�g�{�b�N�X�̐���
    HSSFTextbox box1 = _patr.createTextbox(
          new HSSFClientAnchor(0, 0, 0, 0, 
              (short)3, 3, (short) 8, 6));
    box1.setString(getContent(
      "org.apache.poi.ss.usermodel.Workbook" +
      "\n�C���^�[�t�F�[�X"));
    box1.setVerticalAlignment(
      HSSFTextbox.HORIZONTAL_ALIGNMENT_CENTERED);
    box1.setHorizontalAlignment(
      HSSFTextbox.VERTICAL_ALIGNMENT_CENTER);

    HSSFTextbox box2 = _patr.createTextbox(
                new HSSFClientAnchor(0, 0, 0, 0, 
                (short)1, 10, (short) 5, 13));
    box2.setString(getContent(
      "org.apache.poi.hssf.usermodel." +
      "\nHSSFWorkbook�N���X"));
    box2.setVerticalAlignment(
      HSSFTextbox.HORIZONTAL_ALIGNMENT_CENTERED);
    box2.setHorizontalAlignment(
      HSSFTextbox.VERTICAL_ALIGNMENT_CENTER);

    HSSFTextbox box3 = _patr.createTextbox(
                new HSSFClientAnchor(0, 0, 0, 0, 
                (short)6, 10, (short) 10, 13));
    box3.setString(getContent(
      "org.apache.poi.xssf.usermodel.\n" +
      "XSSFWorkbook�N���X"));
    box3.setVerticalAlignment(
      HSSFTextbox.HORIZONTAL_ALIGNMENT_CENTERED);
    box3.setHorizontalAlignment(
      HSSFTextbox.VERTICAL_ALIGNMENT_CENTER);
    // ���C���̐���
    HSSFSimpleShape shape1 = _patr.createSimpleShape(
                new HSSFClientAnchor(0, 0, 0, 0,
                (short)5, 6, (short)3,10));
    shape1.setShapeType(HSSFSimpleShape.OBJECT_TYPE_LINE);
    shape1.setLineStyle(HSSFShape.LINESTYLE_LONGDASHGEL);
    HSSFSimpleShape shape2 = _patr.createSimpleShape(
                new HSSFClientAnchor(0, 0, 0, 0,
                (short)6, 6, (short)8,10));
    shape2.setShapeType(HSSFSimpleShape.OBJECT_TYPE_LINE);
    shape2.setLineStyle(HSSFShape.LINESTYLE_LONGDASHGEL);
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(
            this.getClass().getName() + "_Book1.xls");
      _workBook.write(out);
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
  /**
   * �G���g���[�|�C���g
   */
  public static void main(String args[]) {
    new MakeInheritSheet().Run();
    System.out.print("���^�[���L�[�ŏI���c�c");

    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
