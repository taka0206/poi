import java.io.*;
import org.apache.poi.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * �V�[�g�ɂ��̑��}�`��\��t����(hssf�̂�)
 */ 
public class SetShapeOther {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // ���[�N�V�[�g����
    Sheet sheet = workBook.createSheet("�V�F�C�v");
    // �e��V�F�C�v�����
    if (mode.equals("2003")) {
      HSSFPatriarch _patr2003 = 
        ((HSSFSheet)sheet).createDrawingPatriarch();
      // COMBOBOX �� �g�p����Ӗ��Ȃ��F
      // �o��Workbook���J���Ƃ��Ƀ��b�Z�[�W
      // (���̃t�@�C�����J�����Ƃ����Ƃ��ɁAOffice �t�@�C�����؋@�\�ɂ���Ė�肪���o����܂����B
      // ���̃t�@�C�����J���̂̓Z�L�����e�B��댯�ł���\��������܂��B)
      // ���A�Ȃɂ����삵�Ȃ��B
      HSSFClientAnchor anchorRectCombo = 
        new HSSFClientAnchor(0, 0, 0, 0, 
                  (short)1, 1, (short)3, 4);
      // Cell�ɕ����Ĉړ��E���T�C�Y
      anchorRectCombo.setAnchorType(0); 
      HSSFSimpleShape rShapeCombo = 
        _patr2003.createSimpleShape(anchorRectCombo);
      rShapeCombo.setShapeType(
        HSSFSimpleShape.OBJECT_TYPE_COMBO_BOX);
/*
      // PICTURE ���@�g�p�s��:�������ݎ�ClassCastException
      HSSFClientAnchor anchorRectPic = new HSSFClientAnchor(0, 0, 0, 0, 
                                  (short)1, 1, (short)3, 4);
      anchorRectPic.setAnchorType(0); // Cell�ɕ����Ĉړ��E���T�C�Y
      HSSFSimpleShape rShapePic = _patr2003.createSimpleShape(anchorRectPic);
      rShapePic.setShapeType(HSSFSimpleShape.OBJECT_TYPE_PICTURE);
      // COMMENT �� �g�p�s�F�������ݎ�IllegalArgumentException
      HSSFClientAnchor anchorComm = new HSSFClientAnchor(0, 0, 0, 0, 
                                  (short)1, 1, (short)3, 4);
      anchorComm.setAnchorType(0); // Cell�ɕ����Ĉړ��E���T�C�Y
      HSSFSimpleShape rShapeComm = _patr2003.createSimpleShape(anchorComm);
      rShapeComm.setShapeType(HSSFSimpleShape.OBJECT_TYPE_COMMENT);
*/
    }
    else {
      // �����ł͏������Ȃ��B
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
      System.out.println("�G���[�F���[�h��2003�̂ݎw�肵�ĉ������B");
      return;
    }
    // �����̎��s
    new SetShapeOther().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
