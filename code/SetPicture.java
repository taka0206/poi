import java.io.*;
import org.apache.poi.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;

/**
 * �V�[�g�ɉ摜��\��t����
 */ 
public class SetPicture {

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
    Workbook workBook = mode.equals("2003") ?
                  new HSSFWorkbook() : 
                  new XSSFWorkbook();
 
    // ���[�N�V�[�g����
    Sheet sheet = workBook.createSheet("Sheet1");
    // �摜�t�@�C����ǂݍ���
    byte bytes[];
    try {
      bytes =  IOUtils.toByteArray(
        new FileInputStream("./project-logo.jpg"));
    }
    catch (Exception e) {
      System.out.println("�摜�t�@�C���Ǎ��G���[" + 
                      e.toString());
      return;
    }
    int picIdx = workBook.addPicture(bytes, 
                Workbook.PICTURE_TYPE_JPEG);
    ClientAnchor anchor;
    // �摜�̓\��t��
    if (mode.equals("2003")) {
      _patr2003 = (
        (HSSFSheet)sheet).createDrawingPatriarch();
      anchor = new HSSFClientAnchor(40, 40, 980, 220, 
                (short)1, 1, (short)3, 12);
      // Cell�ɕ����Ĉړ��E���T�C�Y
      anchor.setAnchorType(0); 
      // �摜�̓\��t��
      _patr2003.createPicture(anchor, picIdx);
    }
    else {
      _patr2007 = (
        (XSSFSheet)sheet).createDrawingPatriarch();
      anchor = new XSSFClientAnchor(40, 40, 980, 220, 
                (short)1, 1, (short)3, 12);
      // Cell�ɕ����Ĉړ��E���T�C�Y
      anchor.setAnchorType(0); 
      // �摜�̓\��t��
      _patr2007.createPicture(anchor, picIdx);
    }

   // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? 
                      this.getClass().getName() + "_Book1.xls" : 
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
    else if ( !args[0].equals("2003") && 
              !args[0].equals("2007") ) {
      System.out.println(
        "�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    // �����̎��s
    new SetPicture().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
