import java.io.*;
import java.util.*;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*; 
import org.apache.poi.hssf.record.crypto.*;
/**
 * �`��I�u�W�F�N�g�̍폜�e�X�g
 */
class RemoveDrawObjectTest {
  /** �����̎��s
   * @param ���[�h
   */
  public void Run(String mode) {
    FileInputStream fis = null;
    Workbook workBook = null;
    try {
      fis = new FileInputStream( mode.equals("2003") ? "./input/RemoveDrawObject_in.xls" : "./input/RemoveDrawObject_in.xlsx");
      workBook = mode.equals("2003") ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
      fis.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + e.toString());
      return;
    }
    // �V�[�g�̎擾
    Sheet sheet = workBook.getSheetAt(0);
		if (mode.equals("2003")) {
			// �`�挳���߂̎擾
			//HSSFPatriarch _patr2003 = ((HSSFSheet)sheet).createDrawingPatriarch();
			// �e�L�X�g�{�b�N�X�̂ݎ�菜��
			//int ret = _patr2003.removeTextbox();
			//System.out.println("�폜�����e�L�X�g�{�b�N�X��" + ret + "�ł��B");
		}
    // ���[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream( mode.equals("2003") ? "./" + this.getClass().getName() + "_Book1.xls" : 
                      "./" + this.getClass().getName() + "_Book1.xlsx");
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
      System.out.println("�G���[�F���[�h���w�肵�Ă��������B");
      return;
    }
    else if ( !args[0].equals("2003")) {
      System.out.println("�G���[�F���[�h��2003���w�肵�ĉ������B");
      return;
    }
    // �����̎��s
    new RemoveDrawObjectTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
    
  }
}
