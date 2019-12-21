import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 * ���[�N�u�b�N�C���X�^���X�̕���
 */
class DuplicateBook {
  /** �����̎��s
   * @param ���[�h
   */
  public void Run(String mode) {
    FileInputStream fisA = null;
    FileInputStream fisB = null;
    Workbook workBookA = null;
    Workbook workBookB = null;
    // ���[�N�u�b�N��ǂݍ���
    try {
      fisA = new FileInputStream( 
          mode.equals("2003") ? 
            "./input/SampleLauncherORG.xls" : 
            "./input/SampleLauncherORG.xlsm");
      workBookA = mode.equals("2003") ? 
            new HSSFWorkbook(fisA) : 
            new XSSFWorkbook(fisA);
      fisA.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" + 
                        e.toString());
      return;
    }
    // ���O��ύX���ă��[�N�u�b�N�����o��
    FileOutputStream out = null;
    try{
      out = new FileOutputStream(
          mode.equals("2003") ? 
          this.getClass().getName() + "_Book1.xls" : 
          this.getClass().getName() + "_Book1.xlsm");
      workBookA.write(out);
    }catch(IOException e){
      System.out.println(e.toString());
    }finally{
      try {
        out.close();
      }catch(IOException e) {
        System.out.println(e.toString());
      }
    }
    // �������ꂽ���[�N�u�b�N��ǂݍ���
    try {
      fisB = new FileInputStream(
          mode.equals("2003") ? 
          this.getClass().getName() + "_Book1.xls" : 
          this.getClass().getName() + "_Book1.xlsm");
      workBookB = mode.equals("2003") ?
            new HSSFWorkbook(fisB) : 
            new XSSFWorkbook(fisB);
      fisB.close();
    }
    catch(Exception e) {
      System.out.println("�u�b�N�̓ǂݍ��݂Ɏ��s���܂����B\n" +
                        e.toString());
      return;
    }
    System.out.println("���[�N�u�b�N�̕������������܂����B");
  }
  /** �G���g���[�|�C���g */
  public static void main(String[] args) {
    if (args.length != 1) {
      System.out.println("�G���[�F���[�h���w�肵�Ă��������B");
      return;
    }
    else if ( !args[0].equals("2003") &&
              !args[0].equals("2007") ) {
      System.out.println(
      "�G���[�F���[�h��2003�܂���2007���w�肵�ĉ������B");
      return;
    }
    // �����̎��s
    new DuplicateBook().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }
  }
}
