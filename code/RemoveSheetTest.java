import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * �V�[�g�폜�̃e�X�g
 */
public class RemoveSheetTest {

  /** 
   * �����̎��s
   * @param mode ���샂�[�h
   */
  public void Run(String mode) {
    // ���[�N�u�b�N�̐���
    Workbook workBook = mode.equals("2003") ? new HSSFWorkbook() : 
                                  new XSSFWorkbook();
    // �V�[�g�̐��� 
    for (int i=0; i<5; i++) {
      workBook.createSheet();
    }
    // �V�[�g�̍폜 - �O���� - NG
    // - NG(IllegalArgumentException)
    /*
    for (int i=0; i<5; i++) {
      workBook.removeSheetAt(i);
    }
    */
    // Sheet�C�e���[�^�[�ŏ��� 
    // - NG(ConcurrentModificationException)
    /*
    if (mode.equals("2007")) {
      for(XSSFSheet sheet : (XSSFWorkbook)workBook) {
        // �V�[�g���폜
        workBook.removeSheetAt(
          workBook.getSheetIndex(sheet));
      }
    }
    */
    // �V�[�g�̍폜(1���c��) - ��납�� - OK 
    for (int i=4; i>0; i--) {
      workBook.removeSheetAt(i);
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
    new RemoveSheetTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
