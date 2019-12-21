import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

/**
 * �V�[�g�̃A�N�e�B�u��Ԃ𔻒肷��e�X�g
 */
public class IsActiveTest {

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
    // �K���ȃV�[�g���A�N�e�B�u�ɂ���
    Random rand = new Random();
    workBook.setActiveSheet(rand.nextInt(5));
    // �A�N�e�B�u�V�[�g��T��
    if (mode.equals("2003")) {
      for (int i=0; i<5; i++) {
        HSSFSheet sheet = 
              (HSSFSheet)workBook.getSheetAt(i);
        if (sheet.isActive()) {
          System.out.println("�V�[�g" + i + 
                "���A�N�e�B�u�ɂȂ��Ă��܂��B");
          break;
        }
      }
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
    new IsActiveTest().Run(args[0]);

    System.out.print("���^�[���L�[�ŏI���c�c");
    try {
      int c = System.in.read();
    }
    catch (Exception e) {
    }

  }
}
